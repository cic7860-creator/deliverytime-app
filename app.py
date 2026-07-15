from flask import Flask, render_template, request, redirect, url_for, send_file, session
from models import db, Dispatch, Center, SmsTemplate
import pandas as pd
import io
from datetime import datetime, timedelta
import urllib.parse
import requests
import urllib3
import json
import concurrent.futures
from openpyxl.comments import Comment

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'jette_super_secret_admin_key' 

db.init_app(app)
with app.app_context():
    db.create_all()

KAKAO_API_KEY = 'f70047282a8b7f30cd02fd2cfc00f029'
kakao_session = requests.Session()
route_time_cache = {}

def get_coords(address):
    if not address or address.strip() == '': return None, None
    encoded_address = urllib.parse.quote(address)
    url = f"https://dapi.kakao.com/v2/local/search/address.json?query={encoded_address}"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    try:
        resp = kakao_session.get(url, headers=headers, verify=False, timeout=5).json()
        if resp.get('documents'): return resp['documents'][0]['x'], resp['documents'][0]['y']
    except Exception: pass
    return None, None

def get_driving_time(start_x, start_y, end_x, end_y):
    try:
        sx, sy, ex, ey = round(float(start_x), 4), round(float(start_y), 4), round(float(end_x), 4), round(float(end_y), 4)
        cache_key = f"{sx},{sy}_{ex},{ey}"
        if cache_key in route_time_cache: return route_time_cache[cache_key]
    except Exception: cache_key = None
    url = f"https://apis-navi.kakaomobility.com/v1/directions?origin={start_x},{start_y}&destination={end_x},{end_y}"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    try:
        resp = kakao_session.get(url, headers=headers, verify=False, timeout=2).json()
        if resp.get('routes'): 
            duration = resp['routes'][0]['summary']['duration']
            if cache_key: route_time_cache[cache_key] = duration
            return duration
    except Exception: pass
    return 25 * 60 

def update_etas_for_driver(driver_name):
    dispatches = Dispatch.query.filter_by(driver_name=driver_name).order_by(Dispatch.delivery_seq).all()
    if not dispatches: return
    base_depart_time = dispatches[0].center_depart_time
    if not base_depart_time: return 
        
    routes_to_fetch = []
    current_x, current_y = get_coords(dispatches[0].center_address if dispatches[0].center_address else dispatches[0].store_address)
    for d in dispatches:
        if d.is_departed and d.departure_time:
            current_x, current_y = d.store_x, d.store_y
        else:
            if current_x and current_y and d.store_x and d.store_y:
                routes_to_fetch.append((current_x, current_y, d.store_x, d.store_y))
            current_x, current_y = d.store_x, d.store_y
            
    duration_map = {}
    def fetch_route_time(route):
        sx, sy, ex, ey = route
        return route, get_driving_time(sx, sy, ex, ey)
        
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        results = executor.map(fetch_route_time, list(set(routes_to_fetch)))
        for route, duration in results: duration_map[route] = duration

    current_departure = base_depart_time
    current_x, current_y = get_coords(dispatches[0].center_address if dispatches[0].center_address else dispatches[0].store_address)
    for d in dispatches:
        if d.is_departed and d.departure_time:
            current_departure = d.departure_time
            current_x, current_y = d.store_x, d.store_y
        else:
            if current_x and current_y and d.store_x and d.store_y:
                duration_sec = duration_map.get((current_x, current_y, d.store_x, d.store_y), 25 * 60)
                arrival_time = current_departure + timedelta(seconds=duration_sec)
            else:
                arrival_time = current_departure + timedelta(minutes=25)
            d.estimated_arrival = arrival_time
            current_departure = arrival_time + timedelta(minutes=d.buffer_time)
            current_x, current_y = d.store_x, d.store_y

@app.route('/')
def home(): return redirect(url_for('admin_login'))

@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        if request.form.get('password') == 'a13579!!':
            session['is_admin'] = True
            return redirect(url_for('admin'))
        else:
            return "<script>alert('비밀번호가 틀렸습니다.'); history.back();</script>"
    return '''
        <div style="text-align:center; margin-top:150px; font-family:'Malgun Gothic', sans-serif;">
            <h2 style="color:#082c84;">JETTE 관리자 로그인</h2>
            <form method="post">
                <input type="password" name="password" placeholder="비밀번호 입력" required style="padding:10px; font-size:16px; border: 1px solid #ccc; border-radius: 4px;">
                <button type="submit" style="padding:10px 20px; font-size:16px; background:#082c84; color:white; border:none; border-radius: 4px; cursor:pointer;">접속</button>
            </form>
        </div>
    '''

@app.route('/admin_logout')
def admin_logout():
    session.pop('is_admin', None)
    return redirect(url_for('admin_login'))

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    
    if request.method == 'POST':
        excel_text = request.form.get('excel_text')
        if not excel_text or excel_text.strip() == '': return "입력된 데이터가 없습니다.", 400
        try:
            df = pd.read_csv(io.StringIO(excel_text), sep='\t')
            df.columns = df.columns.str.replace(' ', '')
            driver_seq_counter = {}
            address_cache = {}
            
            if '매장주소' in df.columns:
                unique_addresses = df['매장주소'].dropna().astype(str).str.strip().unique()
                unique_addresses = [addr for addr in unique_addresses if addr]
                def fetch_coord(addr): return addr, get_coords(addr)
                with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                    results = executor.map(fetch_coord, unique_addresses)
                    for addr, (sx, sy) in results: address_cache[addr] = (sx, sy)

            for index, row in df.iterrows():
                raw_date = str(row.get('배송일자', '')).strip()
                clean_date = raw_date.replace('년', '-').replace('월', '-').replace('일', '').replace(' ', '')
                if len(clean_date.split('-')) == 2: clean_date = f"{datetime.now().year}-{clean_date}"
                try: delivery_date = pd.to_datetime(clean_date).date()
                except: delivery_date = datetime.now().date()
                
                driver_name = str(row.get('기사명', '')).strip()
                if driver_name not in driver_seq_counter: driver_seq_counter[driver_name] = 1

                if '배송순서' in df.columns and pd.notna(row['배송순서']):
                    seq_value = int(row['배송순서'])
                    driver_seq_counter[driver_name] = seq_value + 1
                else:
                    seq_value = driver_seq_counter[driver_name]
                    driver_seq_counter[driver_name] += 1

                center_name_val = str(row.get('센터명', '')).strip()
                center_obj = Center.query.filter_by(name=center_name_val).first()
                center_addr_val = center_obj.address if center_obj else ''

                store_address_val = str(row.get('매장주소', '')).strip()
                buffer_time_val = int(row['상하차시간']) if '상하차시간' in df.columns and pd.notna(row['상하차시간']) else 10
                sx, sy = address_cache.get(store_address_val, (None, None))

                dispatch_entry = Dispatch(
                    delivery_date=delivery_date, 
                    center_name=center_name_val,
                    center_address=center_addr_val, 
                    vehicle_num=str(row.get('차량번호', '')).strip(),
                    driver_name=driver_name, 
                    store_code=str(row.get('매장코드', '')).strip(), 
                    store_name=str(row.get('매장명', '')).strip(),
                    store_address=store_address_val, 
                    delivery_seq=seq_value, 
                    buffer_time=buffer_time_val, 
                    store_x=sx, store_y=sy,
                    driver_phone=str(row.get('기사전화번호', '')).strip(),
                    store_phone=str(row.get('매장전화번호', '')).strip(),
                    template_name=str(row.get('템플릿양식', '')).strip()
                )
                db.session.add(dispatch_entry)
            db.session.commit()
            return redirect(url_for('admin'))
        except Exception as e:
            db.session.rollback()
            return f"오류 발생: {str(e)} <br><br><a href='/admin'>돌아가기</a>"
            
    all_data = Dispatch.query.order_by(Dispatch.delivery_date, Dispatch.driver_name, Dispatch.delivery_seq).all()
    centers = Center.query.all()
    return render_template('admin.html', dispatches=all_data, centers=centers)

@app.route('/admin/add_center', methods=['POST'])
def add_center():
    name = request.form.get('center_name').strip()
    address = request.form.get('center_address').strip()
    if name and address:
        if not Center.query.filter_by(name=name).first():
            db.session.add(Center(name=name, address=address))
            db.session.commit()
    return redirect(url_for('admin'))

@app.route('/admin/delete_center/<int:center_id>', methods=['POST'])
def delete_center(center_id):
    c = Center.query.get(center_id)
    if c:
        db.session.delete(c)
        db.session.commit()
    return redirect(url_for('admin'))

@app.route('/admin/delete_by_center', methods=['POST'])
def delete_by_center():
    center_name = request.form.get('center_name')
    if center_name:
        Dispatch.query.filter_by(center_name=center_name).delete()
        db.session.commit()
    return redirect(url_for('admin'))

@app.route('/admin/delete_all', methods=['POST'])
def delete_all():
    db.session.query(Dispatch).delete()
    db.session.commit()
    return redirect(url_for('admin'))

@app.route('/admin/delete/<int:dispatch_id>', methods=['POST'])
def delete_dispatch(dispatch_id):
    d = Dispatch.query.get(dispatch_id)
    if d:
        db.session.delete(d)
        db.session.commit()
    return redirect(url_for('admin'))

@app.route('/admin/update_buffer/<int:dispatch_id>', methods=['POST'])
def update_buffer(dispatch_id):
    dispatch = Dispatch.query.get(dispatch_id)
    if dispatch:
        new_time = request.form.get('buffer_time', type=int)
        if new_time is not None:
            dispatch.buffer_time = new_time
            db.session.commit()
            update_etas_for_driver(dispatch.driver_name) 
            db.session.commit()
    return redirect(url_for('admin'))

@app.route('/admin/update_driver/<int:dispatch_id>', methods=['POST'])
def update_driver(dispatch_id):
    dispatch = Dispatch.query.get(dispatch_id)
    if dispatch:
        new_driver = request.form.get('new_driver_name', '').strip()
        new_vehicle = request.form.get('new_vehicle_num', '').strip()
        
        if new_driver and new_vehicle:
            old_driver = dispatch.driver_name
            dispatch.driver_name = new_driver
            dispatch.vehicle_num = new_vehicle
            if new_driver != old_driver:
                max_seq_dispatch = Dispatch.query.filter_by(driver_name=new_driver).order_by(Dispatch.delivery_seq.desc()).first()
                if max_seq_dispatch: dispatch.delivery_seq = max_seq_dispatch.delivery_seq + 1
                else: dispatch.delivery_seq = 1
            db.session.commit()
            update_etas_for_driver(old_driver)
            if new_driver != old_driver: update_etas_for_driver(new_driver)
            db.session.commit()
    return redirect(url_for('admin'))

@app.route('/driver', methods=['GET'])
def driver():
    name = request.args.get('driver_name')
    dispatches = []
    route_chunks = [] 
    date_str = ""
    if name:
        dispatches = Dispatch.query.filter_by(driver_name=name).order_by(Dispatch.delivery_seq).all()
        if not dispatches:
            return f"<script>alert('{name} 기사님의 배차 내역이 존재하지 않습니다.'); window.location.href='/driver';</script>"
        display_date = datetime.now().date()
        if dispatches and dispatches[0].delivery_date: display_date = dispatches[0].delivery_date
        weekdays = ['월', '화', '수', '목', '금', '토', '일']
        date_str = display_date.strftime('%y%m%d') + f"({weekdays[display_date.weekday()]})"
        valid_dispatches = [d for d in dispatches if d.store_x and d.store_y and not d.is_departed]
        chunk_size = 5
        for i in range(0, len(valid_dispatches), chunk_size):
            chunk = valid_dispatches[i:i+chunk_size]
            dest_d = chunk[-1]
            ep_y, ep_x = float(dest_d.store_y), float(dest_d.store_x)
            kakaomap_app_url = f"kakaomap://route?ep={ep_y},{ep_x}&by=CAR"
            for idx, wp in enumerate(chunk[:-1]):
                vp_key = 'vp' if idx == 0 else f"vp{idx+1}"
                kakaomap_app_url += f"&{vp_key}={float(wp.store_y)},{float(wp.store_x)}"
            center_addr = dispatches[0].center_address if dispatches[0].center_address else dispatches[0].store_address
            origin_encoded = urllib.parse.quote(center_addr)
            dest_encoded = urllib.parse.quote(dest_d.store_address)
            google_map_url = f"https://www.google.com/maps/dir/?api=1&origin={origin_encoded}&destination={dest_encoded}"
            waypoint_addrs = [v.store_address for v in chunk[:-1]]
            if waypoint_addrs:
                wp_encoded = urllib.parse.quote("|".join(waypoint_addrs))
                google_map_url += f"&waypoints={wp_encoded}"
            route_chunks.append({
                'title': f"📱 코스 ({chunk[0].delivery_seq}~{chunk[-1].delivery_seq}번)",
                'url': kakaomap_app_url, 'pc_url': google_map_url 
            })
    return render_template('driver.html', dispatches=dispatches, driver_name=name, route_chunks=route_chunks, date_str=date_str)

@app.route('/depart_center', methods=['POST'])
def depart_center():
    driver_name = request.form.get('driver_name')
    manual_time_str = request.form.get('manual_time') 
    dispatches = Dispatch.query.filter_by(driver_name=driver_name).order_by(Dispatch.delivery_seq).all()
    if not dispatches: return "데이터 없음", 404
    depart_dt = datetime.now() 
    if manual_time_str:
        today = datetime.now().date()
        time_obj = datetime.strptime(manual_time_str, '%H:%M').time()
        depart_dt = datetime.combine(today, time_obj)
    for d in dispatches: d.center_depart_time = depart_dt
    db.session.commit()
    update_etas_for_driver(driver_name)
    db.session.commit()
    return redirect(url_for('driver', driver_name=driver_name))

@app.route('/cancel_depart', methods=['POST'])
def cancel_depart():
    driver_name = request.form.get('driver_name')
    for d in Dispatch.query.filter_by(driver_name=driver_name).all():
        d.center_depart_time = None
        d.estimated_arrival = None
    db.session.commit()
    return redirect(url_for('driver', driver_name=driver_name))

@app.route('/complete/<int:dispatch_id>', methods=['POST'])
def complete_delivery(dispatch_id):
    dispatch = Dispatch.query.get(dispatch_id)
    if dispatch:
        dispatch.is_departed = True
        dispatch.departure_time = datetime.now()
        db.session.commit()
        update_etas_for_driver(dispatch.driver_name) 
        db.session.commit()
        return redirect(url_for('driver', driver_name=dispatch.driver_name))
    return "데이터 없음", 404

@app.route('/update_seq/<int:dispatch_id>', methods=['POST'])
def update_seq(dispatch_id):
    dispatch = Dispatch.query.get(dispatch_id)
    if dispatch:
        new_seq = request.form.get('new_seq', type=int)
        if new_seq:
            completed_count = Dispatch.query.filter_by(driver_name=dispatch.driver_name, is_departed=True).count()
            if new_seq <= completed_count: new_seq = completed_count + 1
            dispatch.delivery_seq = new_seq
            db.session.commit()
            update_etas_for_driver(dispatch.driver_name)
            db.session.commit()
        return redirect(url_for('driver', driver_name=dispatch.driver_name))
    return "데이터 없음", 404

@app.route('/update_order', methods=['POST'])
def update_order():
    order_data = request.json
    driver_name = None
    if order_data:
        first_dispatch = Dispatch.query.get(int(order_data[0]))
        if first_dispatch:
            driver_name = first_dispatch.driver_name
            completed_count = Dispatch.query.filter_by(driver_name=driver_name, is_departed=True).count()
            for index, item_id in enumerate(order_data):
                dispatch = Dispatch.query.get(int(item_id))
                if dispatch and not dispatch.is_departed:
                    dispatch.delivery_seq = completed_count + index + 1
        db.session.commit()
        if driver_name:
            update_etas_for_driver(driver_name)
            db.session.commit()
    return {"status": "success"}

@app.route('/dashboard')
def dashboard():
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    unique_centers = [c.name for c in Center.query.all()]
    all_dispatches = Dispatch.query.order_by(Dispatch.driver_name, Dispatch.delivery_seq).all()
    stats = {}
    for d in all_dispatches:
        name = d.driver_name
        if name not in stats: stats[name] = {'total': 0, 'completed': 0, 'remaining': 0, 'vehicle': d.vehicle_num, 'details': []}
        stats[name]['total'] += 1
        if d.is_departed: stats[name]['completed'] += 1
        else: stats[name]['remaining'] += 1
        stats[name]['details'].append(d)
    for name, data in stats.items():
        data['progress'] = int((data['completed'] / data['total']) * 100) if data['total'] > 0 else 0
    vehicle_stats = {
        'total': len(stats),
        'completed': sum(1 for data in stats.values() if data['remaining'] == 0),
        'pending': len(stats) - sum(1 for data in stats.values() if data['remaining'] == 0)
    }
    return render_template('dashboard.html', stats=stats, vehicle_stats=vehicle_stats, unique_centers=unique_centers)

# 💡 [수정] 다운로드 시 센터 필터 적용
@app.route('/download_excel')
def download_excel():
    center_filter = request.args.get('center_name', '')
    query = Dispatch.query
    if center_filter:
        query = query.filter_by(center_name=center_filter)
    all_data = query.order_by(Dispatch.driver_name, Dispatch.delivery_seq).all()
    
    data_list = []
    for d in all_data:
        data_list.append({
            '센터명': d.center_name, '출발센터주소': d.center_address, '배송일자': d.delivery_date,
            '차량번호': d.vehicle_num, '기사명': d.driver_name, '매장코드': d.store_code,
            '매장명': d.store_name, '도착예정시간': d.estimated_arrival.strftime('%H:%M') if d.estimated_arrival else '-',
            '실제완료시간': d.departure_time.strftime('%H:%M') if d.departure_time else '미완료',
            '센터출발시간': d.center_depart_time.strftime('%H:%M') if d.center_depart_time else '미출발',
            '완료여부': '완료' if d.is_departed else '대기', '배송순서': d.delivery_seq,
            '상하차시간(분)': d.buffer_time, '매장주소': d.store_address
        })
    df = pd.DataFrame(data_list)
    today_str = datetime.now().strftime('%Y-%m-%d')
    if all_data and all_data[0].delivery_date: today_str = all_data[0].delivery_date.strftime('%Y-%m-%d')
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='배송시간')
        worksheet = writer.sheets['배송시간']
        for column_cells in worksheet.columns:
            max_len = 0
            for cell in column_cells:
                if cell.value:
                    val_str = str(cell.value)
                    val_len = sum(2 if ord(c) > 127 else 1.2 for c in val_str)
                    if val_len > max_len: max_len = val_len
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_len + 2
    output.seek(0)
    return send_file(output, download_name=f"{center_filter}_{today_str}_배송시간.xlsx" if center_filter else f"{today_str}_배송시간.xlsx", as_attachment=True)

@app.route('/download_template')
def download_template():
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    today_str = datetime.now().strftime('%Y-%m-%d')
    example_data = {
        '센터명': ['밀양센터', '오산센터'], '배송일자': [today_str, today_str],
        '차량번호': ['임시00임 0000', '임시00임 1111'], '기사명': ['홍길동', '이순신'],
        '매장코드': ['S001', 'S002'], '매장명': ['강남A', '강남B'],
        '매장주소': ['서울특별시 강남구 테헤란로 123', '서울특별시 강남구 테헤란로 456'],
        '배송순서': [1, 2], '상하차시간(분)': [10, 10], 
        '기사전화번호': ['010-1234-5678', '010-9876-5432'], '매장전화번호': ['02-111-2222', '031-333-4444'],
        '템플릿양식': ['A 양식', 'B 양식']
    }
    df = pd.DataFrame(example_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='업로드양식')
        worksheet = writer.sheets['업로드양식']
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_idx)
            if col_name == '배송순서': cell.comment = Comment('공란이어도 됩니다.', 'Admin')
            elif col_name == '상하차시간(분)': cell.comment = Comment('공란이어도 됩니다. (기본값: 10분)', 'Admin')
        for column_cells in worksheet.columns:
            max_len = 0
            for cell in column_cells:
                if cell.value:
                    val_str = str(cell.value)
                    val_len = sum(2 if ord(c) > 127 else 1.2 for c in val_str)
                    if val_len > max_len: max_len = val_len
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_len + 2
    output.seek(0)
    return send_file(output, download_name="JETTE_배차업로드양식.xlsx", as_attachment=True)

# ==========================================
# 💡 [신규] SMS 템플릿 및 발송 관리 라우트
# ==========================================
@app.route('/sms')
def sms_page():
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    templates = SmsTemplate.query.all()
    unique_centers = [c.name for c in Center.query.all()]
    
    # 뷰 렌더링용으로 모든 출발건 가져오기 (필터링은 화면에서 JS로)
    departed_dispatches = Dispatch.query.filter(Dispatch.center_depart_time != None).order_by(Dispatch.driver_name, Dispatch.delivery_seq).all()
    return render_template('sms.html', dispatches=departed_dispatches, templates=templates, unique_centers=unique_centers)

@app.route('/sms/add_template', methods=['POST'])
def add_template():
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    name = request.form.get('template_name').strip()
    content = request.form.get('template_content').strip()
    
    if name and content:
        existing = SmsTemplate.query.filter_by(name=name).first()
        if existing:
            existing.content = content
        else:
            db.session.add(SmsTemplate(name=name, content=content))
        db.session.commit()
    return redirect(url_for('sms_page'))

@app.route('/sms/delete_template/<int:template_id>', methods=['POST'])
def delete_template(template_id):
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    t = SmsTemplate.query.get(template_id)
    if t:
        db.session.delete(t)
        db.session.commit()
    return redirect(url_for('sms_page'))

@app.route('/download_sms_excel')
def download_sms_excel():
    if not session.get('is_admin'): return redirect(url_for('admin_login'))
    center_filter = request.args.get('center_name', '')
    
    query = Dispatch.query.filter(Dispatch.center_depart_time != None)
    if center_filter:
        query = query.filter_by(center_name=center_filter)
        
    departed_dispatches = query.order_by(Dispatch.driver_name, Dispatch.delivery_seq).all()
    templates_dict = {t.name: t.content for t in SmsTemplate.query.all()}
    
    data_list = []
    for d in departed_dispatches:
        eta_str = d.estimated_arrival.strftime('%H시 %M분') if d.estimated_arrival else "계산중"
        content_template = templates_dict.get(d.template_name, "안녕하세요 {매장명} 점주님!\n도착예정시간: {도착예정시간}\n기사명: {기사명}\n연락처: {기사전화번호}")
        
        # 💡 치환(Replace) 로직 적용
        sms_content = content_template.replace('{매장명}', d.store_name)\
                                      .replace('{도착예정시간}', eta_str)\
                                      .replace('{기사명}', d.driver_name)\
                                      .replace('{차량번호}', d.vehicle_num)\
                                      .replace('{기사전화번호}', d.driver_phone if d.driver_phone else "번호없음")
        
        store_phone = d.store_phone if d.store_phone else "번호없음"
        
        data_list.append({
            '수신인': d.store_name,
            '연락처': store_phone,
            '제목': f"[(주)제때] {d.store_name} 배송예정시간 안내",
            '내용': sms_content,
            '발신번호': '1668-3136'
        })
        
    df = pd.DataFrame(data_list)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='문자발송양식')
        worksheet = writer.sheets['문자발송양식']
        worksheet.column_dimensions['A'].width = 25
        worksheet.column_dimensions['B'].width = 20
        worksheet.column_dimensions['C'].width = 35
        worksheet.column_dimensions['D'].width = 50
        worksheet.column_dimensions['E'].width = 15
        
        from openpyxl.styles import Alignment
        for row in worksheet.iter_rows(min_col=4, max_col=4, min_row=2):
            for cell in row: cell.alignment = Alignment(wrap_text=True)
                
    output.seek(0)
    today_str = datetime.now().strftime('%Y%m%d')
    filename = f"{center_filter}_{today_str}_알림톡발송양식.xlsx" if center_filter else f"{today_str}_알림톡발송양식.xlsx"
    return send_file(output, download_name=filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
