from flask import Flask, render_template, request, redirect, url_for, send_file
from models import db, Dispatch
import pandas as pd
import io
from datetime import datetime, timedelta
import urllib.parse
import requests
import urllib3
import json

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db.init_app(app)

with app.app_context():
    db.create_all()

# ==========================================
# 🔑 카카오 API 설정 영역
# ==========================================
KAKAO_API_KEY = '여기에_복사한_REST_API_키를_붙여넣으세요'

def get_coords(address):
    if not address or address.strip() == '':
        return None, None
    encoded_address = urllib.parse.quote(address)
    url = f"https://dapi.kakao.com/v2/local/search/address.json?query={encoded_address}"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    try:
        resp = requests.get(url, headers=headers, verify=False).json()
        if resp.get('documents'):
            return resp['documents'][0]['x'], resp['documents'][0]['y']
    except Exception:
        pass
    return None, None

def get_driving_time(start_x, start_y, end_x, end_y):
    url = f"https://apis-navi.kakaomobility.com/v1/directions?origin={start_x},{start_y}&destination={end_x},{end_y}"
    headers = {"Authorization": f"KakaoAK {KAKAO_API_KEY}"}
    try:
        resp = requests.get(url, headers=headers, verify=False).json()
        if resp.get('routes'):
            return resp['routes'][0]['summary']['duration']
    except Exception:
        pass
    return 40 * 60 

@app.route('/')
def home():
    return "<h1>배차 관리 시스템 서버 가동 중!</h1><p><a href='/admin'>1. 관리자 페이지 (데이터 입력)</a></p><p><a href='/driver'>2. 기사님 페이지 (배차 조회)</a></p><p><a href='/dashboard'>3. 실시간 현황판 보기</a></p>"

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST':
        excel_text = request.form.get('excel_text')
        if not excel_text or excel_text.strip() == '':
            return "입력된 데이터가 없습니다.", 400

        try:
            df = pd.read_csv(io.StringIO(excel_text), sep='\t')
            df.columns = df.columns.str.replace(' ', '')
            driver_seq_counter = {}

            for index, row in df.iterrows():
                raw_date = str(row['배송일자']).strip()
                clean_date = raw_date.replace('년', '-').replace('월', '-').replace('일', '').replace(' ', '')
                if len(clean_date.split('-')) == 2:
                    clean_date = f"{datetime.now().year}-{clean_date}"
                delivery_date = pd.to_datetime(clean_date).date()
                
                driver_name = str(row['기사명']).strip()
                
                if driver_name not in driver_seq_counter:
                    driver_seq_counter[driver_name] = 1

                if '배송순서' in df.columns and pd.notna(row['배송순서']):
                    seq_value = int(row['배송순서'])
                    driver_seq_counter[driver_name] = seq_value + 1
                else:
                    seq_value = driver_seq_counter[driver_name]
                    driver_seq_counter[driver_name] += 1

                center_addr_val = str(row['출발센터주소']).strip() if '출발센터주소' in df.columns and pd.notna(row['출발센터주소']) else ''
                store_code_val = str(row['매장코드']).strip() if '매장코드' in df.columns and pd.notna(row['매장코드']) else ''
                store_address_val = str(row['매장주소']).strip() if '매장주소' in df.columns and pd.notna(row['매장주소']) else ''
                buffer_time_val = int(row['상하차시간']) if '상하차시간' in df.columns and pd.notna(row['상하차시간']) else 15

                sx, sy = get_coords(store_address_val)

                dispatch_entry = Dispatch(
                    delivery_date=delivery_date,
                    center_address=center_addr_val,
                    vehicle_num=str(row['차량번호']).strip(),
                    driver_name=driver_name,
                    store_code=store_code_val,
                    store_name=str(row['매장명']).strip(),
                    store_address=store_address_val,
                    delivery_seq=seq_value,
                    buffer_time=buffer_time_val,
                    store_x=sx,
                    store_y=sy
                )
                db.session.add(dispatch_entry)
            
            db.session.commit()
            return redirect(url_for('admin'))
        except Exception as e:
            db.session.rollback()
            return f"오류 발생: {str(e)} <br><br><a href='/admin'>돌아가기</a>"
            
    all_data = Dispatch.query.order_by(Dispatch.delivery_date, Dispatch.driver_name, Dispatch.delivery_seq).all()
    return render_template('admin.html', dispatches=all_data)

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

@app.route('/driver', methods=['GET'])
def driver():
    name = request.args.get('driver_name')
    dispatches = []
    route_chunks = [] 
    
    if name:
        dispatches = Dispatch.query.filter_by(driver_name=name).order_by(Dispatch.delivery_seq).all()
        valid_dispatches = [d for d in dispatches if d.store_x and d.store_y and not d.is_departed]
        
        chunk_size = 5
        for i in range(0, len(valid_dispatches), chunk_size):
            chunk = valid_dispatches[i:i+chunk_size]
            dest_d = chunk[-1]
            
            via_list = []
            for v in chunk[:-1]:
                via_list.append({
                    "name": v.store_name,
                    "x": float(v.store_x),
                    "y": float(v.store_y)
                })
                
            pars = {
                "destination": {
                    "name": dest_d.store_name,
                    "x": float(dest_d.store_x),
                    "y": float(dest_d.store_y)
                },
                "viaList": via_list,
                "coordType": "wgs84"
            }
            
            pars_encoded = urllib.parse.quote(json.dumps(pars, ensure_ascii=False))
            kakaonavi_app_url = f"kakaonavi://shareRoute?pars={pars_encoded}"
            
            center_addr = dispatches[0].center_address if dispatches[0].center_address else dispatches[0].store_address
            origin_encoded = urllib.parse.quote(center_addr)
            dest_encoded = urllib.parse.quote(dest_d.store_address)
            pc_web_url = f"https://map.kakao.com/?sName={origin_encoded}&eName={dest_encoded}"
            
            start_num = chunk[0].delivery_seq
            end_num = chunk[-1].delivery_seq
            
            route_chunks.append({
                'title': f"📱 코스 ({start_num}~{end_num}번)",
                'url': kakaonavi_app_url,
                'pc_url': pc_web_url 
            })

    return render_template('driver.html', dispatches=dispatches, driver_name=name, route_chunks=route_chunks)

@app.route('/depart_center', methods=['POST'])
def depart_center():
    driver_name = request.form.get('driver_name')
    manual_time_str = request.form.get('manual_time') 
    
    dispatches = Dispatch.query.filter_by(driver_name=driver_name).order_by(Dispatch.delivery_seq).all()
    if not dispatches:
        return "데이터 없음", 404
        
    depart_dt = datetime.now() 
    if manual_time_str:
        today = datetime.now().date()
        time_obj = datetime.strptime(manual_time_str, '%H:%M').time()
        depart_dt = datetime.combine(today, time_obj)
        
    current_time = depart_dt
    center_addr = dispatches[0].center_address if dispatches[0].center_address else dispatches[0].store_address
    current_x, current_y = get_coords(center_addr)

    for d in dispatches:
        d.center_depart_time = depart_dt
        if current_x and current_y and d.store_x and d.store_y:
            duration_sec = get_driving_time(current_x, current_y, d.store_x, d.store_y)
            current_time = current_time + timedelta(seconds=duration_sec) + timedelta(minutes=d.buffer_time)
        else:
            current_time = current_time + timedelta(minutes=d.buffer_time + 25)
            
        d.estimated_arrival = current_time
        current_x, current_y = d.store_x, d.store_y
        
    db.session.commit()
    return redirect(url_for('driver', driver_name=driver_name))

@app.route('/complete/<int:dispatch_id>', methods=['POST'])
def complete_delivery(dispatch_id):
    dispatch = Dispatch.query.get(dispatch_id)
    if dispatch:
        dispatch.is_departed = True
        dispatch.departure_time = datetime.now()
        db.session.commit()
        return redirect(url_for('driver', driver_name=dispatch.driver_name))
    return "데이터 없음", 404

@app.route('/update_seq/<int:dispatch_id>', methods=['POST'])
def update_seq(dispatch_id):
    dispatch = Dispatch.query.get(dispatch_id)
    if dispatch:
        new_seq = request.form.get('new_seq', type=int)
        if new_seq:
            dispatch.delivery_seq = new_seq
            db.session.commit()
        return redirect(url_for('driver', driver_name=dispatch.driver_name))
    return "데이터 없음", 404

# 💡 [신규] 손가락 드래그 앤 드롭 순서 변경 반영 API
@app.route('/update_order', methods=['POST'])
def update_order():
    order_data = request.json
    if order_data:
        for index, item_id in enumerate(order_data):
            dispatch = Dispatch.query.get(int(item_id))
            if dispatch:
                # 드래그된 순서대로 1번부터 새롭게 번호를 매깁니다.
                dispatch.delivery_seq = index + 1
        db.session.commit()
    return {"status": "success"}

@app.route('/dashboard')
def dashboard():
    all_dispatches = Dispatch.query.order_by(Dispatch.driver_name, Dispatch.delivery_seq).all()
    stats = {}
    for d in all_dispatches:
        name = d.driver_name
        if name not in stats:
            stats[name] = {'total': 0, 'completed': 0, 'remaining': 0, 'vehicle': d.vehicle_num, 'details': []}
        
        stats[name]['total'] += 1
        if d.is_departed:
            stats[name]['completed'] += 1
        else:
            stats[name]['remaining'] += 1
            
        stats[name]['details'].append(d)
            
    for name, data in stats.items():
        if data['total'] > 0:
            data['progress'] = int((data['completed'] / data['total']) * 100)
        else:
            data['progress'] = 0
            
    return render_template('dashboard.html', stats=stats)

@app.route('/download_excel')
def download_excel():
    all_data = Dispatch.query.order_by(Dispatch.driver_name, Dispatch.delivery_seq).all()
    data_list = []
    for d in all_data:
        data_list.append({
            '배송일자': d.delivery_date,
            '기사명': d.driver_name,
            '차량번호': d.vehicle_num,
            '출발센터주소': d.center_address,
            '배송순서': d.delivery_seq,
            '매장코드': d.store_code,
            '매장명': d.store_name,
            '센터출발시간': d.center_depart_time.strftime('%H:%M') if d.center_depart_time else '미출발',
            '도착예정시간': d.estimated_arrival.strftime('%H:%M') if d.estimated_arrival else '-',
            '실제완료시간': d.departure_time.strftime('%H:%M') if d.departure_time else '미완료',
            '완료여부': '완료' if d.is_departed else '대기',
            '상하차시간(분)': d.buffer_time 
        })
    df = pd.DataFrame(data_list)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='배송시간')
    output.seek(0)
    return send_file(output, download_name="배송시간.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
