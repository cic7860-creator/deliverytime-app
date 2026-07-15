from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Center(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    address = db.Column(db.String(255), nullable=False)

# 💡 [신규] 문자 양식 템플릿 보관함
class SmsTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True) # 양식 이름 (예: A 양식)
    content = db.Column(db.Text, nullable=False)                  # 양식 내용

class Dispatch(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    delivery_date = db.Column(db.Date)
    center_name = db.Column(db.String(100))
    center_address = db.Column(db.String(255))
    vehicle_num = db.Column(db.String(50))
    driver_name = db.Column(db.String(100))
    store_code = db.Column(db.String(50))
    store_name = db.Column(db.String(100))
    store_address = db.Column(db.String(255))
    delivery_seq = db.Column(db.Integer)
    buffer_time = db.Column(db.Integer, default=10)
    store_x = db.Column(db.String(50))
    store_y = db.Column(db.String(50))
    is_departed = db.Column(db.Boolean, default=False)
    departure_time = db.Column(db.DateTime)
    estimated_arrival = db.Column(db.DateTime)
    center_depart_time = db.Column(db.DateTime)
    
    # 💡 [신규] 엑셀에서 바로 받아오는 추가 정보들
    driver_phone = db.Column(db.String(50))
    store_phone = db.Column(db.String(50))
    template_name = db.Column(db.String(100))
