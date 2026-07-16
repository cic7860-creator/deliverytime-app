from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class Center(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    address = db.Column(db.String(255), nullable=False)

class SmsTemplate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    subject = db.Column(db.String(255))
    sender_phone = db.Column(db.String(50))
    content = db.Column(db.Text, nullable=False)

class Notice(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text, nullable=False)
    target_drivers = db.Column(db.String(500), default='') # 💡 신규: 특정 기사님 지정 (쉼표로 구분)
    created_at = db.Column(db.DateTime, default=datetime.now)
    is_active = db.Column(db.Boolean, default=True)

# 💡 신규: 배송 완료 축하 팝업 설정
class CompletionSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    content = db.Column(db.Text) # 복사/붙여넣기 한 이미지 데이터 저장

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
    driver_phone = db.Column(db.String(50))
    store_phone = db.Column(db.String(50))
    template_name = db.Column(db.String(100))
