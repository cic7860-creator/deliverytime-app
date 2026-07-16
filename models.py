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

# 💡 [업데이트] 공지사항 모델 (이미지 필드 추가)
class Notice(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text, nullable=False)
    images_str = db.Column(db.Text)                               # 💡 신규: 이미지 파일명들을 구분자(|)로 묶어서 저장
    created_at = db.Column(db.DateTime, default=datetime.now)
    is_active = db.Column(db.Boolean, default=True)

    # 💡 신규: HTML 템플릿에서 이미지 목록을 바로 반복문 돌릴 수 있도록 리스트로 반환
    @property
    def image_list(self):
        if self.images_str:
            return self.images_str.split('|')
        return []

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
