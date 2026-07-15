from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

# 센터 테이블
class Center(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    address = db.Column(db.String(255), nullable=False)

# 💡 [신규] 기사 및 매장 연락처 마스터 테이블
class Contact(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    contact_type = db.Column(db.String(20)) # '기사' 또는 '매장'
    name = db.Column(db.String(100))        # 기사명 또는 매장명
    phone = db.Column(db.String(50))        # 연락처

# 배차 테이블
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
