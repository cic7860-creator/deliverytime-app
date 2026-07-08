from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Dispatch(db.Model):
    __tablename__ = 'dispatch'
    
    id = db.Column(db.Integer, primary_key=True)
    delivery_date = db.Column(db.Date, nullable=False)
    center_address = db.Column(db.String(200), nullable=True) 
    vehicle_num = db.Column(db.String(20), nullable=False)
    driver_name = db.Column(db.String(50), nullable=False)
    store_code = db.Column(db.String(50), nullable=True)
    store_name = db.Column(db.String(100), nullable=False)
    store_address = db.Column(db.String(200), nullable=True)
    delivery_seq = db.Column(db.Integer, default=1)
    buffer_time = db.Column(db.Integer, default=15)
    
    # 💡 [신규] 카카오내비 연동을 위한 GPS 좌표 저장 컬럼
    store_x = db.Column(db.String(50), nullable=True) # 경도
    store_y = db.Column(db.String(50), nullable=True) # 위도
    
    center_depart_time = db.Column(db.DateTime, nullable=True)
    estimated_arrival = db.Column(db.DateTime, nullable=True)
    departure_time = db.Column(db.DateTime, nullable=True)
    is_departed = db.Column(db.Boolean, default=False)

    def __repr__(self):
        return f"<Dispatch {self.driver_name} -> {self.store_name}>"