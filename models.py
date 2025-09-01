from app import db
from datetime import datetime

class Shipment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    shopiny_number = db.Column(db.String(50), unique=True, nullable=False)
    receipt_number = db.Column(db.String(50))
    order_number = db.Column(db.String(50))
    delivery_date = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    from_governorate = db.Column(db.String(100), nullable=False)
    to_governorate = db.Column(db.String(100), nullable=False)
    carrier_company = db.Column(db.String(100))
    notes = db.Column(db.Text)
    items = db.relationship('ShipmentItem', backref='shipment', lazy=True)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

class ShipmentType(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    items = db.relationship('ShipmentItem', backref='shipment_type', lazy=True)

class Department(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    items = db.relationship('ShipmentItem', backref='department', lazy=True)

class ShipmentItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    shipment_id = db.Column(db.Integer, db.ForeignKey('shipment.id'), nullable=False)
    shipment_type_id = db.Column(db.Integer, db.ForeignKey('shipment_type.id'), nullable=False)
    department_id = db.Column(db.Integer, db.ForeignKey('department.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    cost = db.Column(db.Float, nullable=False)
    boxes_count = db.Column(db.Integer, nullable=False)
    total = db.Column(db.Float, nullable=False)
    notes = db.Column(db.Text)