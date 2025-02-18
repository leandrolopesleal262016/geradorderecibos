from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class ReceiptSequence(db.Model):
    __tablename__ = 'receipt_sequence'
    id = db.Column(db.Integer, primary_key=True)
    last_number = db.Column(db.Integer, nullable=False, default=0)

class ReciboGerado(db.Model):
    __tablename__ = 'recibos_gerados'
    id = db.Column(db.Integer, primary_key=True)
    numero_recibo = db.Column(db.String(10), unique=True, nullable=False)
    modelo_id = db.Column(db.Integer, nullable=False)
    cliente_nome = db.Column(db.String(200), nullable=False)
    valor = db.Column(db.Float, nullable=False)  # Mudando para Float
    data_geracao = db.Column(db.DateTime, default=datetime.utcnow)
    documento_blob = db.Column(db.LargeBinary)    
    documento_blob = db.Column(db.LargeBinary)