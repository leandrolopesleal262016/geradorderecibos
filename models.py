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

class ModeloRecibo(db.Model):
    __tablename__ = 'modelos_recibo'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    conteudo = db.Column(db.Text, nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
    data_atualizacao = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class Cliente(db.Model):
    __tablename__ = 'clientes'
    id = db.Column(db.Integer, primary_key=True)
    razao_social = db.Column(db.String(200), nullable=False)
    cpf_cnpj = db.Column(db.String(20), nullable=False, unique=True)
    tipo = db.Column(db.String(10), nullable=False)  # 'empresa' or 'pessoa'
