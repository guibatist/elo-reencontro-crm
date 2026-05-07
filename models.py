from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class User(db.Model, UserMixin):
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default='CONSULTOR')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Lead(db.Model):
    __tablename__ = 'leads'
    
    id = db.Column(db.Integer, primary_key=True)
    nome_decisor = db.Column(db.String(100), nullable=False)
    telefone = db.Column(db.String(20), nullable=False)
    nome_paciente = db.Column(db.String(100))
    status = db.Column(db.String(50), default='Novo Lead')
    clinica_id = db.Column(db.Integer)
    consultor_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    ultima_atualizacao = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class Atividade(db.Model):
    __tablename__ = 'atividades'
    
    id = db.Column(db.Integer, primary_key=True)
    lead_id = db.Column(db.Integer, db.ForeignKey('leads.id', ondelete='CASCADE'))
    consultor_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    tipo_acao = db.Column(db.String(50)) # Ex: 'WhatsApp', 'Ligação'
    descricao = db.Column(db.Text)
    data_registro = db.Column(db.DateTime, default=datetime.utcnow)