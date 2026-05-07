import os
from flask import Flask, render_template, redirect, url_for, request
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from dotenv import load_dotenv
from models import db, User, Lead, Atividade

# Carrega senhas e o DATABASE_URL do arquivo .env
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'chave-padrao-temporaria')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Inicia o Banco de Dados
db.init_app(app)

# Configura o sistema de Login
login_manager = LoginManager()
login_manager.login_view = 'login'
login_manager.init_app(app)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ==========================================
# ROTAS PÚBLICAS (A MÁQUINA DE VENDAS)
# ==========================================

@app.route('/')
def index():
    # O Flask procura o index.html automaticamente dentro da pasta "templates"
    return render_template('index.html')

@app.route('/politica-de-privacidade')
def privacidade():
    # Esta rota servirá o arquivo privacidade.html
    return render_template('privacidade.html')

@app.route('/termos-de-uso')
def termos():
    return render_template('termos.html')

@app.route('/agente-de-cura')
def agente_de_cura():
    return render_template('agente.html')

# ==========================================
# ROTAS DE ACESSO (LOGIN)
# ==========================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    # Aqui depois faremos a tela de colocar login e senha
    return "Em breve: Tela de Login do CRM"

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# ==========================================
# ROTAS DO CRM (PREFIXO /WORKBENCH)
# ==========================================

@app.route('/workbench/dashboard')
# @login_required  <-- Comentado temporariamente para você poder testar sem precisar logar
def dashboard():
    return "Bem-vindo ao CRM. Aqui você verá os cards dos Leads em breve!"

@app.route('/workbench/leads')
def leads_lista():
    return "Lista de todos os contatos capturados."


if __name__ == '__main__':
    app.run()