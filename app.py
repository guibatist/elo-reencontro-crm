import os
from flask import Flask, render_template, redirect, url_for, request, flash, make_response, send_from_directory
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import check_password_hash, generate_password_hash
from dotenv import load_dotenv
from datetime import date, datetime, timedelta
import calendar
from sqlalchemy import extract, func
import io
import uuid
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from werkzeug.utils import secure_filename
import random
import string
from flask_mail import Mail, Message

load_dotenv()

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'elo-secret-123')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# No bloco de configurações de e-mail do seu app.py
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USER')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASS')
app.config['MAIL_DEFAULT_SENDER'] = os.getenv('MAIL_USER') 
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 
mail = Mail(app)

# Configuração de Upload (Crie a pasta 'uploads' no seu projeto)
UPLOAD_FOLDER = 'static/uploads/chamados'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Configuração (Adicione no topo do app.py)
app.config['UPLOAD_FOLDER'] = 'uploads/anexos'
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # 5MB
ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg'}

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login_crm'

# ==========================================
# 1. MODELOS DE BANCO DE DADOS
# ==========================================

class Usuario(UserMixin, db.Model):
    __tablename__ = 'usuarios'
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(50), nullable=False)
    sobrenome = db.Column(db.String(50), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    senha = db.Column(db.String(255), nullable=False)
    # A user_key agora é o USERNAME de login (ex: joao.almeida)
    user_key = db.Column(db.String(100), unique=True, nullable=False)
    role = db.Column(db.String(20), default='consultor')
    is_active = db.Column(db.Boolean, default=True)
    data_criacao = db.Column(db.DateTime, default=db.func.current_timestamp())

    precisa_mudar_senha = db.Column(db.Boolean, default=True) 
    codigo_verificacao = db.Column(db.String(6))
    data_criacao = db.Column(db.DateTime, default=db.func.current_timestamp())

    def gerar_username(self):
        # Transforma "João" e "Almeida" em "joao.almeida"
        import unicodedata
        import re
        
        base = f"{self.nome}.{self.sobrenome}".lower()
        # Remove acentos e caracteres especiais
        nfkd = unicodedata.normalize('NFKD', base)
        username = "".join([c for c in nfkd if not unicodedata.combining(c)])
        return re.sub(r'[^a-z0-9.]', '', username)

class Decisor(db.Model):
    __tablename__ = 'decisores'
    id = db.Column(db.Integer, primary_key=True)
    nome_completo = db.Column(db.String(100), nullable=False)
    whatsapp = db.Column(db.String(20), nullable=False)
    parentesco = db.Column(db.String(50)) # <-- O campo que estava faltando
    
    # Mapeando os outros campos que existem no SQL para evitar erros futuros
    idade = db.Column(db.Integer)
    email = db.Column(db.String(120))
    cep = db.Column(db.String(10))
    endereco = db.Column(db.Text)
    rg_url = db.Column(db.Text)

class Paciente(db.Model):
    __tablename__ = 'pacientes'
    id = db.Column(db.Integer, primary_key=True)
    nome_completo = db.Column(db.String(100), nullable=False)

class Acolhimento(db.Model):
    __tablename__ = 'acolhimentos'
    id = db.Column(db.Integer, primary_key=True)
    
    paciente_id = db.Column(db.Integer, db.ForeignKey('pacientes.id'))
    decisor_id = db.Column(db.Integer, db.ForeignKey('decisores.id'))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    
    # Dados da Oportunidade (Estes estavam faltando no Python!)
    status = db.Column(db.String(50), default='Novo')
    tipo_internacao = db.Column(db.String(50))
    substancias_uso = db.Column(db.Text)
    comorbidades = db.Column(db.Text)
    investimento_estimado = db.Column(db.String(50))
    urgencia = db.Column(db.String(20), default='Média')
    previsao_internacao = db.Column(db.Date)
    previsao_cura = db.Column(db.Date)

    data_inicio = db.Column(db.DateTime, default=db.func.current_timestamp())

    # Relacionamentos para facilitar a busca no HTML
    paciente = db.relationship('Paciente', backref='acolhimentos')
    decisor = db.relationship('Decisor', backref='acolhimentos')
    agente = db.relationship('Usuario', backref='meus_acolhimentos')
    anexos = db.relationship('Anexo', backref='acolhimento', lazy=True)
    # Relação com tarefas
    tarefas = db.relationship('Tarefa', backref='acolhimento', order_by='Tarefa.data_vencimento.asc()')

class Tarefa(db.Model):
    __tablename__ = 'tarefas'
    id = db.Column(db.Integer, primary_key=True)
    acolhimento_id = db.Column(db.Integer, db.ForeignKey('acolhimentos.id', ondelete='CASCADE'))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    descricao = db.Column(db.Text, nullable=False)
    data_vencimento = db.Column(db.Date, nullable=False)
    hora_vencimento = db.Column(db.Time)
    prioridade = db.Column(db.String(20), default='Média')
    status = db.Column(db.String(20), default='Pendente')
    relatorio = db.Column(db.Text) 

class Atividade(db.Model):
    __tablename__ = 'atividades'
    id = db.Column(db.Integer, primary_key=True)
    acolhimento_id = db.Column(db.Integer, db.ForeignKey('acolhimentos.id', ondelete='CASCADE'))
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    anotacao = db.Column(db.Text, nullable=False)
    data_criacao = db.Column(db.DateTime, default=db.func.current_timestamp())

    # Relações para podermos puxar o nome de quem anotou
    usuario = db.relationship('Usuario', backref='atividades')
    # O backref abaixo permite fazermos "ac.atividades" direto no HTML
    acolhimento = db.relationship('Acolhimento', backref=db.backref('atividades', order_by='Atividade.data_criacao.desc()', cascade='all, delete-orphan'))

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(Usuario, int(user_id))

class Chamado(db.Model):
    __tablename__ = 'chamado'
    id = db.Column(db.Integer, primary_key=True)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'))
    tipo = db.Column(db.String(50)) # Bug, Sugestão, Outros
    assunto = db.Column(db.String(100))
    descricao = db.Column(db.Text)
    anexo = db.Column(db.String(255))
    status = db.Column(db.String(20), default='Aberto') # Aberto, Em Análise, Resolvido
    data_criacao = db.Column(db.DateTime, default=db.func.current_timestamp())
    
    usuario = db.relationship('Usuario', backref='chamados')

class Anexo(db.Model):
    __tablename__ = 'anexos' # Nome da tabela no banco
    id = db.Column(db.Integer, primary_key=True)
    nome_original = db.Column(db.String(255), nullable=False)
    caminho = db.Column(db.String(255), nullable=False)
    data_criacao = db.Column(db.DateTime, default=datetime.utcnow)
    
    # AJUSTE AS STRINGS ABAIXO PARA O NOME REAL DAS TABELAS NO SQL
    acolhimento_id = db.Column(db.Integer, db.ForeignKey('acolhimentos.id'), nullable=False)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
# ==========================================
# 2. ROTAS INSTITUCIONAIS (PÚBLICO)
# ==========================================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/politica-de-privacidade')
def privacidade():
    return render_template('privacidade.html')

@app.route('/termos-de-uso')
def termos():
    return render_template('termos.html')

@app.route('/agente-de-cura')
def agente_de_cura():
    return render_template('agente.html')

# ==========================================
# 3. ROTAS DO CRM (WORKBENCH)
# ==========================================

@app.route('/workbench/acolhimentos')
@login_required
def acolhimentos():
    if current_user.role == 'admin':
        lista = Acolhimento.query.order_by(Acolhimento.data_inicio.desc()).all()
    else:
        lista = Acolhimento.query.filter_by(usuario_id=current_user.id).order_by(Acolhimento.data_inicio.desc()).all()
    
    return render_template('crm/acolhimentos.html', acolhimentos=lista)

@app.route('/api/capturar-lead', methods=['POST'])
def capturar_lead():
    novo_decisor = Decisor(
        nome_completo=request.form.get('nome_familiar'),
        whatsapp=request.form.get('whatsapp'),
        parentesco=request.form.get('parentesco')
    )
    db.session.add(novo_decisor)
    db.session.flush()

    novo_paciente = Paciente(
        nome_completo=request.form.get('nome_paciente')
    )
    db.session.add(novo_paciente)
    db.session.flush()

    # O dono é quem está logado. Se vier da Landing Page, fica sem dono (None).
    dono_id = current_user.id if current_user.is_authenticated else None

    novo_acolhimento = Acolhimento(
        paciente_id=novo_paciente.id,
        decisor_id=novo_decisor.id,
        urgencia=request.form.get('urgencia', 'Média'),
        status='Novo',
        usuario_id=dono_id
    )
    db.session.add(novo_acolhimento)
    db.session.commit()
    
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/workbench/acolhimento/<int:id>')
@login_required
def ver_acolhimento(id):
    # Puxa o acolhimento e já traz o paciente e decisor junto
    oportunidade = Acolhimento.query.get_or_404(id)
    consultores = Usuario.query.filter_by(is_active=True).all() if current_user.role == 'admin' else []
    return render_template('crm/detalhes_acolhimento.html', ac=oportunidade, consultores=consultores)

@app.route('/api/atualizar-acolhimento/<int:id>', methods=['POST'])
@login_required
def atualizar_acolhimento(id):
    ac = Acolhimento.query.get_or_404(id)
    
    # 1. Atualização de Datas (Previsão de Internação e Cura)
    # Usamos try/except ou verificação simples para não quebrar se a data vier vazia
    prev_int = request.form.get('previsao_internacao')
    if prev_int:
        try:
            ac.previsao_internacao = datetime.strptime(prev_int, '%Y-%m-%d').date()
        except ValueError:
            pass # Formato inválido ou vazio
    else:
        ac.previsao_internacao = None

    prev_cura = request.form.get('previsao_cura')
    if prev_cura:
        try:
            ac.previsao_cura = datetime.strptime(prev_cura, '%Y-%m-%d').date()
        except ValueError:
            pass
    else:
        ac.previsao_cura = None

    # 2. Dados da Oportunidade (Acolhimento)
    ac.status = request.form.get('status', ac.status)
    ac.urgencia = request.form.get('urgencia', ac.urgencia)
    ac.tipo_internacao = request.form.get('tipo_internacao', ac.tipo_internacao)
    ac.substancias_uso = request.form.get('substancias_uso', ac.substancias_uso)
    ac.comorbidades = request.form.get('comorbidades', ac.comorbidades)
    ac.investimento_estimado = request.form.get('investimento_estimado', ac.investimento_estimado)

    # 3. Dados do Decisor (Relação entrelaçada)
    if ac.decisor:
        ac.decisor.nome_completo = request.form.get('decisor_nome', ac.decisor.nome_completo)
        ac.decisor.whatsapp = request.form.get('decisor_whatsapp', ac.decisor.whatsapp)
        ac.decisor.parentesco = request.form.get('decisor_parentesco', ac.decisor.parentesco)
        ac.decisor.email = request.form.get('decisor_email', ac.decisor.email)
        ac.decisor.cep = request.form.get('decisor_cep', ac.decisor.cep)
        ac.decisor.endereco = request.form.get('decisor_endereco', ac.decisor.endereco)

    # 4. Dados do Paciente (Relação entrelaçada)
    if ac.paciente:
        ac.paciente.nome_completo = request.form.get('paciente_nome', ac.paciente.nome_completo)
    
    # Salva tudo no Neon
    db.session.commit()
    
    flash('Dossiê atualizado com sucesso!')
    return redirect(url_for('ver_acolhimento', id=ac.id))

@app.route('/api/acolhimento/<int:id>/tarefa', methods=['POST'])
@login_required
def adicionar_tarefa(id):
    # A tarefa pertence obrigatoriamente a quem a criou
    nova_tarefa = Tarefa(
        acolhimento_id=id,
        usuario_id=current_user.id,
        descricao=request.form.get('descricao'),
        data_vencimento=request.form.get('date'),
        prioridade=request.form.get('prioridade', 'Média'),
        status='Pendente'
    )
    db.session.add(nova_tarefa)
    db.session.commit()
    return redirect(url_for('ver_acolhimento', id=id))

@app.route('/workbench/tarefa/<int:id>/concluir', methods=['GET', 'POST'])
@login_required
def concluir_tarefa(id):
    tarefa = Tarefa.query.get_or_404(id)
    
    if request.method == 'POST':
        # Monta o relatório com as respostas do usuário
        resumo = request.form.get('resumo')
        proximo_passo = request.form.get('proximo_passo')
        temperatura = request.form.get('temperatura')
        
        relatorio_completo = f"[{temperatura}] Resumo: {resumo}\nPróximo Passo: {proximo_passo}"
        
        # Salva a tarefa e fecha
        tarefa.relatorio = relatorio_completo
        tarefa.status = 'Concluída'
        
        # Opcional: Já joga o relatório na Timeline do Acolhimento também!
        nota_automatica = Atividade(
            acolhimento_id=tarefa.acolhimento_id,
            usuario_id=current_user.id,
            anotacao=f"✔️ Tarefa Concluída: {tarefa.descricao}\n{relatorio_completo}"
        )
        db.session.add(nota_automatica)
        
        db.session.commit()
        return redirect(url_for('ver_acolhimento', id=tarefa.acolhimento_id))
        
    return render_template('crm/concluir_tarefa.html', tarefa=tarefa)

@app.route('/workbench')
def workbench_index():
    return redirect(url_for('login'))

# 1. Funções Auxiliares
def gerar_senha_inicial(sobrenome):
    numeros = ''.join(random.choices(string.digits, k=5))
    return f"{numeros}{sobrenome.capitalize()}"

def gerar_codigo_2fa():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username_input = request.form.get('username').lower().strip()
        senha_input = request.form.get('password')
        
        user = Usuario.query.filter_by(user_key=username_input).first()
        
        if user and check_password_hash(user.senha, senha_input):
            login_user(user)
            
            # Se for o primeiro acesso ou reset, envia código 2FA
            if user.precisa_mudar_senha:
                user.codigo_verificacao = gerar_codigo_2fa()
                db.session.commit()
                
                # DISPARO DO E-MAIL MARKETING DE SEGURANÇA (HTML)
                try:
                    msg = Message(
                        subject=f"CÓDIGO: {user.codigo_verificacao} - Verificação Elo",
                        sender=app.config['MAIL_DEFAULT_SENDER'],
                        recipients=[user.email]
                    )
                    
                    # Carrega o HTML com o design de elite
                    msg.html = render_template('emails/codigo_seguranca.html', 
                                               nome=user.nome, 
                                               codigo=user.codigo_verificacao)
                    
                    mail.send(msg)
                    print(f"✅ Código de segurança enviado para {user.email}")
                    
                except Exception as e:
                    print(f"❌ Erro ao enviar e-mail de segurança: {str(e)}")
                    flash("Erro ao enviar código de segurança. Tente novamente.", "error")
                    return redirect(url_for('login'))
                
                return redirect(url_for('verificar_2fa'))
            
            return redirect(url_for('dashboard'))
        
        flash("Credenciais inválidas.", "error")
    
    return render_template('crm/login.html')

@app.route('/workbench/dashboard')
@login_required
def dashboard():
    hoje = date.today()
    mes_atual = hoje.month
    ano_atual = hoje.year

    # === SEGURANÇA: DEFINE A BASE DE DADOS BASEADA NO CARGO ===
    if current_user.role == 'admin':
        # Superadmin: Vê absolutamente tudo
        q_acolhimentos = Acolhimento.query
        q_tarefas = Tarefa.query
    else:
        # Consultor (João): Vê apenas o que pertence ao ID dele
        q_acolhimentos = Acolhimento.query.filter_by(usuario_id=current_user.id)
        q_tarefas = Tarefa.query.filter_by(usuario_id=current_user.id)

    # KPIs (Usando data_inicio conforme definido na sua Classe na linha 58)
    leads_hoje = q_acolhimentos.filter(func.date(Acolhimento.data_inicio) == hoje).count()
    em_atendimento = q_acolhimentos.filter(~Acolhimento.status.in_(['Internado (Ganho)', 'Perdido'])).count()
    internacoes_mes = q_acolhimentos.filter(
        Acolhimento.status == 'Internado (Ganho)',
        func.extract('month', Acolhimento.data_inicio) == mes_atual,
        func.extract('year', Acolhimento.data_inicio) == ano_atual
    ).count()

    # BUSCA DE DADOS PARA A LISTAGEM (O que faz os dados aparecerem no HTML)
    acolhimentos = q_acolhimentos.order_by(Acolhimento.data_inicio.desc()).all()
    tarefas = q_tarefas.filter_by(status='Pendente').all()

    # B.I. 1: Linha do Tempo (7 dias)
    datas_timeline = []
    dados_timeline = []
    for i in range(6, -1, -1):
        d = hoje - timedelta(days=i)
        datas_timeline.append(d.strftime('%d/%m'))
        count_dia = q_acolhimentos.filter(func.date(Acolhimento.data_inicio) == d).count()
        dados_timeline.append(count_dia)

    # B.I. 4: Performance por Consultor
    labels_consultor = []
    dados_consultor = []
    if current_user.role == 'admin':
        consultores_raw = db.session.query(Acolhimento.usuario_id, func.count(Acolhimento.id)).group_by(Acolhimento.usuario_id).all()
        for uid, count in consultores_raw:
            user = db.session.get(Usuario, uid) if uid else None
            labels_consultor.append(user.nome if user else "Site Público")
            dados_consultor.append(int(count))
    else:
        labels_consultor = [current_user.nome]
        dados_consultor = [len(acolhimentos)]

    # Taxa de Conversão Global
    total_historico = q_acolhimentos.count()
    total_ganhos = q_acolhimentos.filter_by(status='Internado (Ganho)').count()
    taxa_conversao = int((total_ganhos / total_historico * 100)) if total_historico > 0 else 0

    return render_template(
        'crm/dashboard.html',
        acolhimentos=acolhimentos, tarefas=tarefas, # <--- ENVIANDO OS DADOS PARA A TELA
        leads_hoje=leads_hoje, em_atendimento=em_atendimento, internacoes_mes=internacoes_mes,
        datas_timeline=datas_timeline, dados_timeline=dados_timeline,
        labels_consultor=labels_consultor, dados_consultor=dados_consultor,
        taxa_conversao=taxa_conversao,
        labels_funil=[], dados_funil=[], labels_urgencia=[], dados_urgencia=[] # Fallback para não quebrar o Chart.js
    )

@app.route('/workbench/pacientes')
@login_required
def pacientes():
    return None

# Rotas auxiliares da Sidebar para não dar 404
@app.route('/workbench/tarefas')
@login_required
def listar_tarefas():
    if current_user.role == 'admin':
        tarefas_pendentes = Tarefa.query.filter_by(status='Pendente').order_by(Tarefa.data_vencimento.asc()).all()
        tarefas_concluidas = Tarefa.query.filter_by(status='Concluída').order_by(Tarefa.id.desc()).limit(10).all()
    else:
        tarefas_pendentes = Tarefa.query.filter_by(usuario_id=current_user.id, status='Pendente').order_by(Tarefa.data_vencimento.asc()).all()
        tarefas_concluidas = Tarefa.query.filter_by(usuario_id=current_user.id, status='Concluída').order_by(Tarefa.id.desc()).limit(10).all()
        
    return render_template('crm/tarefas.html', pendentes=tarefas_pendentes, concluidas=tarefas_concluidas)

@app.route('/workbench/suporte')
@login_required
def suporte():
    return "<h1>Suporte - Em construção</h1>"

@app.route('/workbench/perfil')
@login_required
def perfil():
    # 1. Definir Mês e Ano para a Navegação do Calendário
    hoje = datetime.now()
    mes_atual = int(request.args.get('mes', hoje.month))
    ano_atual = int(request.args.get('ano', hoje.year))

    # Proteção de virada de ano (se avançar de Dezembro ou voltar de Janeiro)
    if mes_atual > 12:
        mes_atual = 1
        ano_atual += 1
    elif mes_atual < 1:
        mes_atual = 12
        ano_atual -= 1

    # 2. Gerar a Matriz do Calendário (Dias da semana)
    cal = calendar.monthcalendar(ano_atual, mes_atual)
    nomes_meses = ['', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    nome_mes = nomes_meses[mes_atual]

    # 3. MOTOR DE PERFORMANCE (Estatísticas Completas e à prova de falhas)
    # Total de oportunidades deste consultor
    total_acolhimentos = Acolhimento.query.filter_by(usuario_id=current_user.id).count()
    
    # Contagem de Fechamentos (Usa IN para garantir que pegue qualquer variação de nomenclatura)
    ganhos = Acolhimento.query.filter(
        Acolhimento.usuario_id == current_user.id,
        Acolhimento.status.in_(['Internado (Ganho)', 'Ganho', 'Internado'])
    ).count()

    # Cálculo da taxa de conversão (com proteção contra divisão por zero)
    taxa = round((ganhos / total_acolhimentos * 100), 1) if total_acolhimentos > 0 else 0

    stats = {
        'total_acolhimentos': total_acolhimentos,
        'taxa_conversao': taxa,
        'ganhos': ganhos
    }

    # 4. BUSCA DE TAREFAS (Apenas do mês que está sendo visualizado)
    tarefas_do_mes = Tarefa.query.filter(
        Tarefa.usuario_id == current_user.id,
        extract('month', Tarefa.data_vencimento) == mes_atual,
        extract('year', Tarefa.data_vencimento) == ano_atual
    ).all()

    # Agrupar as tarefas por dia para o HTML ler facilmente (ex: dia 15 tem 2 tarefas)
    tarefas_por_dia = {}
    for t in tarefas_do_mes:
        dia = t.data_vencimento.day
        if dia not in tarefas_por_dia:
            tarefas_por_dia[dia] = []
        tarefas_por_dia[dia].append(t)

    # 5. Enviar tudo pronto para o HTML renderizar
    return render_template(
        'crm/perfil.html', 
        user=current_user, 
        stats=stats, 
        cal=cal, 
        mes=mes_atual, 
        ano=ano_atual, 
        nome_mes=nome_mes, 
        tarefas_por_dia=tarefas_por_dia
    )

@app.route('/workbench/relatorios', methods=['GET', 'POST'])
@login_required
def relatorios():
    # Definição dos campos disponíveis por tabela
    MAPA_CAMPOS = {
        'acolhimentos': {
            'id': 'ID da Oportunidade',
            'status': 'Status do Funil',
            'urgencia': 'Urgência',
            'investimento_estimado': 'Valor Estimado',
            'paciente.nome_completo': 'Nome do Paciente',
            'decisor.nome_completo': 'Nome do Familiar',
            'usuario.nome': 'Consultor Responsável',
            'data_inicio': 'Data de Abertura'
        },
        'tarefas': {
            'id': 'ID da Tarefa',
            'descricao': 'Descrição',
            'prioridade': 'Prioridade',
            'status': 'Status da Atividade',
            'data_vencimento': 'Data de Prazo',
            'acolhimento.paciente.nome_completo': 'Paciente Vinculado'
        }
    }

    if request.method == 'POST':
        tabela = request.form.get('tabela')
        colunas_selecionadas_raw = request.form.getlist('colunas')
        filtro_status = request.form.get('filtro_status')
        exportar = request.form.get('exportar') == 'true'

        # FILTRO DE SEGURANÇA: Limpa colunas invisíveis enviadas por engano pelo HTML
        colunas_selecionadas = [c for c in colunas_selecionadas_raw if c in MAPA_CAMPOS.get(tabela, {})]

        # 1. Início da Query
        if tabela == 'acolhimentos':
            query = Acolhimento.query
        else:
            query = Tarefa.query

        # 2. Regra de Segurança (Superadmin vs Consultor)
        if getattr(current_user, 'role', 'consultor') != 'admin':
            query = query.filter_by(usuario_id=current_user.id)

        # 3. Filtros Dinâmicos (Exemplo por Status)
        if filtro_status:
            query = query.filter_by(status=filtro_status)

        resultados = query.all()

        # 4. Se for apenas para visualizar na tela (Report Builder)
        if not exportar:
            return render_template('crm/relatorios.html', 
                                   mapa=MAPA_CAMPOS, 
                                   resultados=resultados, 
                                   colunas=colunas_selecionadas,
                                   tabela_ativa=tabela)

        # ==========================================
        # 5. GERADOR DE EXCEL PREMIUM (.xlsx)
        # ==========================================
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Relatório de {tabela.capitalize()}"

        # Estilos Corporativos Elo Reencontro
        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_alignment = Alignment(horizontal="center", vertical="center")

        # Escreve o Cabeçalho traduzido
        cabecalho = [MAPA_CAMPOS[tabela][c] for c in colunas_selecionadas]
        ws.append(cabecalho)

        # Aplica o estilo na primeira linha
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment

        # Preenche os dados reais
        for item in resultados:
            linha = []
            for col in colunas_selecionadas:
                val = item
                # Navega por atributos aninhados (Ex: entra no paciente e pega o nome)
                for part in col.split('.'):
                    val = getattr(val, part, 'N/A') if val else 'N/A'
                
                # Tratamento visual para moeda, se necessário
                if col == 'investimento_estimado' and val != 'N/A':
                    linha.append(f"R$ {val}")
                else:
                    linha.append(str(val) if val != None else '-')
            ws.append(linha)

        # Ajuste automático da largura das colunas do Excel
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter # Ex: 'A', 'B'
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            # Dá margem de respiro ao texto na célula
            ws.column_dimensions[col_letter].width = max_length + 2

        # Salva o arquivo final na memória
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        # ==========================================
        # 6. NOMENCLATURA E ENVIO AO NAVEGADOR
        # ==========================================
        id_relatorio = uuid.uuid4().hex[:4].upper()
        data_emissao = datetime.now().strftime('%d%m%Y_%H%M')
        nome_arquivo = f"{id_relatorio}_{data_emissao}_{tabela.capitalize()}.xlsx"

        output = make_response(out.getvalue())
        output.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
        output.headers["Content-type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        
        return output

    # Retorno Padrão (Entrada Inicial na Página)
    return render_template('crm/relatorios.html', mapa=MAPA_CAMPOS, resultados=None)

# --- ROTAS DE CHAMADOS ---
@app.route('/workbench/chamados')
@login_required
def chamados():
    # Admin vê todos os chamados de todo mundo. Consultor vê só os seus.
    if getattr(current_user, 'role', 'consultor') == 'admin':
        lista_chamados = Chamado.query.order_by(Chamado.data_criacao.desc()).all()
    else:
        lista_chamados = Chamado.query.filter_by(usuario_id=current_user.id).order_by(Chamado.data_criacao.desc()).all()
    
    return render_template('crm/chamados.html', chamados=lista_chamados)

@app.route('/workbench/chamados/novo', methods=['POST'])
@login_required
def novo_chamado():
    arquivo = request.files.get('anexo')
    nome_arquivo = None
    
    if arquivo and arquivo.filename != '':
        nome_arquivo = secure_filename(f"{uuid.uuid4().hex[:8]}_{arquivo.filename}")
        arquivo.save(os.path.join(app.config['UPLOAD_FOLDER'], nome_arquivo))

    novo = Chamado(
        usuario_id = current_user.id,
        tipo = request.form.get('tipo'),
        assunto = request.form.get('assunto'),
        descricao = request.form.get('descricao'),
        anexo = nome_arquivo
    )
    db.session.add(novo)
    db.session.commit()
    return redirect(url_for('chamados'))

# --- ROTA DE CONFIGURAÇÕES ---
@app.route('/workbench/configuracoes')
@login_required
def configuracoes():
    if current_user.role != 'admin':
        flash("Acesso restrito a administradores.", "error")
        return redirect(url_for('dashboard'))
    
    usuarios = Usuario.query.order_by(Usuario.nome).all()
    return render_template('crm/configuracoes.html', usuarios=usuarios)

@app.route('/api/usuarios/novo', methods=['POST'])
@login_required
def criar_usuario():
    if current_user.role != 'admin': return "Acesso Negado", 403
    
    nome = request.form.get('nome')
    sobrenome = request.form.get('sobrenome')
    email = request.form.get('email')
    
    # Lógica de Senha Inicial (5 números + Sobrenome)
    senha_temp = gerar_senha_inicial(sobrenome)
    
    novo_user = Usuario(
        nome=nome,
        sobrenome=sobrenome,
        email=email,
        senha=generate_password_hash(senha_temp),
        role=request.form.get('role', 'consultor'),
        precisa_mudar_senha=True,
        is_active=True
    )
    novo_user.user_key = novo_user.gerar_username()
    
    db.session.add(novo_user)
    db.session.commit()

    # Envio do E-mail Marketing de Convite
    try:
        msg = Message("Bem-vindo à Elo - Suas Credenciais", recipients=[email])
        msg.html = render_template('emails/boas_vindas.html', 
                                   nome=nome, 
                                   user_key=novo_user.user_key, 
                                   senha_temporaria=senha_temp)
        mail.send(msg)
        flash(f"Usuário {nome} criado e convite enviado!", "success")
    except Exception as e:
        flash(f"Usuário criado, mas erro ao enviar e-mail: {e}", "warning")

    return redirect(url_for('configuracoes'))

@app.route('/api/usuarios/<int:id>/editar', methods=['POST'])
@login_required
def editar_usuario(id):
    if current_user.role != 'admin': return "Acesso Negado", 403
    
    user = db.session.get(Usuario, id)
    user.nome = request.form.get('nome')
    user.email = request.form.get('email')
    
    # Lógica do Toggle Ativo/Inativo (Checkbox HTML)
    user.is_active = 'is_active' in request.form
    
    db.session.commit()
    flash("Alterações salvas com sucesso!", "success")
    return redirect(url_for('configuracoes'))

# Rota para Bloquear/Editar (Apenas Admin)
@app.route('/workbench/configuracoes/usuario/<int:id>/status', methods=['POST'])
@login_required
def alterar_status_usuario(id):
    if current_user.role != 'admin':
        return "Acesso negado", 403
    
    user = db.session.get(Usuario, id)
    # Lógica simples: se você tiver uma coluna 'ativo' no banco:
    # user.ativo = not user.ativo 
    db.session.commit()
    return redirect(url_for('configuracoes'))

@app.route('/verificar-2fa', methods=['GET', 'POST'])
@login_required
def verificar_2fa():
    if request.method == 'POST':
        codigo = request.form.get('codigo').upper()
        nova_senha = request.form.get('nova_senha')
        confirmar = request.form.get('confirmar_senha')
        
        if codigo == current_user.codigo_verificacao:
            if nova_senha == confirmar and len(nova_senha) >= 8:
                current_user.senha = generate_password_hash(nova_senha)
                current_user.precisa_mudar_senha = False
                current_user.codigo_verificacao = None
                db.session.commit()
                flash("Senha atualizada com sucesso!", "success")
                return redirect(url_for('dashboard'))
            else:
                flash("As senhas não coincidem ou são muito curtas.", "error")
        else:
            flash("Código de verificação inválido.", "error")
            
    return render_template('crm/verificar_2fa.html')

@app.route('/api/usuarios/reset-total', methods=['POST'])
@login_required
def reset_total():
    if current_user.role != 'admin': return "Acesso Negado", 403
    
    usuarios = Usuario.query.filter(Usuario.role != 'admin').all()
    for u in usuarios:
        # Reset Total: 5 números + Nome (conforme solicitado)
        nums = ''.join(random.choices(string.digits, k=5))
        senha_reset = f"{nums}{u.nome.capitalize()}"
        
        u.senha = generate_password_hash(senha_reset)
        u.precisa_mudar_senha = True
        
        # Enviar e-mail simples de aviso
        msg = Message("Alerta de Segurança - Reset de Senha", recipients=[u.email])
        msg.body = f"Todas as senhas foram resetadas pelo admin. Sua nova senha: {senha_reset}"
        mail.send(msg)
        
    db.session.commit()
    flash("Toda a equipe foi resetada!", "success")
    return redirect(url_for('configuracoes'))

@app.route('/api/usuarios/<int:id>/reset-senha', methods=['POST'])
@login_required
def resetar_senha_usuario(id):
    if current_user.role != 'admin': return "Acesso Negado", 403
    
    user = db.session.get(Usuario, id)
    # Reset Individual: 5 números + Sobrenome
    nova_senha = gerar_senha_inicial(user.sobrenome)
    
    user.senha = generate_password_hash(nova_senha)
    user.precisa_mudar_senha = True
    db.session.commit()
    
    # E-mail de Notificação de Reset
    try:
        msg = Message("Sua Senha Elo foi Resetada", recipients=[user.email])
        msg.body = f"Olá {user.nome}, sua nova senha temporária é: {nova_senha}"
        mail.send(msg)
        flash(f"Senha de {user.nome} resetada!", "success")
    except Exception as e:
        flash("Senha resetada no banco, mas erro no e-mail.", "error")
        
    return redirect(url_for('configuracoes'))

@app.route('/api/acolhimento/<int:id>/anexo', methods=['POST'])
@login_required
def upload_anexo(id):
    if 'arquivo' not in request.files:
        # CORREÇÃO AQUI: de 'detalhes_acolhimento' para 'ver_acolhimento'
        return redirect(url_for('ver_acolhimento', id=id))
    
    file = request.files['arquivo']
    if file.filename == '':
        return redirect(url_for('ver_acolhimento', id=id))

    if file:
        filename = secure_filename(file.filename)
        unique_name = f"{int(datetime.now().timestamp())}_{filename}"
        
        if not os.path.exists(app.config['UPLOAD_FOLDER']):
            os.makedirs(app.config['UPLOAD_FOLDER'])
            
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], unique_name))

        novo_anexo = Anexo(
            nome_original=filename,
            caminho=unique_name,
            acolhimento_id=id,
            usuario_id=current_user.id
        )
        db.session.add(novo_anexo)
        db.session.commit()
        
        # Opcional: Garante que a sessão está limpa para a próxima leitura
        db.session.expire_all() 
        
        flash("Documento anexado!", "success")
    return redirect(url_for('ver_acolhimento', id=id))

@app.route('/uploads/anexos/<filename>')
@login_required
def custom_static(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route('/api/anexo/<int:id>/deletar', methods=['POST'])
@login_required
def deletar_anexo(id):
    anexo = Anexo.query.get_or_404(id)
    id_acolhimento = anexo.acolhimento_id
    
    try:
        # 1. Tenta apagar o arquivo físico da pasta uploads
        caminho_arquivo = os.path.join(app.config['UPLOAD_FOLDER'], anexo.caminho)
        if os.path.exists(caminho_arquivo):
            os.remove(caminho_arquivo)
            
        # 2. Apaga o registro no banco de dados Neon
        db.session.delete(anexo)
        db.session.commit()
        flash("Arquivo removido com sucesso!", "success")
        
    except Exception as e:
        db.session.rollback()
        print(f"Erro ao deletar: {e}")
        flash("Não foi possível remover o arquivo físico.", "error")

    # 3. Redireciona para a página do paciente (ver_acolhimento)
    return redirect(url_for('ver_acolhimento', id=id_acolhimento))

@app.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('index'))

if __name__ == "__main__":
    # O código aqui dentro será IGNORADO pela Vercel
    app.run(debug=True)