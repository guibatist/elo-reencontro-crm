import os
from flask import Flask, render_template, redirect, url_for, request, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import check_password_hash
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL')
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'elo-secret-123')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

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
    user_key = db.Column(db.String(100), unique=True, nullable=False)

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
    return Usuario.query.get(int(user_id))


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
    # Puxa todos os acolhimentos ativos
    lista = Acolhimento.query.order_by(Acolhimento.data_inicio.desc()).all()
    return render_template('crm/acolhimentos.html', acolhimentos=lista)

# Atualize a captura do Lead para criar os registros entrelaçados
@app.route('/api/capturar-lead', methods=['POST'])
def capturar_lead():
    # Cria o Decisor
    novo_decisor = Decisor(
        nome_completo = request.form.get('nome_familiar'),
        whatsapp = request.form.get('whatsapp'),
        parentesco = request.form.get('parentesco')
    )
    db.session.add(novo_decisor)
    db.session.flush()

    # Cria o Paciente
    novo_paciente = Paciente(
        nome_completo = request.form.get('nome_paciente')
    )
    db.session.add(novo_paciente)
    db.session.flush()

    # Cria o Acolhimento
    novo_acolhimento = Acolhimento(
        paciente_id = novo_paciente.id,
        decisor_id = novo_decisor.id,
        urgencia = request.form.get('urgencia', 'Média'),
        status = 'Novo'
    )
    db.session.add(novo_acolhimento)
    db.session.commit()
    
    # Se veio do CRM, volta pro CRM. Se veio do site, volta pro site.
    return redirect(request.referrer or url_for('index'))

@app.route('/workbench/acolhimento/<int:id>')
@login_required
def ver_acolhimento(id):
    # Puxa o acolhimento e já traz o paciente e decisor junto
    oportunidade = Acolhimento.query.get_or_404(id)
    return render_template('crm/detalhes_acolhimento.html', ac=oportunidade)

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

@app.route('/api/acolhimento/<int:id>/nota', methods=['POST'])
@login_required
def adicionar_nota(id):
    texto = request.form.get('anotacao')
    if texto:
        nova_nota = Atividade(
            acolhimento_id=id,
            usuario_id=current_user.id,
            anotacao=texto
        )
        db.session.add(nova_nota)
        db.session.commit()
    return redirect(url_for('ver_acolhimento', id=id))

@app.route('/api/acolhimento/<int:id>/tarefa', methods=['POST'])
@login_required
def adicionar_tarefa(id):
    # Pega os valores do formulário HTML
    data_form = request.form.get('date') # O HTML manda 'date'
    hora_form = request.form.get('hora') # Vamos adicionar esse campo no HTML
    
    nova_tarefa = Tarefa(
        acolhimento_id=id,
        usuario_id=current_user.id,
        descricao=request.form.get('descricao'),
        data_vencimento=data_form,
        hora_vencimento=hora_form if hora_form else None,
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
    return redirect(url_for('login_crm'))

@app.route('/workbench/login', methods=['GET', 'POST'])
def login_crm():
    if request.method == 'POST':
        u_key = request.form.get('user_key')
        u_pass = request.form.get('password')
        agente = Usuario.query.filter_by(user_key=u_key).first()
        if agente and check_password_hash(agente.senha, u_pass):
            login_user(agente)
            return redirect(url_for('dashboard'))
        flash('Acesso Negado. Verifique os dados.')
    return render_template('crm/login.html')

@app.route('/workbench/dashboard')
@login_required
def dashboard():
    # Agora o Dashboard busca dados da nova tabela 'Acolhimento'
    total_acolhimentos = Acolhimento.query.count()
    novos = Acolhimento.query.filter_by(status='Novo').count()
    
    return render_template('crm/dashboard.html', total=total_acolhimentos, novos=novos)

@app.route('/workbench/pacientes')
@login_required
def pacientes():
    return None

# Rotas auxiliares da Sidebar para não dar 404
@app.route('/workbench/tarefas')
@login_required
def listar_tarefas():
    # Puxa todas as tarefas pendentes do usuário logado
    tarefas_pendentes = Tarefa.query.filter_by(usuario_id=current_user.id, status='Pendente').order_by(Tarefa.data_vencimento.asc()).all()
    tarefas_concluidas = Tarefa.query.filter_by(usuario_id=current_user.id, status='Concluída').order_by(Tarefa.data_vencimento.desc()).limit(10).all()
    return render_template('crm/tarefas.html', pendentes=tarefas_pendentes, concluidas=tarefas_concluidas)

@app.route('/workbench/suporte')
@login_required
def suporte():
    return "<h1>Suporte - Em construção</h1>"

@app.route('/workbench/perfil')
@login_required
def perfil():
    # Pegamos o total de acolhimentos deste usuário para as estatísticas
    total = Acolhimento.query.filter_by(usuario_id=current_user.id).count()
    ganhos = Acolhimento.query.filter_by(usuario_id=current_user.id, status='Internado (Ganho)').count()
    
    # Cálculo simples de conversão
    taxa = round((ganhos / total * 100), 1) if total > 0 else 0
    
    # Criamos um dicionário de estatísticas para o HTML usar
    stats = {
        'total_acolhimentos': total,
        'taxa_conversao': taxa
    }

    # PASSAMOS O current_user COM O NOME DE 'user' PARA O HTML
    return render_template('crm/perfil.html', user=current_user, stats=stats)

@app.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)