import os
from flask import Flask, render_template, redirect, url_for, request, flash, make_response
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import check_password_hash
from dotenv import load_dotenv
from datetime import date, datetime, timedelta
import calendar
from sqlalchemy import extract, func
import io
import uuid
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


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
    role = db.Column(db.String(20), default='consultor')

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

    # A MÁGICA ACONTECE AQUI:
    # Ele tenta pegar o ID oculto que o HTML enviou. 
    # Se não tiver (porque veio da Landing Page pública), ele salva como None (sem dono).
    usuario_id_form = request.form.get('usuario_id')
    dono_id = int(usuario_id_form) if usuario_id_form else None

    # Cria o Acolhimento
    novo_acolhimento = Acolhimento(
        paciente_id = novo_paciente.id,
        decisor_id = novo_decisor.id,
        urgencia = request.form.get('urgencia', 'Média'),
        status = 'Novo',
        usuario_id = dono_id  # <--- Salva o ID capturado
    )
    db.session.add(novo_acolhimento)
    db.session.commit()
    
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
    hoje = date.today()
    mes_atual = hoje.month
    ano_atual = hoje.year

    # KPIs Básicos (Cartões do Topo)
    leads_hoje = Acolhimento.query.filter(func.date(Acolhimento.data_inicio) == hoje).count()
    em_atendimento = Acolhimento.query.filter(~Acolhimento.status.in_(['Internado (Ganho)', 'Perdido'])).count()
    internacoes_mes = Acolhimento.query.filter(
        Acolhimento.status == 'Internado (Ganho)',
        func.extract('month', Acolhimento.data_inicio) == mes_atual,
        func.extract('year', Acolhimento.data_inicio) == ano_atual
    ).count()

    # B.I. 1: Linha do Tempo (Últimos 7 dias)
    datas_timeline = []
    dados_timeline = []
    for i in range(6, -1, -1):
        d = hoje - timedelta(days=i)
        datas_timeline.append(d.strftime('%d/%m'))
        # Conta leads exatos daquele dia
        count_dia = Acolhimento.query.filter(func.date(Acolhimento.data_inicio) == d).count()
        dados_timeline.append(count_dia)

    # B.I. 2: Funil de Status
    status_counts = db.session.query(Acolhimento.status, func.count(Acolhimento.id)).group_by(Acolhimento.status).all()
    labels_funil = [str(s[0]) if s[0] else "Sem Status" for s in status_counts]
    dados_funil = [int(s[1]) for s in status_counts]

    # B.I. 3: Mapa de Urgência
    urgencia_counts = db.session.query(Acolhimento.urgencia, func.count(Acolhimento.id)).group_by(Acolhimento.urgencia).all()
    labels_urgencia = [str(u[0]) if u[0] else "N/A" for u in urgencia_counts]
    dados_urgencia = [int(u[1]) for u in urgencia_counts]

    # B.I. 4: Performance por Consultor
    consultores_raw = db.session.query(Acolhimento.usuario_id, func.count(Acolhimento.id)).group_by(Acolhimento.usuario_id).all()
    labels_consultor = []
    dados_consultor = []
    for uid, count in consultores_raw:
        if uid:
            user = db.session.get(Usuario, uid)
            labels_consultor.append(user.nome.split()[0] if user else "Sistema")
        else:
            labels_consultor.append("Site Público")
        dados_consultor.append(int(count))

    # B.I. 5: Taxa de Conversão Global
    total_historico = Acolhimento.query.count()
    total_ganhos = Acolhimento.query.filter_by(status='Internado (Ganho)').count()
    taxa_conversao = int((total_ganhos / total_historico * 100)) if total_historico > 0 else 0

    return render_template(
        'crm/dashboard.html',
        leads_hoje=leads_hoje, em_atendimento=em_atendimento, internacoes_mes=internacoes_mes,
        datas_timeline=datas_timeline, dados_timeline=dados_timeline,
        labels_funil=labels_funil, dados_funil=dados_funil,
        labels_urgencia=labels_urgencia, dados_urgencia=dados_urgencia,
        labels_consultor=labels_consultor, dados_consultor=dados_consultor,
        taxa_conversao=taxa_conversao
    )

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

@app.route('/logout')
def logout():
    logout_user()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)