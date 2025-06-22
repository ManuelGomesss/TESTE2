import os
import re
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
from werkzeug.utils import secure_filename
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from werkzeug.security import generate_password_hash, check_password_hash

# --- Inicialização da App ---
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET', 'troque_esta_chave_por_uma_segura')
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- Configuração do Banco de Dados ---
engine = create_engine('sqlite:///ranking.db', connect_args={'check_same_thread': False})
Base = declarative_base()
SessionLocal = sessionmaker(bind=engine)

class User(Base):
    __tablename__ = 'users'
    id = Column(Integer, primary_key=True)
    student_number = Column(String, unique=True)
    name = Column(String)
    password_hash = Column(String)
    average = Column(Float)
    public_choice = Column(String)

Base.metadata.create_all(engine)

# --- Decorator para rotas protegidas ---
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

# --- Rotas de Autenticação ---
@app.route('/', methods=['GET','POST'])
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        number = request.form['student_number'].strip()
        password = request.form['password']
        db = SessionLocal()
        user = db.query(User).filter_by(student_number=number).first()
        db.close()
        if user and check_password_hash(user.password_hash, password):
            session.clear()
            session['user_id'] = user.id
            return redirect(url_for('import_6ano'))
        flash('Credenciais inválidas.')
    return render_template('login.html')

@app.route('/register', methods=['GET','POST'])
def register():
    if request.method == 'POST':
        name   = request.form.get('name','').strip()
        number = request.form.get('student_number','').strip()
        password = request.form.get('password','')
        accept = request.form.get('accept')
        if not (name and number and password and accept):
            flash('Preencha todos os campos e aceite os termos.')
        elif not re.fullmatch(r'2019\d{3}', number):
            flash('Nº de aluno inválido (deve começar por 2019).')
        else:
            db = SessionLocal()
            if db.query(User).filter_by(student_number=number).first():
                flash('Nº de aluno já registado. Faça login.')
                db.close()
                return redirect(url_for('login'))
            hashed = generate_password_hash(password)
            user = User(student_number=number, name=name, password_hash=hashed)
            db.add(user)
            db.commit()
            db.close()
            flash('Registo efetuado! Inicie sessão.')
            return redirect(url_for('login'))
    return render_template('register.html')

@app.route('/forgot', methods=['GET','POST'])
def forgot():
    if request.method=='POST':
        number = request.form['student_number'].strip()
        print(f"[ADMIN] Pedido recuperação password: {number}")
        flash('Admin notificado. Será contactado em breve.')
        return redirect(url_for('login'))
    return render_template('forgot.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# --- Função interna de finalização do 6º ano ---
def _finalize_6ano(df):
    df2 = df.rename(columns={'ECTS UC':'ECTS','Avaliação Nota':'Grade'}) if 'ECTS UC' in df.columns else df.copy()
    df2 = df2.dropna(subset=['ECTS','Grade'])
    df2['ECTS'] = pd.to_numeric(df2['ECTS'], errors='coerce')
    df2['Grade'] = pd.to_numeric(df2['Grade'], errors='coerce')
    sum_ects = df2['ECTS'].sum()
    total_weighted = (df2['Grade'] * df2['ECTS']).sum()
    final_avg = round(total_weighted / sum_ects, 2)
    session['sum_ects_6ano'] = sum_ects
    session['num_6ano'] = total_weighted
    session['Y'] = final_avg
    return redirect(url_for('results_6ano'))

# --- Rotas do fluxo 6º Ano ---
@app.route('/import-6ano', methods=['GET','POST'])
@login_required
def import_6ano():
    if request.method == 'POST':
        f = request.files.get('file')
        if not f:
            return render_template('import_6ano.html', error='Nenhum ficheiro selecionado.')
        filename = secure_filename(f.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        f.save(path)

        try:
            xls = pd.ExcelFile(path)
        except Exception as e:
            return render_template('import_6ano.html', error=f'Erro ao ler Excel: {e}')

        required = ['UC', 'ECTS UC', 'Avaliação Nota']
        df = None
        for sheet in xls.sheet_names:
            tmp = pd.read_excel(path, sheet_name=sheet)
            if all(col in tmp.columns for col in required):
                df = tmp[required].copy()
                break

        if df is None:
            return render_template('import_6ano.html',
                                   error='Não encontrei UC, ECTS UC e Avaliação Nota em nenhuma folha.')

        df = df[~(df['UC'].astype(str).str.startswith('Opcional', na=False) &
                  df['Avaliação Nota'].isna())]

        # Guarda DataFrame e filtra missing corretamente
        session['import_df'] = df.to_dict('records')
        df2 = pd.DataFrame(session['import_df'])
        missing = df2[df2['Avaliação Nota'].isna()][['UC', 'ECTS UC']].to_dict('records')

        if any('Estágio Profissionalizante' in r['UC'] for r in missing):
            return redirect(url_for('manual_input'))
        if missing:
            return render_template('fill_missing.html', missing=missing)
        return _finalize_6ano(df2)

    return render_template('import_6ano.html')

@app.route('/fill-missing', methods=['POST'])
@login_required
def fill_missing():
    records = session.get('import_df', [])
    df = pd.DataFrame.from_records(records)
    for uc in df['UC'].unique():
        if df.loc[df['UC']==uc,'Avaliação Nota'].isna().any():
            val = request.form.get(uc)
            try:
                n = float(val); assert 0<=n<=20
            except:
                return render_template('fill_missing.html', missing=[{'UC':uc,'ECTS UC':df.loc[df['UC']==uc,'ECTS UC'].iloc[0]}], error=f'Nota inválida para {uc}')
            df.loc[df['UC']==uc,'Avaliação Nota'] = n
    return _finalize_6ano(df)

@app.route('/manual-input', methods=['GET','POST'])
@login_required
def manual_input():
    records = session.get('import_df', [])
    df = pd.DataFrame.from_records(records)
    defaults = {'preparacao':17,'opcional4':20,'C':18,'GO':18,'MI':20,'MGF':16,'PED':19,'SM':17,'RF':18}
    if request.method=='POST':
        try:
            vals = {k: float(request.form.get(k)) for k in defaults}
            if not all(0<=v<=20 for v in vals.values()): raise ValueError
        except:
            return render_template('manual_input.html', defaults=defaults, error='Notas inválidas. Utilize 0–20.')
        stage_df = pd.DataFrame([
            {'UC':'Preparação Para a Prática Clínica','ECTS':3,'Grade':vals['preparacao']},
            {'UC':'Opcional 4 (6º ano)','ECTS':3,'Grade':vals['opcional4']},
            {'UC':'Cirurgia Geral','ECTS':8,'Grade':vals['C']},
            {'UC':'Ginecologia e Obstetrícia','ECTS':6,'Grade':vals['GO']},
            {'UC':'Medicina Interna','ECTS':9,'Grade':vals['MI']},
            {'UC':'Medicina Geral e Familiar','ECTS':6,'Grade':vals['MGF']},
            {'UC':'Pediatria','ECTS':7,'Grade':vals['PED']},
            {'UC':'Saúde Mental','ECTS':6,'Grade':vals['SM']},
            {'UC':'Relatório Final','ECTS':12,'Grade':vals['RF']},
        ])
        df_clean = df[~df['UC'].str.contains('Estágio Profissionalizante', na=False)]
        df_clean = df_clean.rename(columns={'ECTS UC':'ECTS','Avaliação Nota':'Grade'})
        full_df = pd.concat([df_clean, stage_df], ignore_index=True)
        return _finalize_6ano(full_df)
    return render_template('manual_input.html', defaults=defaults)

@app.route('/results-6ano', methods=['GET','POST'])
@login_required
def results_6ano():
    Y = session.get('Y')
    sum_ects = session.get('sum_ects_6ano')
    total_num = session.get('num_6ano')
    if None in (Y, sum_ects, total_num):
        return redirect(url_for('import_6ano'))
    if request.method=='POST':
        choice = request.form.get('public_choice','none')
        db = SessionLocal()
        user = db.query(User).filter_by(id=session['user_id']).first()
        user.average = Y
        user.public_choice = choice
        db.commit(); db.close()
        return redirect(url_for('ranking'))
    return render_template('results_6ano.html', Y=Y, sum_ects=sum_ects, num=total_num)

@app.route('/ranking')
@login_required
def ranking():
    db = SessionLocal()
    # só traz quem já tem média calculada
    users = (
        db.query(User)
          .filter(User.average.isnot(None))
          .order_by(User.average.desc())
          .all()
    )
    db.close()

    # calcula a média geral
    if users:
        course_mean = round(
            sum(u.average for u in users) / len(users),
            2
        )
    else:
        course_mean = None

    return render_template(
        'ranking.html',
        users=users,
        course_mean=course_mean
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
