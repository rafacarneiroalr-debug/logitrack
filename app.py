from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from openpyxl import load_workbook
from datetime import date
import os

app = Flask(__name__)
app.secret_key = 'logitrack-secret-2026'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'data', 'entregas.xlsx')
EXCEL_PASSWORD = 'logitrack123'

STATUS_LABELS = {
    'pendente': 'Pendente',
    'em_rota':  'Em Rota',
    'entregue': 'Entregue',
    'falhou':   'Falhou',
}

STATUS_COLORS = {
    'pendente': 'warning',
    'em_rota':  'primary',
    'entregue': 'success',
    'falhou':   'danger',
}

COLUNAS = ['numero_pedido', 'nome_cliente', 'regiao', 'cidade', 'data_prevista', 'status']
REGIOES = ['Norte', 'Nordeste', 'Centro-Oeste', 'Sudeste', 'Sul']


def load_data():
    df = pd.read_excel(EXCEL_PATH, dtype={'numero_pedido': str})
    df['data_prevista'] = pd.to_datetime(df['data_prevista']).dt.date
    df['status'] = df['status'].fillna('pendente').str.strip()
    return df


def enrich(df):
    today = date.today()
    records = df.to_dict('records')
    for r in records:
        r['data_fmt'] = r['data_prevista'].strftime('%d/%m/%Y')
        r['atrasado'] = (
            r['data_prevista'] < today and
            r['status'] not in ('entregue', 'falhou')
        )
        r['status_label'] = STATUS_LABELS.get(r['status'], r['status'])
        r['status_color'] = STATUS_COLORS.get(r['status'], 'secondary')
    return records


def protect_sheet(ws):
    ws.protection.sheet = True
    ws.protection.password = EXCEL_PASSWORD
    ws.protection.selectLockedCells = False


def add_pedido_excel(dados):
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.append([dados[c] for c in COLUNAS])
    protect_sheet(ws)
    wb.save(EXCEL_PATH)


def update_status_excel(numero_pedido, novo_status):
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active
    header = [cell.value for cell in ws[1]]
    col_numero = header.index('numero_pedido')
    col_status = header.index('status')

    for row in ws.iter_rows(min_row=2):
        if str(row[col_numero].value) == str(numero_pedido):
            row[col_status].value = novo_status
            break

    protect_sheet(ws)
    wb.save(EXCEL_PATH)


# ── Dashboard ──────────────────────────────────────────────────────────────────

@app.route('/')
def dashboard():
    df = load_data()
    today = date.today()

    por_status = df['status'].value_counts().to_dict()
    por_regiao = df['regiao'].value_counts().to_dict()

    atrasados = enrich(
        df[(df['data_prevista'] < today) & (~df['status'].isin(['entregue', 'falhou']))]
    )

    return render_template('dashboard.html',
        total=len(df),
        por_status=por_status,
        por_regiao=por_regiao,
        atrasados=atrasados,
        hoje=today.strftime('%d/%m/%Y'),
    )


# ── Pedidos ────────────────────────────────────────────────────────────────────

@app.route('/pedidos')
def pedidos():
    df = load_data()

    filtro_status = request.args.get('status', '')
    filtro_regiao = request.args.get('regiao', '')
    busca = request.args.get('busca', '')

    if filtro_status:
        df = df[df['status'] == filtro_status]
    if filtro_regiao:
        df = df[df['regiao'] == filtro_regiao]
    if busca:
        mask = (
            df['numero_pedido'].str.contains(busca, case=False, na=False) |
            df['nome_cliente'].str.contains(busca, case=False, na=False)
        )
        df = df[mask]

    return render_template('pedidos.html',
        pedidos=enrich(df),
        filtro_status=filtro_status,
        filtro_regiao=filtro_regiao,
        busca=busca,
        status_labels=STATUS_LABELS,
        regioes=REGIOES,
    )


@app.route('/pedidos/novo', methods=['GET', 'POST'])
def novo_pedido():
    if request.method == 'POST':
        numero  = request.form.get('numero_pedido', '').strip()
        nome    = request.form.get('nome_cliente', '').strip()
        regiao  = request.form.get('regiao', '').strip()
        cidade  = request.form.get('cidade', '').strip()
        data_s  = request.form.get('data_prevista', '').strip()
        status  = request.form.get('status', 'pendente').strip()

        if not all([numero, nome, regiao, cidade, data_s]):
            flash('Preencha todos os campos obrigatórios.', 'danger')
            return render_template('novo_pedido.html', regioes=REGIOES, status_labels=STATUS_LABELS)

        df = load_data()
        if numero in df['numero_pedido'].values:
            flash(f'O número de pedido "{numero}" já existe.', 'danger')
            return render_template('novo_pedido.html', regioes=REGIOES, status_labels=STATUS_LABELS)

        try:
            add_pedido_excel({
                'numero_pedido': numero,
                'nome_cliente':  nome,
                'regiao':        regiao,
                'cidade':        cidade,
                'data_prevista': data_s,
                'status':        status,
            })
            flash(f'Pedido {numero} cadastrado com sucesso!', 'success')
            return redirect(url_for('pedidos'))
        except PermissionError:
            flash('Não foi possível salvar: feche o arquivo entregas.xlsx no Excel e tente novamente.', 'danger')

    return render_template('novo_pedido.html', regioes=REGIOES, status_labels=STATUS_LABELS)


@app.route('/pedidos/<numero_pedido>/status', methods=['POST'])
def atualizar_status(numero_pedido):
    novo_status = request.form.get('novo_status', '').strip()

    if novo_status not in STATUS_LABELS:
        flash('Status inválido.', 'danger')
        return redirect(url_for('pedidos'))

    try:
        update_status_excel(numero_pedido, novo_status)
        flash(
            f'Status de {numero_pedido} atualizado para "{STATUS_LABELS[novo_status]}".',
            'success'
        )
    except PermissionError:
        flash('Não foi possível salvar: feche o arquivo entregas.xlsx no Excel e tente novamente.', 'danger')

    return redirect(url_for('pedidos'))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
