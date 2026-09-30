import pandas as pd
from openpyxl import load_workbook
import os
from datetime import date, timedelta

os.makedirs('data', exist_ok=True)

hoje = date.today()

dados = [
    ('PED-001', 'João Silva',        'Sudeste',      'São Paulo',      hoje,                  'entregue'),
    ('PED-002', 'Maria Souza',       'Nordeste',     'Fortaleza',      hoje,                  'em_rota'),
    ('PED-003', 'Carlos Lima',       'Sul',          'Curitiba',       hoje - timedelta(1),   'pendente'),
    ('PED-004', 'Ana Costa',         'Norte',        'Belém',          hoje - timedelta(3),   'falhou'),
    ('PED-005', 'Pedro Mendes',      'Centro-Oeste', 'Brasília',       hoje,                  'pendente'),
    ('PED-006', 'Lucia Ferreira',    'Sudeste',      'Rio de Janeiro', hoje - timedelta(2),   'entregue'),
    ('PED-007', 'Roberto Santos',    'Nordeste',     'Recife',         hoje - timedelta(1),   'em_rota'),
    ('PED-008', 'Fernanda Oliveira', 'Sul',          'Porto Alegre',   hoje,                  'pendente'),
    ('PED-009', 'Marcos Alves',      'Norte',        'Manaus',         hoje - timedelta(2),   'entregue'),
    ('PED-010', 'Juliana Rocha',     'Centro-Oeste', 'Campo Grande',   hoje - timedelta(3),   'falhou'),
    ('PED-011', 'André Pereira',     'Sudeste',      'Belo Horizonte', hoje - timedelta(1),   'em_rota'),
    ('PED-012', 'Camila Torres',     'Nordeste',     'Salvador',       hoje,                  'pendente'),
    ('PED-013', 'Diego Ribeiro',     'Sul',          'Florianópolis',  hoje - timedelta(2),   'entregue'),
    ('PED-014', 'Patricia Nunes',    'Norte',        'Porto Velho',    hoje - timedelta(1),   'pendente'),
    ('PED-015', 'Thiago Carvalho',   'Sudeste',      'Campinas',       hoje - timedelta(3),   'em_rota'),
    ('PED-016', 'Rafael Teste',      'Nordeste',     'Natal',          hoje,                  'pendente'),
]

df = pd.DataFrame(dados, columns=[
    'numero_pedido', 'nome_cliente', 'regiao', 'cidade', 'data_prevista', 'status'
])

df.to_excel('data/entregas.xlsx', index=False)

# Protege a planilha com senha
wb = load_workbook('data/entregas.xlsx')
ws = wb.active
ws.protection.sheet = True
ws.protection.password = 'logitrack123'
ws.protection.selectLockedCells = False
wb.save('data/entregas.xlsx')

print('Planilha criada e protegida: data/entregas.xlsx')
print('Senha para edição no Excel: logitrack123')
print(f'\n{len(df)} pedidos registrados (incluindo PED-016 de teste).')
