'''
2 fontes de informações:
- Dados do painel de monitoramento RAPP (feito no GAS)
- Dados do Avalia Potiguar

Dados do painel de monitoramento RAPP (feito no GAS):
Acesso à planilha "Monitoramento_RAPP_2026_2º Ciclo" > Arquivo > Baixar > Microsoft Excel (.xlsx)

Nos dados do painel de monitoramento RAPP (feito no GAS), os dados de prova são segmentados por componente-estudante.
Os dados de agendamento são segmentados por componente-turma.

Dados do Avalia Potiguar:
Login na painel > Acompanhamento RAPP (barra lateral esquerda) > Exportar Excel.

Os dados do Avalia Potiguar são segmentados por componente-estudante.

São informações iguais para 2 bases de dados diferentes.
RAPP: slides 5, 6 e 7
Avalia Potiguar: slides 8, 9 e 10

São informações de quantidade realizada e quantidade total/ prevista para:
- Agendamentos;
- Estudantes;
- Componentes.

Variáveis para serem criadas para a apresentação do GT:
- perc_agend_avalia
- agend_avalia
- barra_agend_avalia
- perc_estud_avalia
- estud_avalia
- barra_estud_avalia
- perc_comp_avalia
- comp_avalia
- barra_comp_avalia
- graf_perc_agendamentos_direc_avalia
- liberad_avalia
- graf_perc_liberados_direc_avalia
- perc_agend_plurall
- agend_plurall
- barra_agend_plurall
- perc_estud_plurall
- estud_plurall
- barra_estud_plurall
- perc_comp_plurall
- comp_plurall
- barra_comp_plurall
- graf_perc_agendamentos_direc_plurall
- liberad_plurall
- graf_perc_liberados_direc_plurall


'''
# Importação das bibliotecas
import os
import glob
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import warnings
warnings.filterwarnings('ignore')
from tqdm import tqdm
import matplotlib.ticker as mtick
from matplotlib.patches import FancyBboxPatch
from datetime import datetime


# AVALIA POTIGUAR
'''
Variáveis:
- perc_agend_avalia
- agend_avalia
- barra_agend_avalia
- perc_estud_avalia
- estud_avalia
- barra_estud_avalia
- perc_comp_avalia
- comp_avalia
- barra_comp_avalia
- graf_perc_agendamentos_direc_avalia
- liberad_avalia
- graf_perc_liberados_direc_avalia
'''

# Data de hoje
data_hoje = datetime.now().strftime('%d/%m/%Y')

# Pasta de destino dos gráficos
pasta_graficos = os.path.join(os.getcwd(), 'graficos_gerados')

# Carregar base de dados do Avalia Potiguar:
df_avalia = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\relatorio_acompanhamento_rapp_ept_div.xlsx', sheet_name='Relatório RAPP')

# AGENDAMENTO - AVALIA POTIGUAR
# 'agend_avalia' é composto por 'componentes agendados' e total previstos de agendamentos (agendados + não agendados)
# Não agendados é composto por: 'SITUAÇÃO DA AVALIAÇÃO' = 'NÃO AGENDADA', 'REPROVADA' e 'CANCELADA'. Restante é agendado.

# Total de componentes
total_agend_avalia = df_avalia['SITUAÇÃO DA AVALIAÇÃO'].count()

# Componentes agendados
# Situações que representam componentes não agendados
situacoes_nao_agendadas = [
    'NÃO AGENDADA',
    'NÃO AGENDADO',
    'REPROVADA',
    'REPROVADO',
    'CANCELADA',
    'CANCELADO'
]

# Componentes agendados
agendados_avalia = (
    ~df_avalia['SITUAÇÃO DA AVALIAÇÃO'].isin(situacoes_nao_agendadas)
).sum()

# perc_agend_avalia: componentes agendados / total de componentes
perc_agend_avalia = round((agendados_avalia / total_agend_avalia) * 100)

# Variável combinada para a apresentação
agend_avalia = f"{int(agendados_avalia)} / {int(total_agend_avalia)}"

# barra_agend_avalia
# Valor do percentual
percentual = perc_agend_avalia

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#1d9e75',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color='#1d9e75'
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'barra_agend_avalia.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# ESTUDANTE - AVALIA POTIGUAR
# 'estud_avalia' é composto por 'estudantes liberados' (liberad_avalia) e 'total de estudantes'
# Total de Estudantes
total_estud_avalia = df_avalia['MATRÍCULA SISTEMA'].nunique()

# Estudantes Liberados
# Estudante liberado (liberad_avalia) é aquele 'APROVADO' em todos seus componentes curriulares
estudantes_liberados_avalia = (
    df_avalia
    .groupby('MATRÍCULA SISTEMA')['SITUAÇÃO DA AVALIAÇÃO']
    .apply(lambda x: x.eq('APROVADO').all())
)

# Quantidade de estudantes liberados
liberad_avalia = estudantes_liberados_avalia.sum()


# perc_estud_avalia: estudantes liberados / total de estudantes
perc_estud_avalia = round((liberad_avalia / total_estud_avalia) * 100)

# Variável combinada para a apresentação
estud_avalia = f"{int(liberad_avalia)} / {int(total_estud_avalia)}"

# barra_estud_avalia
# Valor do percentual
percentual = perc_estud_avalia

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#000000',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#000000"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'barra_estud_avalia.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# COMPONENTES - AVALIA POTIGUAR
# 'comp_avalia' é composto por 'componentes aprovados' e 'total de componentes'

# Total de componentes
total_comp_avalia = df_avalia['COMPONENTE CURRICULAR'].count()

# Componentes aprovados
aprovados_comp_avalia = df_avalia[df_avalia['SITUAÇÃO DA AVALIAÇÃO'] == 'APROVADO']['COMPONENTE CURRICULAR'].count()

# perc_comp_avalia: componentes aprovados / total de componentes
perc_comp_avalia = round((aprovados_comp_avalia / total_comp_avalia) * 100)

# Variável combinada para a apresentação
comp_avalia = f"{int(aprovados_comp_avalia)} / {int(total_comp_avalia)}"

# barra_comp_avalia
# Valor do percentual
percentual = perc_comp_avalia

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#c47e3c',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#ffffff"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'barra_comp_avalia.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# graf_perc_agendamentos_direc_avalia: Gráfico de barras com percentual de componentes agendamentos por DIREC
# Situações que representam componentes não agendados
situacoes_nao_agendadas = [
    'NÃO AGENDADA',
    'REPROVADA',
    'CANCELADA'
]

# Identificar componentes agendados
df_agendamento_avalia = df_avalia.copy()

df_agendamento_avalia['AGENDADO'] = (
    ~df_agendamento_avalia['SITUAÇÃO DA AVALIAÇÃO']
    .isin(situacoes_nao_agendadas)
)

# Calcular componentes agendados e total por DIREC
df_grafico = (
    df_agendamento_avalia
    .groupby('DIREC')
    .agg(
        agendados=('AGENDADO', 'sum'),
        total=('AGENDADO', 'count')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['agendados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#cc8a42',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, agendados, total in zip(
    barras,
    df_grafico['agendados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(agendados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'grafico_agendamento_direc_avalia.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)


# graf_perc_liberados_direc_avalia: Gráfico de barras com percentual de estudantes liberados por DIREC
df_estudantes_liberados_avalia = (
    df_avalia
    .groupby('MATRÍCULA SISTEMA')
    .agg(
        DIREC=('DIREC', 'first'),
        LIBERADO=('SITUAÇÃO DA AVALIAÇÃO', lambda x: x.eq('APROVADO').all())
    )
    .reset_index()
)

# Calcular estudantes liberados e total por DIREC
df_grafico = (
    df_estudantes_liberados_avalia
    .groupby('DIREC')
    .agg(
        liberados=('LIBERADO', 'sum'),
        total=('MATRÍCULA SISTEMA', 'count')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['liberados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#58a75b',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, liberados, total in zip(
    barras,
    df_grafico['liberados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(liberados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'grafico_estudantes_liberados_direc_avalia.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

###########################################
# PLURALL
'''
Variáveis:
- perc_agend_plurall
- agend_plurall
- barra_agend_plurall
- perc_estud_plurall
- estud_plurall
- barra_estud_plurall
- perc_comp_plurall
- comp_plurall
- barra_comp_plurall
- graf_perc_agendamentos_direc_plurall
- liberad_plurall
- graf_perc_liberados_direc_plurall
'''

# Carregar base de dados do painel de monitoramento RAPP (feito no GAS):
df_plurall_agend = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\Monitoramento_RAPP_2026_2º Ciclo.xlsx', sheet_name='Agendamento')
df_plurall_prova = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\Monitoramento_RAPP_2026_2º Ciclo.xlsx', sheet_name='Dados_brutos', usecols=range(13))

# AGENDAMENTO - PLURALL
# agend_plurall é composto por 'turmas agendadas' e total previstos de agendamentos (agendadas + não agendadas)
'''
A turma é contabilizada de acordo com o ID TURMA
Uma turma tem vários componentes diferentes, ou seja, várias linhas.
A turma para ser considerada 'Agendada' deve ter todos os componentes agendados.
Para ser considerada 'Não Agendada' basta ter 1 componente não agendado.
'''
# Total de turmas de agendamento
total_agend_plurall = df_plurall_agend['ID TURMA'].nunique()

# Turmas agendadas
# Considerar uma turma agendada somente quando todos os registros (linhas) estiverem como AGENDOU = True
# Converter AGENDOU para booleano
df_plurall_agend['AGENDOU_BOOL'] = (
    df_plurall_agend['AGENDOU']
    .astype('string')
    .str.strip()
    .str.upper()
    .map({
        'VERDADEIRO': True,
        'FALSO': False,
        'TRUE': True,
        'FALSE': False
    })
)

# Identifica as turmas que tem todos os registros (linhas) AGENDADOS (True)
plurall_turmas_agendadas = (
    df_plurall_agend
    .groupby(['ID TURMA'])['AGENDOU_BOOL']
    .all()
)

# Quantidade de turmas agendadas
qtd_plurall_turmas_agendadas = plurall_turmas_agendadas.sum()

# perc_agend_plurall: turmas agendadas / total de turmas de agendamento
perc_agend_plurall = round((qtd_plurall_turmas_agendadas / total_agend_plurall) * 100)

# Variável combinada para a apresentação
agend_plurall = f"{int(qtd_plurall_turmas_agendadas)} / {int(total_agend_plurall)}"

# barra_agend_plurall
# Valor do percentual
percentual = perc_agend_plurall

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#1d9e75',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#1d9e75"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'barra_agend_plurall.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# ESTUDANTE - PLURALL
# estud_plurall é composto por 'estudantes liberados' (liberad_plurall) e 'total de estudantes'

# Total de estudantes
total_estud_plurall = df_plurall_prova['MATRÍCULA'].nunique()

# Estudante liberado (liberad_plurall) é aquele 'APROVADO' em todos seus componentes curriulares
estudantes_liberados_plurall = (
    df_plurall_prova
    .groupby('MATRÍCULA')['Situação']
    .apply(lambda x: ~x.isin(['Não Avaliado', 'Não Aprovado']).any())
)

# Quantidade de estudantes liberados
liberad_plurall = estudantes_liberados_plurall.sum()


# perc_estud_plurall: estudantes liberados / total de estudantes
perc_estud_plurall = round((liberad_plurall / total_estud_plurall) * 100)

# Variável combinada para a apresentação
estud_plurall = f"{int(liberad_plurall)} / {int(total_estud_plurall)}"


# barra_estud_plurall
# Valor do percentual
percentual = perc_estud_plurall

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#000000',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#000000"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'barra_estud_plurall.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# COMPONENTES - PLURALL
# 'comp_plurall' é composto por 'componentes aprovados' e 'total de componentes'

# Total de componentes
total_comp_plurall = df_plurall_prova['COMPONENTE CURRICULAR'].count()

# Componentes aprovados
aprovados_comp_plurall = df_plurall_prova[df_plurall_prova['Situação'] == 'Aprovado']['COMPONENTE CURRICULAR'].count()

# perc_comp_plurall: componentes aprovados / total de componentes
perc_comp_plurall = round((aprovados_comp_plurall / total_comp_plurall) * 100)

# Variável combinada para a apresentação
comp_plurall = f"{int(aprovados_comp_plurall)} / {int(total_comp_plurall)}"

# barra_comp_plurall
# Valor do percentual
percentual = perc_comp_plurall

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#c47e3c',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#ffffff"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'barra_comp_plurall.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# graf_perc_agendamentos_direc_plurall: Gráfico de barras com percentual de agendamento por DIREC
# Converter AGENDOU para booleano
df_plurall_agend['AGENDOU_BOOL'] = (
    df_plurall_agend['AGENDOU']
    .astype('string')
    .str.strip()
    .str.upper()
    .map({'VERDADEIRO': True, 'FALSO': False, 'TRUE': True, 'FALSE': False})
)

# Calcular agendamentos por DIREC
df_grafico = (
    df_plurall_agend
    .groupby('DIREC')
    .agg(
        total=('AGENDOU_BOOL', 'size'),
        agendados=('AGENDOU_BOOL', 'sum')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['agendados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#cc8a42',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, agendados, total in zip(
    barras,
    df_grafico['agendados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(agendados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'grafico_agendamento_direc_plurall.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)


# graf_perc_liberados_direc_plurall: Gráfico de barras com percentual de estudantes liberados por DIREC
# Identificar estudantes liberados e recuperar a DIREC
df_estudantes_liberados = (
    df_plurall_prova
    .groupby('MATRÍCULA')
    .agg(
        DIREC=('DIREC', 'first'),
        LIBERADO=('Situação', lambda x: ~x.isin(
            ['Não Avaliado', 'Não Aprovado']
        ).any())
    )
    .reset_index()
)

# Calcular estudantes liberados e total por DIREC
df_grafico = (
    df_estudantes_liberados
    .groupby('DIREC')
    .agg(
        liberados=('LIBERADO', 'sum'),
        total=('MATRÍCULA', 'count')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['liberados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#58a75b',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, liberados, total in zip(
    barras,
    df_grafico['liberados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(liberados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'grafico_estudantes_liberados_direc_plurall.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)






























































###################################################################################################################
'''
PLURALL: 1º + 2º ciclos
Dados e gráficos para a Avaliação Plurall, na base que considera também os estudantes do 1º ciclo

Variáveis para serem criadas para a apresentação do GT:
- t_perc_agend_t
- t-agend_t
- t_barra_agend_t
- t_perc_estud_t
- t_estud_t
- t_barra_estud_t
- t_perc_comp_t
- t_comp_t
- t_barra_comp_t
- t_graf_perc_agendamentos_direc_t
- t_liberad_t
- t_graf_perc_liberados_direc_t

'''
# Carregar base de dados do painel de monitoramento RAPP (feito no GAS):
df_plurall_agend = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\Monitoramento_RAPP_2026_2º Ciclo.xlsx', sheet_name='Agendamento')
df_plurall_prova = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\Monitoramento_RAPP_2026_2º Ciclo.xlsx', sheet_name='Dados_brutos', usecols=range(13))

# Dados do 1º ciclo
df_plurall_prova_1ciclo = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\1ciclo_Monitoramento_RAPP_2026.xlsx', sheet_name='Dados_brutos', usecols=range(12))


# Selecionar somente as linhas com 'Situação' = 'Aprovado'
df_plurall_prova_1ciclo = df_plurall_prova_1ciclo[df_plurall_prova_1ciclo['Situação'] == 'Aprovado']

# Adicionar as linhas de 'Situação'= 'Aprovado' do 1º ciclo na base do 2º ciclo
df_plurall_prova = pd.concat([df_plurall_prova, df_plurall_prova_1ciclo], ignore_index=True)


# AGENDAMENTO - PLURALL
# t_agend_t é composto por 'turmas agendadas' e total previstos de agendamentos (agendadas + não agendadas)
'''
A turma é contabilizada de acordo com o ID TURMA
Uma turma tem vários componentes diferentes, ou seja, várias linhas.
A turma para ser considerada 'Agendada' deve ter todos os componentes agendados.
Para ser considerada 'Não Agendada' basta ter 1 componente não agendado.
'''
# Total de turmas de agendamento
total_agend_plurall = df_plurall_agend['ID TURMA'].nunique()

# Turmas agendadas
# Considerar uma turma agendada somente quando todos os registros (linhas) estiverem como AGENDOU = True
# Converter AGENDOU para booleano
df_plurall_agend['AGENDOU_BOOL'] = (
    df_plurall_agend['AGENDOU']
    .astype('string')
    .str.strip()
    .str.upper()
    .map({
        'VERDADEIRO': True,
        'FALSO': False,
        'TRUE': True,
        'FALSE': False
    })
)

# Identifica as turmas que tem todos os registros (linhas) AGENDADOS (True)
plurall_turmas_agendadas = (
    df_plurall_agend
    .groupby(['ID TURMA'])['AGENDOU_BOOL']
    .all()
)

# Quantidade de turmas agendadas
qtd_plurall_turmas_agendadas = plurall_turmas_agendadas.sum()

# perc_agend_plurall: turmas agendadas / total de turmas de agendamento
t_perc_agend_t = round((qtd_plurall_turmas_agendadas / total_agend_plurall) * 100)

# Variável combinada para a apresentação
t_agend_t = f"{int(qtd_plurall_turmas_agendadas)} / {int(total_agend_plurall)}"

# barra_agend_plurall
# Valor do percentual
percentual = t_perc_agend_t

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#1d9e75',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#1d9e75"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 't_barra_agend_t.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# ESTUDANTE - PLURALL
# t_estud_t é composto por 'estudantes liberados' (liberad_plurall) e 'total de estudantes'

# Total de estudantes
total_estud_t = df_plurall_prova['MATRÍCULA'].nunique()

# Estudante liberado (liberad_t) é aquele 'APROVADO' em todos seus componentes curriulares
estudantes_liberados_t = (
    df_plurall_prova
    .groupby('MATRÍCULA')['Situação']
    .apply(lambda x: ~x.isin(['Não Avaliado', 'Não Aprovado']).any())
)

# Quantidade de estudantes liberados
t_liberad_t = estudantes_liberados_t.sum()


# t_perc_estud_t: estudantes liberados / total de estudantes
t_perc_estud_t = round((t_liberad_t / total_estud_t) * 100)

# Variável combinada para a apresentação
t_estud_t = f"{int(t_liberad_t)} / {int(total_estud_t)}"


# barra_estud_plurall
# Valor do percentual
percentual = t_perc_estud_t

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#000000',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#000000"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 't_barra_estud_t.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# COMPONENTES - PLURALL
# 't_comp_t' é composto por 'componentes aprovados' e 'total de componentes'

# Total de componentes
total_comp_t = df_plurall_prova['COMPONENTE CURRICULAR'].count()

# Componentes aprovados
aprovados_comp_t = df_plurall_prova[df_plurall_prova['Situação'] == 'Aprovado']['COMPONENTE CURRICULAR'].count()

# t_perc_comp_t: componentes aprovados / total de componentes
t_perc_comp_t = round((aprovados_comp_t / total_comp_t) * 100)

# Variável combinada para a apresentação
t_comp_t = f"{int(aprovados_comp_t)} / {int(total_comp_t)}"

# barra_comp_t
# Valor do percentual
percentual = t_perc_comp_t

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#c47e3c',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#ffffff"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 't_barra_comp_t.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# t_graf_perc_agendamentos_direc_t: Gráfico de barras com percentual de agendamento por DIREC
# Converter AGENDOU para booleano
df_plurall_agend['AGENDOU_BOOL'] = (
    df_plurall_agend['AGENDOU']
    .astype('string')
    .str.strip()
    .str.upper()
    .map({'VERDADEIRO': True, 'FALSO': False, 'TRUE': True, 'FALSE': False})
)

# Calcular agendamentos por DIREC
df_grafico = (
    df_plurall_agend
    .groupby('DIREC')
    .agg(
        total=('AGENDOU_BOOL', 'size'),
        agendados=('AGENDOU_BOOL', 'sum')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['agendados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#cc8a42',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, agendados, total in zip(
    barras,
    df_grafico['agendados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(agendados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 't_grafico_agendamento_direc_t.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)


# t_graf_perc_liberados_direc_t: Gráfico de barras com percentual de estudantes liberados por DIREC
# Identificar estudantes liberados e recuperar a DIREC
df_estudantes_liberados = (
    df_plurall_prova
    .groupby('MATRÍCULA')
    .agg(
        DIREC=('DIREC', 'first'),
        LIBERADO=('Situação', lambda x: ~x.isin(
            ['Não Avaliado', 'Não Aprovado']
        ).any())
    )
    .reset_index()
)

# Calcular estudantes liberados e total por DIREC
df_grafico = (
    df_estudantes_liberados
    .groupby('DIREC')
    .agg(
        liberados=('LIBERADO', 'sum'),
        total=('MATRÍCULA', 'count')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['liberados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#58a75b',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, liberados, total in zip(
    barras,
    df_grafico['liberados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(liberados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 't_grafico_estudantes_liberados_direc_t.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)


























































###################################################################################################################
'''
AVALIA + PLURALL
Dados e gráficos para a mesclagem da Avaliação Plurall e Avalia Potiguar. 
A base da Plurall considera os estudantes do 1º e 2º ciclos.

Variáveis para serem criadas para a apresentação do GT:
- pa_perc_agend_plurall
- pa_agend_plurall
- pa_barra_agend_plurall
- pa_perc_estud_plurall
- pa_estud_plurall
- pa_barra_estud_plurall
- pa_perc_comp_plurall
- pa_comp_plurall
- pa_barra_comp_plurall
- pa_graf_perc_agendamentos_direc_plurall
- pa_liberad_plurall
- pa_graf_perc_liberados_direc_plurall
'''

# Carregar base de dados do painel de monitoramento RAPP (feito no GAS):
df_plurall_agend = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\Monitoramento_RAPP_2026_2º Ciclo.xlsx', sheet_name='Agendamento')
df_plurall_prova = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\Monitoramento_RAPP_2026_2º Ciclo.xlsx', sheet_name='Dados_brutos', usecols=range(13))

# Dados do 1º ciclo
df_plurall_prova_1ciclo = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\1ciclo_Monitoramento_RAPP_2026.xlsx', sheet_name='Dados_brutos', usecols=range(12))


# Selecionar somente as linhas com 'Situação' = 'Aprovado'
df_plurall_prova_1ciclo = df_plurall_prova_1ciclo[df_plurall_prova_1ciclo['Situação'] == 'Aprovado']

# Adicionar as linhas de 'Situação'= 'Aprovado' do 1º ciclo na base do 2º ciclo
df_plurall_prova = pd.concat([df_plurall_prova, df_plurall_prova_1ciclo], ignore_index=True)


# Carregar base de dados do Avalia Potiguar:
df_avalia = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\relatorio_acompanhamento_rapp_ept_div.xlsx', sheet_name='Relatório RAPP')


# Criar o dataframe combinado para agendamento: df_pa_agend

# Preparar o df_plurall_agend
df_plurall_pa = df_plurall_agend[[
    'DIREC',
    'ESCOLA',
    'ID ESCOLA',
    'TURMA',
    'ID TURMA',
    'AGENDOU',
    'COMPONENTE',
    'SERIE',
    'ETAPA_RESUMIDA'
]].copy()

df_plurall_pa = df_plurall_pa.rename(columns={
    'ID ESCOLA': 'ID ESCOLA - INEP ESCOLA',
    'TURMA': 'TURMA - ESTUDANTE',
    'ID TURMA': 'ID TURMA - MATRÍCULA SISTEMA',
    'AGENDOU': 'AGENDOU - SITUAÇÃO DA AVALIAÇÃO',
    'ETAPA_RESUMIDA': 'ETAPA'
})


# Preparar o df_avalia
df_avalia_pa = df_avalia[[
    'DIREC',
    'ESCOLA',
    'INEP ESCOLA',
    'NOME DO ESTUDANTE',
    'MATRÍCULA SISTEMA',
    'SITUAÇÃO DA AVALIAÇÃO',
    'COMPONENTE CURRICULAR',
    'SÉRIE',
    'ETAPA'
]].copy()

df_avalia_pa = df_avalia_pa.rename(columns={
    'INEP ESCOLA': 'ID ESCOLA - INEP ESCOLA',
    'NOME DO ESTUDANTE': 'TURMA - ESTUDANTE',
    'MATRÍCULA SISTEMA': 'ID TURMA - MATRÍCULA SISTEMA',
    'SITUAÇÃO DA AVALIAÇÃO': 'AGENDOU - SITUAÇÃO DA AVALIAÇÃO',
    'COMPONENTE CURRICULAR': 'COMPONENTE',
    'SÉRIE': 'SERIE'
})

# Concatenar os dataframes
df_pa_agend = pd.concat(
    [df_plurall_pa, df_avalia_pa],
    ignore_index=True)


# Criar o dataframe combinado para prova: df_pa_prova

# Preparar o df_plurall_prova
df_plurall_pa = df_plurall_prova[[
    'NOME',
    'MATRÍCULA',
    'DIREC',
    'ESCOLA',
    'INEP ESCOLA',
    'ETAPA_RESUMIDA',
    'SÉRIE',
    'COMPONENTE CURRICULAR',
    'Situação',
    'rendimento (%)'
]].copy()

df_plurall_pa = df_plurall_pa.rename(columns={
    'NOME': 'ESTUDANTE',
    'ETAPA_RESUMIDA': 'ETAPA',
    'Situação': 'SITUAÇÃO',
    'rendimento (%)': 'RENDIMENTO'
})


# Preparar o df_avalia
df_avalia_pa = df_avalia[[
    'NOME DO ESTUDANTE',
    'MATRÍCULA SISTEMA',
    'DIREC',
    'ESCOLA',
    'INEP ESCOLA',
    'ETAPA',
    'SÉRIE',
    'COMPONENTE CURRICULAR',
    'SITUAÇÃO DA AVALIAÇÃO',
    'NOTA FINAL'
]].copy()

df_avalia_pa = df_avalia_pa.rename(columns={
    'NOME DO ESTUDANTE': 'ESTUDANTE',
    'MATRÍCULA SISTEMA': 'MATRÍCULA',
    'SITUAÇÃO DA AVALIAÇÃO': 'SITUAÇÃO',
    'NOTA FINAL': 'RENDIMENTO'
})

# Concatenar os dataframes
df_pa_prova = pd.concat(
    [df_plurall_pa, df_avalia_pa],
    ignore_index=True
)


# Padronizar algumas colunas do df_pa_prova
# Séries
df_pa_prova['SÉRIE'] = df_pa_prova['SÉRIE'].replace({
    '6º Ano': '6º ANO',
    '7º Ano': '7º ANO',
    '8º Ano': '8º ANO',
    '9º Ano': '9º ANO',
    '1ª Série': '1ª SÉRIE',
    '2ª Série': '2ª SÉRIE',
    '3ª Série': '3ª SÉRIE'
})


# Etapa de Ensino
# Criar a coluna 'ETAPA_RESUMIDO' com base na coluna 'SÉRIE'
mapeamento_etapa = {
    '1ª SÉRIE': 'Ensino Médio',
    '2ª SÉRIE': 'Ensino Médio',
    '3ª SÉRIE': 'Ensino Médio',
    '6º ANO': 'Ens. Fund. - Anos Finais',
    '7º ANO': 'Ens. Fund. - Anos Finais',
    '8º ANO': 'Ens. Fund. - Anos Finais',
    '9º ANO': 'Ens. Fund. - Anos Finais',
    '3º Período (3ª Série)': 'Ensino Médio',
    '2º Período (2ª Série)': 'Ensino Médio',
    '1º Período (1ª Série)': 'Ensino Médio',
    '5º PERÍODO (8° E 9° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    '4º PERÍODO (6° E 7° ANO - ANOS FINAIS)': 'Ens. Fund. - Anos Finais',
    'TURMA II (8° E 9° ANOS)': 'Ens. Fund. - Anos Finais',
    '1º MÓDULO': 'Ensino Médio',
    '2º MÓDULO': 'Ensino Médio',
    '2° MÓDULO': 'Ensino Médio',
    '2º MODULO': 'Ensino Médio',
    '3º MÓDULO': 'Ensino Médio',
    '4º MÓDULO': 'Ensino Médio',
    '2º SEMESTRE': 'Ensino Médio',
    '1º SEMESTRE': 'Ensino Médio',
    '3º SEMESTRE': 'Ensino Médio',
    '3° PERÍODO': 'Ensino Médio',
    'BLOCO B': 'Ensino Médio',
    'BLOCO A': 'Ensino Médio',
    'BLOCO C': 'Ensino Médio',
    'BLOCO D': 'Ensino Médio',
    '4° PERÍODO': 'Ensino Médio'
}

df_pa_prova['ETAPA_RESUMIDA'] = df_pa_prova['SÉRIE'].map(mapeamento_etapa)


# Situação
df_pa_prova['SITUAÇÃO'] = df_pa_prova['SITUAÇÃO'].replace(
    'APROVADO', 'Aprovado'
)


# Exclusão de duplicatas de 'MATRÍCULA' e 'COMPONENTE CURRICULAR' dando preferência para manutenção de 'SITUAÇÃO' = 'Aprovado'
# Verificar se tem valores duplicados nas colunas 'MATRÍCULA' e 'COMPONENTE CURRICULAR'
qtd_duplicatas = df_pa_prova.duplicated(
    subset=['MATRÍCULA', 'COMPONENTE CURRICULAR']
).sum()

print(f'Quantidade de duplicatas: {qtd_duplicatas}')


# Priorizar registros com SITUAÇÃO = 'Aprovado'
df_pa_prova = df_pa_prova.sort_values(
    by='SITUAÇÃO',
    key=lambda x: x.eq('Aprovado'),
    ascending=False
)

# Excluir duplicatas, mantendo a primeira ocorrência
df_pa_prova = df_pa_prova.drop_duplicates(
    subset=['MATRÍCULA', 'COMPONENTE CURRICULAR'],
    keep='first'
).reset_index(drop=True)


# AGENDAMENTO - PLURALL
# pa_agend_plurall é composto por 'turmas/componentes agendadas' e total previstos de agendamentos (agendadas + não agendadas)
'''
A unidade de agendamento é contabilizada de acordo com o 'ID TURMA - MATRÍCULA SISTEMA'
Uma ID tem vários componentes diferentes, ou seja, várias linhas.
A ID para ser considerada 'Agendada' deve ter todas as linhas agendados.
Para ser considerada 'Não Agendada' basta ter 1 linha não agendado.
'''

# Total de ID de agendamento
total_agend_pa = df_pa_agend['ID TURMA - MATRÍCULA SISTEMA'].nunique()


# Turmas agendadas
# Considerar uma turma agendada somente quando todos os registros (linhas) estiverem como AGENDOU = True

# Trocar valores de Agendamento para deixar como Agendou ou não
# Não agendados é composto por: 'SITUAÇÃO DA AVALIAÇÃO' = 'NÃO AGENDADA', 'REPROVADA' e 'CANCELADA'. Restante é agendado.
df_pa_agend['AGENDOU - SITUAÇÃO DA AVALIAÇÃO'] = df_pa_agend['AGENDOU - SITUAÇÃO DA AVALIAÇÃO'].replace({
    'APROVADO': 'VERDADEIRO',
    'APROVADA': 'VERDADEIRO',
    'AGUARDANDO_INICIO': 'VERDADEIRO',
    'NÃO AGENDADA': 'FALSO',
    'NÃO AGENDADO': 'FALSO',
    'AGENDADA': 'VERDADEIRO',
    'AGENDADO': 'VERDADEIRO',
    'REAGENDADA': 'VERDADEIRO',
    'REAGENDADO': 'VERDADEIRO',
    'REPROVADA': 'FALSO',
    'REPROVADO': 'FALSO',
    'CANCELADA': 'FALSO',
    'CANCELADO': 'FALSO',
    'INICIADA': 'VERDADEIRO',
    'INICIADO': 'VERDADEIRO'
})


# Converter AGENDOU para booleano
df_pa_agend['AGENDOU_BOOL'] = (
    df_pa_agend['AGENDOU - SITUAÇÃO DA AVALIAÇÃO']
    .astype('string')
    .str.strip()
    .str.upper()
    .map({
        'VERDADEIRO': True,
        'FALSO': False,
        'TRUE': True,
        'FALSE': False
    })
)

# Identifica as turmas que tem todos os registros (linhas) AGENDADOS (True)
pa_turmas_agendadas = (
    df_pa_agend
    .groupby(['ID TURMA - MATRÍCULA SISTEMA'])['AGENDOU_BOOL']
    .all()
)

# Quantidade de turmas agendadas
qtd_pa_turmas_agendadas = pa_turmas_agendadas.sum()

# perc_agend_plurall: turmas agendadas / total de turmas de agendamento
pa_perc_agend_pa = round((qtd_pa_turmas_agendadas / total_agend_pa) * 100)

# Variável combinada para a apresentação
pa_agend_pa = f"{int(qtd_pa_turmas_agendadas)} / {int(total_agend_pa)}"

# barra_agend_plurall
# Valor do percentual
percentual = pa_perc_agend_pa

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#1d9e75',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#1d9e75"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'pa_barra_agend_pa.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# ESTUDANTE - PLURALL
# pa_estud_pa é composto por 'estudantes liberados' (pa_liberad_pa) e 'total de estudantes'

# Total de estudantes
total_estud_pa = df_pa_prova['MATRÍCULA'].nunique()

# Estudante liberado (pa_liberad_plurall) é aquele 'APROVADO' em todos seus componentes curriculares
estudantes_liberados_pa = (
    df_pa_prova
    .groupby('MATRÍCULA')['SITUAÇÃO']
    .apply(lambda x: x.eq('Aprovado').all())
)

# Quantidade de estudantes liberados
pa_liberad_pa = estudantes_liberados_pa.sum()


# pa_perc_estud_pa: estudantes liberados / total de estudantes
pa_perc_estud_pa = round((pa_liberad_pa / total_estud_pa) * 100)

# Variável combinada para a apresentação
pa_estud_pa = f"{int(pa_liberad_pa)} / {int(total_estud_pa)}"


# barra_estud_pa
# Valor do percentual
percentual = pa_perc_estud_pa

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#000000',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#000000"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'pa_barra_estud_pa.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# COMPONENTES - PLURALL
# 'pa_comp_pa' é composto por 'componentes aprovados' e 'total de componentes'

# Total de componentes
total_comp_pa = df_pa_prova['COMPONENTE CURRICULAR'].count()

# Componentes aprovados
aprovados_comp_pa = df_pa_prova[df_pa_prova['SITUAÇÃO'] == 'Aprovado']['COMPONENTE CURRICULAR'].count()

# pa_perc_comp_pa: componentes aprovados / total de componentes
pa_perc_comp_pa = round((aprovados_comp_pa / total_comp_pa) * 100)

# Variável combinada para a apresentação
pa_comp_pa = f"{int(aprovados_comp_pa)} / {int(total_comp_pa)}"

# barra_comp_pa
# Valor do percentual
percentual = pa_perc_comp_pa

# Criar figura pequena
fig, ax = plt.subplots(figsize=(2.2, 0.35))

# Ocultar eixos
ax.set_xlim(0, 125)
ax.set_ylim(0, 1)
ax.axis('off')

# Dimensões da barra
altura = 0.55
pos_y = 0.22

# Barra de fundo (100%)
fundo = FancyBboxPatch(
    (0, pos_y),
    100,
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#eeece1',
    edgecolor='none'
)

ax.add_patch(fundo)

# Barra preenchida (percentual realizado)
preenchimento = FancyBboxPatch(
    (0, pos_y),
    max(0, min(percentual, 100)),
    altura,
    boxstyle=f"round,pad=0,rounding_size={altura / 2}",
    facecolor='#c47e3c',
    edgecolor='none'
)

ax.add_patch(preenchimento)

# Percentual à direita da barra
ax.text(
    104,
    0.5,
    f'{percentual:.0f}%',
    ha='left',
    va='center',
    fontsize=12,
    fontweight='bold',
    color="#ffffff"
)

# Fundo transparente
fig.patch.set_alpha(0)

# Remover margens
plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'pa_barra_comp_pa.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar imagem
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight',
    pad_inches=0.01
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)

# pa_graf_perc_agendamentos_direc_pa: Gráfico de barras com percentual de agendamento por DIREC
# Converter AGENDOU para booleano
df_pa_agend['AGENDOU_BOOL'] = (
    df_pa_agend['AGENDOU - SITUAÇÃO DA AVALIAÇÃO']
    .astype('string')
    .str.strip()
    .str.upper()
    .map({'VERDADEIRO': True, 'FALSO': False, 'TRUE': True, 'FALSE': False})
)

# Calcular agendamentos por DIREC
df_grafico = (
    df_pa_agend
    .groupby('DIREC')
    .agg(
        total=('AGENDOU_BOOL', 'size'),
        agendados=('AGENDOU_BOOL', 'sum')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['agendados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#cc8a42',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, agendados, total in zip(
    barras,
    df_grafico['agendados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(agendados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'pa_grafico_agendamento_direc_pa.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)


# pa_graf_perc_liberados_direc_pa: Gráfico de barras com percentual de estudantes liberados por DIREC
# Identificar estudantes liberados e recuperar a DIREC
df_estudantes_liberados = (
    df_pa_prova
    .groupby('MATRÍCULA')
    .agg(
        DIREC=('DIREC', 'first'),
        LIBERADO=('SITUAÇÃO', lambda x: x.eq('Aprovado').any())
    )
    .reset_index()
)

# Calcular estudantes liberados e total por DIREC
df_grafico = (
    df_estudantes_liberados
    .groupby('DIREC')
    .agg(
        liberados=('LIBERADO', 'sum'),
        total=('MATRÍCULA', 'count')
    )
    .reset_index()
)

# Calcular percentual
df_grafico['PERCENTUAL'] = (
    df_grafico['liberados'] / df_grafico['total'] * 100
)

# Ordenar DIRECs
df_grafico = df_grafico.sort_values('DIREC')

# Limitar nomes das DIRECs aos primeiros 9 caracteres
nomes_direc = df_grafico['DIREC'].astype(str).str[:9]

# Criar gráfico
fig, ax = plt.subplots(figsize=(14, 7))

barras = ax.bar(
    nomes_direc,
    df_grafico['PERCENTUAL'],
    color='#58a75b',
    width=0.7
)

# Configurar eixo Y
ax.set_ylim(0, 100)
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=100))

# Configurar tamanho das fontes dos eixos
ax.tick_params(axis='x', labelsize=12, rotation=45)
ax.tick_params(axis='y', labelsize=13)

# Alinhar nomes das DIRECs à direita
plt.setp(ax.get_xticklabels(), ha='right')

# Percentuais no topo das barras
ax.bar_label(
    barras,
    labels=[f'{valor:.0f}%' for valor in df_grafico['PERCENTUAL']],
    padding=4,
    fontsize=13,
    fontweight='bold'
)

# Quantidades absolutas dentro das barras
for barra, liberados, total in zip(
    barras,
    df_grafico['liberados'],
    df_grafico['total']
):
    ax.text(
        barra.get_x() + barra.get_width() / 2,
        barra.get_height() / 2,
        f'{int(liberados)} / {int(total)}',
        ha='center',
        va='center',
        rotation=90,
        fontsize=11,
        fontweight='bold',
        color='white'
    )

# Remover títulos dos eixos
ax.set_xlabel('')
ax.set_ylabel('')

# Remover bordas superior e direita
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

# Fundo transparente
fig.patch.set_alpha(0)
ax.patch.set_alpha(0)

# Ajustar layout
plt.tight_layout()

# Definir o caminho exato do arquivo
caminho_grafico = os.path.join(pasta_graficos, 'pa_grafico_estudantes_liberados_direc_pa.png')

# Forçar a remoção do arquivo antigo (se existir) para evitar bloqueios do sistema operacional
if os.path.exists(caminho_grafico):
    os.remove(caminho_grafico)

# Salvar PNG
plt.savefig(
    caminho_grafico,
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()

# Limpar a memória fechando a figura atual
plt.close(fig)































###################################################
# Respostas para o MP:
# Utilizar o dado combinado de Avalia Potiguar, Plurall 1º ciclo + Plurall 2º ciclo
'''
1. Quantidade de estudantes que ingressaram em RAPP por etapa
2. Quantidade de estudantes atualmente em RAPP por etapa
3. Quantidade de estudantes aprovados em RAPP por etapa
4. Distribuição de pendências por etapa
'''

# 1. Quantidade de estudantes que ingressaram em RAPP por etapa
# Quantidade 'MATRÍCULA' distinta por 'ETAPA_RESUMIDA'
qtd_estudantes_etapa = (
    df_pa_prova
    .groupby('ETAPA_RESUMIDA')['MATRÍCULA']
    .nunique()
    .reset_index(name='QUANTIDADE_ESTUDANTES')
)

qtd_estudantes_etapa

# 2. Quantidade de estudantes atualmente em RAPP por etapa
# Identificar matrículas aprovadas em todos os componentes
matriculas_aprovadas = (
    df_pa_prova
    .groupby('MATRÍCULA')['SITUAÇÃO']
    .apply(lambda x: (x == 'Aprovado').all())
)

matriculas_aprovadas = matriculas_aprovadas[
    matriculas_aprovadas
].index

# Excluir estudantes aprovados em todas as linhas
df_pa_rapp_atual = df_pa_prova[
    ~df_pa_prova['MATRÍCULA'].isin(matriculas_aprovadas)
].copy()

# Quantidade de estudantes atualmente em RAPP por etapa
qtd_estudantes_rapp_etapa = (
    df_pa_rapp_atual
    .groupby('ETAPA_RESUMIDA')['MATRÍCULA']
    .nunique()
    .reset_index(name='QTD_ESTUDANTES')
)

qtd_estudantes_rapp_etapa

# 3. Quantidade de estudantes aprovados em RAPP por etapa
# Filtrar estudantes aprovados em todos os componentes
df_pa_aprovados = df_pa_prova[
    df_pa_prova['MATRÍCULA'].isin(matriculas_aprovadas)
].copy()

# Quantidade de estudantes aprovados por etapa
qtd_estudantes_aprovados_etapa = (
    df_pa_aprovados
    .groupby('ETAPA_RESUMIDA')['MATRÍCULA']
    .nunique()
    .reset_index(name='QTD_ESTUDANTES')
)

qtd_estudantes_aprovados_etapa

# 4. Distribuição de pendências por etapa
# Quantidade de estudantes únicos atualmente em RAPP
# Quantidade de componentes por estudante
df_componentes_estudante = (
    df_pa_prova
    .groupby(['ETAPA_RESUMIDA', 'MATRÍCULA'])
    .size()
    .reset_index(name='QTD_COMPONENTES')
)

# Distribuição da quantidade de componentes por etapa
dist_componentes_etapa = (
    df_componentes_estudante
    .groupby(['ETAPA_RESUMIDA', 'QTD_COMPONENTES'])
    .size()
    .reset_index(name='QTD_ESTUDANTES')
)

# Total de estudantes por etapa
dist_componentes_etapa['TOTAL_ETAPA'] = (
    dist_componentes_etapa
    .groupby('ETAPA_RESUMIDA')['QTD_ESTUDANTES']
    .transform('sum')
)

# Percentual perante o total de estudantes
dist_componentes_etapa['PERCENTUAL'] = (
    dist_componentes_etapa['QTD_ESTUDANTES'] /
    dist_componentes_etapa['TOTAL_ETAPA'] * 100
).round(1)

dist_componentes_etapa



































##################################################################################################################################
# Montar apresentação para GT
# Salvar as variáveis para usar na apresentação do GT
# Planilha de dados com os valores das variáveis para apresentação
# Dicionário reunindo o nome da variável e seu valor final
dados_variaveis = {
    'data': data_hoje,
    
    # Avalia Potiguar
    'perc_agend_avalia': perc_agend_avalia,
    'agend_avalia': agend_avalia,
    'perc_estud_avalia': perc_estud_avalia,
    'estud_avalia': estud_avalia,
    'perc_comp_avalia': perc_comp_avalia,
    'comp_avalia': comp_avalia,
    'liberad_avalia': liberad_avalia,
    
    # Plurall
    'perc_agend_plurall': perc_agend_plurall,
    'agend_plurall': agend_plurall,
    'perc_estud_plurall': perc_estud_plurall,
    'estud_plurall': estud_plurall,
    'perc_comp_plurall': perc_comp_plurall,
    'comp_plurall': comp_plurall,
    'liberad_plurall': liberad_plurall,

    # Plurall: 1º + 2º ciclo
    't_perc_agend_t': t_perc_agend_t,
    't_agend_t': t_agend_t,
    't_perc_estud_t': t_perc_estud_t,
    't_estud_t': t_estud_t,
    't_perc_comp_t': t_perc_comp_t,
    't_comp_t': t_comp_t,
    't_liberad_t': t_liberad_t,

    # Avalia + Plurall
    'pa_perc_agend_pa': pa_perc_agend_pa,
    'pa_agend_pa': pa_agend_pa,
    'pa_perc_estud_pa': pa_perc_estud_pa,
    'pa_estud_pa': pa_estud_pa,
    'pa_perc_comp_pa': pa_perc_comp_pa,
    'pa_comp_pa': pa_comp_pa,
    'pa_liberad_pa': pa_liberad_pa

}

# Conversão do dicionário para um DataFrame do Pandas com 2 colunas
df_exportacao = pd.DataFrame(list(dados_variaveis.items()), columns=['Variável', 'Valor'])

# Caminho do arquivo Excel
caminho_planilha = os.path.join(pasta_graficos, 'variaveis_calculadas.xlsx')

# Salvando o DataFrame em um arquivo Excel na mesma pasta dos gráficos
df_exportacao.to_excel(caminho_planilha, index=False)

with pd.ExcelWriter(caminho_planilha, engine='openpyxl') as writer:
    df_exportacao.to_excel(writer, sheet_name='Variáveis', index=False)
    df_pa_prova.to_excel(writer, sheet_name='Prova - Plurall e Avalia', index=False)
    df_pa_agend.to_excel(writer, sheet_name='Agendamento - Plurall e Avalia', index=False)
    qtd_estudantes_etapa.to_excel(writer, sheet_name='Ingressaram por Etapa', index=False)
    qtd_estudantes_rapp_etapa.to_excel(writer, sheet_name='Atualmente em RAPP', index=False)
    qtd_estudantes_aprovados_etapa.to_excel(writer, sheet_name='Aprovados por Etapa', index=False)
    dist_componentes_etapa.to_excel(writer, sheet_name='Distribuição de Pendências', index=False)

















##########################################################################################################################
# Preencher o PPT
# 1. Caminhos dos arquivos
caminho_template = r"D:\Scripts_Python\FGV\RAPP_apresentacao_GT\template_GARN_Reuniao de alinhamento RAPP.pptx"
caminho_saida = r"D:\Scripts_Python\FGV\RAPP_apresentacao_GT\apresentacoes_geradas\GARN_Reuniao de alinhamento RAPP.pptx"

# 2. Carregar a apresentação base
prs = Presentation(caminho_template)

# 3. Dicionários de Substituição
# Mapeamento de TEXTOS
# A chave (lado esquerdo) é exatamente o que está escrito na caixa de texto do PowerPoint
substituicoes_texto = {
    'data_hoje': f"{data_hoje}",

    'perc_agend_avalia': f"{perc_agend_avalia}%",
    'agend_avalia': agend_avalia,
    'perc_estud_avalia': f"{perc_estud_avalia}%",
    'estud_avalia': estud_avalia,
    'liberad_avalia': liberad_avalia,
    'perc_comp_avalia': f"{perc_comp_avalia}%",
    'comp_avalia': comp_avalia,

    'perc_agend_plurall': f"{perc_agend_plurall}%",
    'agend_plurall': agend_plurall,
    'perc_estud_plurall': f"{perc_estud_plurall}%",
    'estud_plurall': estud_plurall,
    'liberad_plurall': liberad_plurall,
    'perc_comp_plurall': f"{perc_comp_plurall}%",
    'comp_plurall': comp_plurall,

    't_perc_agend_t': f"{t_perc_agend_t}%",
    't_agend_t': t_agend_t,
    't_perc_estud_t': f"{t_perc_estud_t}%",
    't_estud_t': t_estud_t,
    't_liberad_t': t_liberad_t,
    't_perc_comp_t': f"{t_perc_comp_t}%",
    't_comp_t': t_comp_t,

    'pa_perc_agend_pa': f"{pa_perc_agend_pa}%",
    'pa_agend_pa': pa_agend_pa,
    'pa_perc_estud_pa': f"{pa_perc_estud_pa}%",
    'pa_estud_pa': pa_estud_pa,
    'pa_liberad_pa': pa_liberad_pa,
    'pa_perc_comp_pa': f"{pa_perc_comp_pa}%",
    'pa_comp_pa': pa_comp_pa,
}

# Mapeamento de IMAGENS (Gráficos)
# A chave é o que está escrito no PowerPoint; o valor é o caminho do PNG gerado.
substituicoes_imagem = {
    'barra_agend_avalia': os.path.join(pasta_graficos, 'barra_agend_avalia.png'),
    'barra_estud_avalia': os.path.join(pasta_graficos, 'barra_estud_avalia.png'),
    'barra_comp_avalia': os.path.join(pasta_graficos, 'barra_comp_avalia.png'),
    'graf_perc_agendamentos_direc_avalia': os.path.join(pasta_graficos, 'grafico_agendamento_direc_avalia.png'),
    'graf_perc_liberados_direc_avalia': os.path.join(pasta_graficos, 'grafico_estudantes_liberados_direc_avalia.png'),

    'barra_agend_plurall': os.path.join(pasta_graficos, 'barra_agend_plurall.png'),
    'barra_estud_plurall': os.path.join(pasta_graficos, 'barra_estud_plurall.png'),
    'barra_comp_plurall': os.path.join(pasta_graficos, 'barra_comp_plurall.png'),
    'graf_perc_agendamentos_direc_plurall': os.path.join(pasta_graficos, 'grafico_agendamento_direc_plurall.png'),
    'graf_perc_liberados_direc_plurall': os.path.join(pasta_graficos, 'grafico_estudantes_liberados_direc_plurall.png'),

    't_barra_agend_t': os.path.join(pasta_graficos, 't_barra_agend_t.png'),
    't_barra_estud_t': os.path.join(pasta_graficos, 't_barra_estud_t.png'),
    't_barra_comp_t': os.path.join(pasta_graficos, 't_barra_comp_t.png'),
    't_graf_perc_agendamentos_direc_t': os.path.join(pasta_graficos, 't_grafico_agendamento_direc_t.png'),
    't_graf_perc_liberados_direc_t': os.path.join(pasta_graficos, 't_grafico_estudantes_liberados_direc_t.png'),

    'pa_barra_agend_pa': os.path.join(pasta_graficos, 'pa_barra_agend_pa.png'),
    'pa_barra_estud_pa': os.path.join(pasta_graficos, 'pa_barra_estud_pa.png'),
    'pa_barra_comp_pa': os.path.join(pasta_graficos, 'pa_barra_comp_pa.png'),
    'pa_graf_perc_agendamentos_direc_pa': os.path.join(pasta_graficos, 'pa_grafico_agendamento_direc_pa.png'),
    'pa_graf_perc_liberados_direc_pa': os.path.join(pasta_graficos, 'pa_grafico_estudantes_liberados_direc_pa.png'),   
}

# Iterar sobre todos os slides e formas (shapes)
for slide in prs.slides:
    # list(slide.shapes) é necessário para podermos deletar shapes durante o loop sem quebrar a iteração
    for shape in list(slide.shapes):
        
        # Ignorar formas que não contêm texto
        if not shape.has_text_frame:
            continue
            
        # -------------------------------------------------------------
        # ETAPA A: Substituição de IMAGENS
        # Se o texto da forma for o nome de um gráfico, trocamos pela imagem
        # -------------------------------------------------------------
        imagem_inserida = False
        for texto_chave, caminho_img in substituicoes_imagem.items():
            if texto_chave in shape.text:
                if os.path.exists(caminho_img):
                    # Pega as dimensões da caixa de texto placeholder
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    # Insere a imagem no mesmo lugar. 
                    # Se você não quiser distorção, pode remover o argumento 'height=height' para ele manter a proporção natural
                    slide.shapes.add_picture(caminho_img, left, top, width=width, height=height)
                    
                    # Remove a caixa de texto original do slide
                    sp = shape._element
                    sp.getparent().remove(sp)
                    
                    imagem_inserida = True
                    break # Interrompe a busca de imagens para este shape
                else:
                    print(f"⚠️ Imagem não encontrada: {caminho_img}")
        
        # Se acabamos de deletar o shape para colocar a imagem, pulamos para o próximo shape
        if imagem_inserida:
            continue

        # -------------------------------------------------------------
        # ETAPA B: Substituição de TEXTOS
        # -------------------------------------------------------------
        # Verifica novamento se o shape ainda tem texto (por segurança)
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for texto_chave, valor_texto in substituicoes_texto.items():
                        if texto_chave in run.text:
                            # Substitui o marcador pelo valor real, mantendo a formatação original (fonte, cor)
                            run.text = run.text.replace(texto_chave, str(valor_texto))

# 5. Salvar a apresentação final gerada
prs.save(caminho_saida)
print(f"✅ Apresentação gerada com sucesso em: {caminho_saida}")


























