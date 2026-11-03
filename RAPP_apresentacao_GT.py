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

# Carregar base de dados do Avalia Potiguar:
df_avalia = pd.read_excel(r'D:\Scripts_Python\FGV\RAPP_apresentacao_GT\dados\relatorio_acompanhamento_rapp_ept_div.xlsx', sheet_name='Relatório RAPP')

# AGENDAMENTO - AVALIA POTIGUAR
# 'agend_avalia' é composto por 'componentes agendados' e total previstos de agendamentos (agendados + não agendados)




# barra_agend_avalia
# (PAUSADO POR ENQUANTO. FAZER QUANDO FOR MONTAR O PPT)


# ESTUDANTE - AVALIA POTIGUAR
# 'estud_avalia' é composto por 'estudantes liberados' (liberad_avalia) e 'total de estudantes'
# Total de Estudantes
total_estud_avalia = df_avalia['MATRÍCULA SISTEMA'].nunique()

# Estudantes Liberados



# Estudante liberado (liberad_avalia) é aquele 'APROVADO' em todos seus componentes curriulares





# barra_estud_avalia
# (PAUSADO POR ENQUANTO. FAZER QUANDO FOR MONTAR O PPT)



# COMPONENTES - AVALIA POTIGUAR
# 'comp_avalia' é composto por 'componentes aprovados' e 'total de componentes'





# barra_comp_avalia
# (PAUSADO POR ENQUANTO. FAZER QUANDO FOR MONTAR O PPT)




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
plurall_turmas_agendadas = (
    df_plurall_agend
    .groupby(['ID ESCOLA', 'ID TURMA'])['AGENDOU']
    .all()
)

# Quantidade de turmas agendadas
qtd_plurall_turmas_agendadas = plurall_turmas_agendadas.sum()

# perc_agend_plurall: turmas agendadas / total de turmas de agendamento
perc_agend_plurall = round((qtd_plurall_turmas_agendadas / total_agend_plurall) * 100, 2)


# barra_agend_plurall
# (PAUSADO POR ENQUANTO. FAZER QUANDO FOR MONTAR O PPT)

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
perc_estud_plurall = round((liberad_plurall / total_estud_plurall) * 100, 2)

# barra_estud_plurall
# (PAUSADO POR ENQUANTO. FAZER QUANDO FOR MONTAR O PPT)


# COMPONENTES - PLURALL
# 'comp_plurall' é composto por 'componentes aprovados' e 'total de componentes'

# Total de componentes
total_comp_plurall = df_plurall_prova['COMPONENTE CURRICULAR'].count()

# Componentes aprovados
aprovados_comp_plurall = df_plurall_prova[df_plurall_prova['Situação'] == 'Aprovado']['COMPONENTE CURRICULAR'].count()

# perc_comp_plurall: componentes aprovados / total de componentes
perc_comp_plurall = round((aprovados_comp_plurall / total_comp_plurall) * 100, 2)

# barra_comp_plurall
# (PAUSADO POR ENQUANTO. FAZER QUANDO FOR MONTAR O PPT)


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
    labels=[f'{valor:.1f}%' for valor in df_grafico['PERCENTUAL']],
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

# Salvar PNG
plt.savefig(
    'grafico_agendamento_direc_plurall.png',
    dpi=300,
    transparent=True,
    bbox_inches='tight'
)

plt.show()



# graf_perc_liberados_direc_plurall: Gráfico de barras com percentual de estudantes liberados por DIREC


