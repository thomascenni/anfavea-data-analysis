#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import datapane as dp
import altair as alt
import ssl

dp.Params.load_defaults('datapane.yaml')
# default year is 2022
year = dp.Params.get('year')

url = 'https://www.anfavea.com.br/docs/siteautoveiculos{}.xlsx'.format(str(year))
print("Downloading Excel from url {}".format(url))

# avoid urlopen error [SSL: CERTIFICATE_VERIFY_FAILED]
ssl._create_default_https_context = ssl._create_unverified_context

# avoid HTTP Error 406
headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.14; rv:66.0) Gecko/20100101 Firefox/66.0"}

df = pd.read_excel(
    url,
    storage_options=headers,
    engine='openpyxl',
    sheet_name='IV. Licenciamento Empresa',
    usecols='B:P',
    header=6,
    names=['Segmento', 'Associada', 'Marca','Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
)


# preenche coluna de Segmento de veiculo
filtered_segmento = df[df['Segmento'].notnull()]
filtered_segmento_index=list(filtered_segmento.index)
filtered_segmento_name=list(filtered_segmento.Segmento)
filtered_segmento_index = [(filtered_segmento_index[i],filtered_segmento_index[i+1],filtered_segmento_name[i]) for i in range(len(filtered_segmento_index)-1)]
for slice_segmento in filtered_segmento_index:
    df.update(df.iloc[range(slice_segmento[0], slice_segmento[1]), 0].fillna(slice_segmento[2]))
# completa as ultimas linhas que ficaram sem preenchimento
df.update(df.loc[:,'Segmento'].fillna(filtered_segmento_name[len(filtered_segmento_name)-1]))


# preenche coluna de Associada
df['Associada'] = df['Associada'].str.strip()

filtered_empresa = df[df['Associada'].notnull()]
filtered_empresa_index=list(filtered_empresa.index)
filtered_empresa_name=list(filtered_empresa.Associada)
filtered_empresa_index = [(filtered_empresa_index[i],filtered_empresa_index[i+1],filtered_empresa_name[i]) for i in range(len(filtered_empresa_index)-1)]
for slice_empresa in filtered_empresa_index:
    df.update(df.iloc[range(slice_empresa[0], slice_empresa[1]), 1].fillna(slice_empresa[2]))


# elimina os indices de Segmento
df.drop(filtered_segmento.index, inplace = True)

# elimina as linhas de totais que tem Associada e Marca vazios (somente a primeira linha !)
to_remove = df[(df.Marca.isnull()) & (df.Associada.isnull())]
df.drop(to_remove.index, inplace = True)

# filtra empresas associadas a anfavea com grupo null e remove
to_remove = df[(df.Marca.isnull()) & (df['Associada'] == 'Empresas associadas à Anfavea')]
df.drop(to_remove.index, inplace = True)


# filtra outras empresas para colocar Marca Outro
df.update(df[df.Associada == 'Outras empresas'].Marca.fillna('Outra'))

# procura sub categorias para caminhoes (na coluna Associada !)
filtered_grupo = df[df['Marca'].isnull()]
filtered_grupo_index=list(filtered_grupo.index)
filtered_grupo_name=list(filtered_grupo.Associada)
filtered_grupo_index = [(filtered_grupo_index[i],filtered_grupo_index[i+1],filtered_grupo_name[i]) for i in range(len(filtered_grupo_index)-1)]

# adiciona uma coluna para essa categoria
df.insert(1, "SubSegmento", np.nan)


# preenche coluna subsegmento
for slice_grupo in filtered_grupo_index:
    df.loc[slice(slice_grupo[0], slice_grupo[1]-1), 'SubSegmento'] = slice_grupo[2]

# preenche as ultimas linhas de Caminhoes
df.SubSegmento = df.SubSegmento.mask(
    (df.Segmento == 'Caminhões') & (df.SubSegmento == ""), 
    filtered_grupo_name[len(filtered_grupo_name)-1]
)

# preenche o subsegmento de automoveis, comerciais leves e onibus
df.update(df.SubSegmento.fillna('Todos'))

# elimina as linhas com Marca null
to_remove = df[df.Marca.isnull()]
df.drop(to_remove.index, inplace = True)


# agora precisa criar a coluna Grupo
df.insert(3, "Grupo", np.nan)


# remove os eventuais espaços na extremidade direita da marca
df['Marca'] = df['Marca'].str.rstrip()


# procura os espaços (indentaçao) na estremidade esquerda para achar grupos apartenentes a grupo
def getGrupo(str):
    if (len(str) - len(str.lstrip(' ')) == 0):
        return(str)
    else:
        return('')

grupo = [getGrupo(item) for item in df.Marca.to_numpy()]

last_item = ''
for idx, val in enumerate(grupo):
    if not val:
        grupo[idx] = last_item
    else:
        last_item = val

df.Grupo = grupo

# grupa e conta as grupos - importante colocar sort  = False para manter a ordem
df_grupo=df.groupby(['Segmento','Grupo', 'SubSegmento'], sort=False).Grupo.agg('count').to_frame('COUNT').reset_index()

# para cada grupo no mesmo segmento
# se count == 1 nao faço nada, mantenho
# se count > 1 elimino a primeira ocorrencia
to_remove = []
for val in df_grupo.COUNT:
    if val == 1:
        to_remove.append(False)
    else:
        sublist = []
        sublist = np.full(val, False)
        sublist[0] = True
        for item in sublist:
            to_remove.append(item)

to_remove = df[to_remove]
df.drop(to_remove.index, inplace = True)



# limpa espaços na esquerda das marcas
df.Marca=df.Marca.str.lstrip()


# elimina as linhas Caminhoes Total Por Empresa pois é uma informação já disponivel no segmento caminhoes
to_remove = df[df['Segmento'] == 'Caminhões - Total por empresa']
df.drop(to_remove.index, inplace = True)

# faz limpeza para Grupo MAN
df.Grupo = df.Grupo.mask(
    df['Grupo'].str.startswith('MAN'), 
    'MAN'
)

# exporta para Excel
df.to_excel('anfavea_data_analysis.xlsx', index=False)
# df.to_csv('anfavea_data_analysis.csv', index=False)

df['Total (YTD)'] = df.sum(axis=1)
df_total_sales = df[['Segmento', 'Marca', 'Total (YTD)']].copy()

# seleciona somente as Automoveis
df_total_sales =  df_total_sales[df_total_sales['Segmento']=='Automóveis']
# df_total_sales = df_total_sales.sort_values(by=['Total (YTD)'], ascending=False)
# seleciona o top 10
top10 = df_total_sales.nlargest(10, 'Total (YTD)')

plot = alt.Chart(top10).mark_bar().encode(
    x='Total (YTD)',
    y = alt.Y('Marca', sort='-x')
)

report = dp.Report(
    dp.Text("## Anfavea data analysis"),
    dp.Text("Year: {}".format(str(year))),
    dp.Text("**Cleaned data original from Excel**"),
    dp.DataTable(df),
    dp.Text("**Top 10 Total Sales YTD** for Automóveis"),
    dp.DataTable(top10),
    dp.Plot(plot)
)

report.upload(name='{} Anfavea Data Analysis'.format(year), publicly_visible = True)
