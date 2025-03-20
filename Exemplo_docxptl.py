# Tabela de artistas retirada de https://www.gazetaesportiva.com/bastidores/confira-os-times-de-coracao-dos-famosos/#foto=33
# (pip install docxtpl)
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
import win32com.client

import pandas as pd

# Ler o arquivo
df = pd.read_csv('C:\\docxptl_exemplo\\times_dos_famosos.csv', sep=';')

modelo = DocxTemplate('C:\\docxptl_exemplo\\doc_xptl_exemplo.docx')

# Preecher a lista de personalidades (nome e respectivos atividade e time)
# E uma lista de dicionarios, cara personalidade tem seu dicionario na lista, com seus dados
personalidades = list()
df = df.reset_index()
for indice, linha in df.iterrows():
  personalidades.append(
    {'atividade' : linha['atividade'],
     'artista'   : linha['artista'],
     'time'      : InlineImage(modelo,
                                           'C:\\docxptl_exemplo\\' + linha['time'] + '.png',
                                            width=Mm(32))}
     )
# next

times = list()
df_times = df.drop_duplicates(subset='time')
df_times.reset_index()
for indice, linha in df.iterrows():
  link = RichText()
  link.add(linha['time'],
           url_id = modelo.build_url_id(linha['pag_time']),
           color='blue')
  
  times.append(
    {'time' : linha['time'],
    'link'  : link}
  )
# next

# Aqui o dicionario do que vai ser preenchido no modelo do Word
# O primeiro item da lista so aparece uma vez.
# O segundo e a lista de personalidades e os dados
# O terceiro e a lista dos times de futebol em RTF para que tenha o link
parametros = {
  'frase_inicial' : 'Aqui começa a lista das personaliades.',
  'personalidades': personalidades,
  'times'         : times
}

modelo.render(parametros)
saida = 'C:\\docxptl_exemplo\\relatorio.docx'
try:
  modelo.save(saida)
  # word = win32com.client.Dispatch("Word.Application")
  # word.Documents.Open(saida)
  # word.Visible  = True
  print('Fim.')
except:
  print('Nao foi possível salvar o relatório; verifique se o arquivo não está aberto.')
# fim try







