#!/usr/bin/env python
# coding: utf-8

# # RPA com Python
# 
# ### Vamos automatizar a extração de informações de um sistema e envio de um relatório por e-mail
# - No nosso caso, para todo mundo conseguir fazer o mesmo programa, o nosso "sistema" vai ser o Gmail, mas o mesmo processo pode ser feito com qualquer programa do seu computador e qualquer sistema
#     - Passo 1: Entrar no sistema (entrar no Gmail)
#     - Passo 2: Exportar o Relatório (Exportar a planilha Financeiro)
#     - Passo 3: Pegar o relatório exportado, tratar e pegar as informações que queremos (análise de dados)
#     - Passo 4: Preencher/Atualizar informações do sistema (No nosso caso, criar um e-mail e enviar)
#     - Passo 5: Criar janela Tkinter para executar a automação em 1 clique

# In[1]:


import pyautogui as gui 
import pyperclip 
import time
import pandas as pd
import os
import shutil
import smtplib
import email.message


# ###### Configurações do pyautogui

# In[2]:


gui.PAUSE = 1
gui.FAILSAFE = True # arrastando o mouse para o canto esquerdo superior pausa a automação em caso de erro
gui.alert('NÃO MEXER NO MOUSE E TECLADO!')


# ###### Passo a passo da automação RPA
# Obs.: esta célula foi desativada do código para não atrasá-lo, já que não foi possível o pyautogui encontrar as imagens pelo método .locateOnScreen().

# # entrar no sistema (Google Drive)
# gui.press('win')
# gui.write('chrome')
# gui.press('enter')
# gui.write('drive')
# gui.press('enter')
# 
# # aguardar o drive abrir
# #while not gui.locateOnScreen('logo_drive.png', grayscale=True):
# #    time.sleep(0.5)
# #print('ok')
#     
# # pela repetição de código, cria-se função para tentar por 5 segundos encontrar e clicar em imagens
# def timer_encontrar_clicar_imagem(imagem):
#     timer = 5
#     while timer:
#         time.sleep(1)
#         
#         try:
#             x, y, largura, altura = gui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
#             gui.click(x+largura/2, y+altura/2)
#         except:
#             print(f'Timer zerado, imagem {imagem} não encontrada')
#         
#         timer -= 1
# 
# # tentar encontrar e clicar na imagem da planilha, 'mais ações', 'fazer download'
# timer_encontrar_clicar_imagem('planilha.png')
# timer_encontrar_clicar_imagem('mais_acoes.png')
# timer_encontrar_clicar_imagem('download.png')
# time.sleep(5) #download
# 
# # fechar o navegador
# gui.hotkey('alt', 'f4')

# ###### Passo a passo da automação RPA
# Aqui utilizou-se o método .click() do pyautogui para a continuidade da automação.

# In[39]:


# caso as imagens não sejam encontradas, pode-se fazê-la de outro maneira:
pasta_downloads = r'C:\Users\W10\Downloads'
lista_arquivos = os.listdir(pasta_downloads)
    
if 'Financeiro.xlsx' not in lista_arquivos:
    print('Arquivo NÃO encontrado na pasta Downloads')

    # entrar no sistema (Google Drive)
    gui.press('win')
    gui.write('chrome')
    gui.press('enter')
    gui.write('drive')
    gui.press('enter')
    time.sleep(2)
    gui.press('f11')

    # aguardar o drive abrir
    #while not gui.locateOnScreen('logo_drive.png', grayscale=True):
    #    time.sleep(0.5)
    #print('ok')

    # clicar nas na planilha e em download
    gui.click(x=641, y=539)
    time.sleep(1)
    gui.click(x=600, y=102)
    time.sleep(3)

else:
    pass

# fechar o navegador
gui.hotkey('alt', 'f4')


# In[21]:


# mover o arquivo baixado para a pasta específica e, então, deletá-lo
arquivo_excel = r'C:\Users\W10\Downloads\Financeiro.xlsx'
local_arquivo_python = os.getcwd()

if os.path.isfile(arquivo_excel) == True:
    print(f'{arquivo_excel} encontrado em Downloads')
    # mover o arquivo
    local_arquivo = arquivo_excel
    local_destino = local_arquivo_python
    shutil.copy2(local_arquivo, local_destino)
    # deletar o arquivo
    os.remove(local_arquivo)    


# Análise de dados:
# - qual aplicativo mais gera renda no período? (gráfico)
# - qual a porcentagem dos lucros destinada aos gastos?

# In[27]:


# importação da base de dados
df = pd.read_excel('Financeiro.xlsx' , dtype={'mês':str})
#display(df)
#df.info()


# In[28]:


# tratamento dos dados: coluna 'mês'
dicionario = {
    '01':'janeiro',
    '02':'fevereiro',
    '03':'março',
    '04':'abril',
    '05':'maio',
    '06':'junho',
    '07':'julho',
    '08':'agosto',
    '09':'setembro',
    '10':'outubro',
    '11':'novembro',
    '12':'dezembro',
}

for data in df['mês']:
    chave = data[5:7]
    df.loc[df['mês']==data, 'mês'] = dicionario[chave]

#display(df)


# In[29]:


# tratamento dos dados: linha 'outubro/2022', coluna 'extra'
df.loc[0, ['uber','extra','99','gastos']] = 0
#display(df)


# In[30]:


# tratamento dos dados: colunas 'extra' e 'uber'
df['extra'] = df['extra'].astype('float')
df['uber'] = df['uber'].astype('float')
df.info()


# In[31]:


# análise de dados: qual aplicativo mais gera renda no período? (gráfico) & qual a porcentagem dos lucros destinada aos gastos?
df_renda = df.iloc[:,[0,2,3,4,5]]
df_renda = df_renda.drop(0, axis=0)
df_renda = df_renda.T
df_renda.columns = df_renda.iloc[0] 
df_renda = df_renda.drop('mês', axis=0)

for coluna in df_renda.columns:
    df_renda[coluna] = df_renda[coluna].astype('float')

df_renda['TOTAL'] = df_renda.sum(axis=1) # soma da linha


df_renda = df_renda.sort_values(by='TOTAL', ascending=False)
#display(df_renda)


# In[32]:


total_gastos = df_renda.loc['gastos','TOTAL']
total_renda_bruta = df_renda['TOTAL'].sum()
total_renda_liquida = total_renda_bruta - total_gastos

n_colunas = len(df_renda.columns)-1
media_mensal = total_renda_liquida / n_colunas
media_gastos = total_gastos / n_colunas

maior_renda = df_renda.iloc[0,-1]
porcentagem_maior_renda = maior_renda / total_renda_liquida

segunda_renda = df_renda.iloc[1,-1]
porcentagem_segunda_renda = segunda_renda / total_renda_liquida

terceira_renda = df_renda.iloc[2,-1]
porcentagem_terceira_renda = terceira_renda / total_renda_bruta

quarta_renda = df_renda.iloc[3,-1]
porcentagem_quarta_renda = quarta_renda / total_renda_liquida


# In[33]:


analises = f"""Análise financeira UBER/99 desde <strong>novembro/2022</strong>:<br>
<br>
<strong>Renda bruta:</strong> R$ {total_renda_bruta:,.2f}<br>
<strong>Renda líquida:</strong> R$ {total_renda_liquida:,.2f}<br>
<strong>Renda Média mensal:</strong> R$ {media_mensal:,.2f}<br>
<strong>Gasto Médio mensal:</strong> R$ {media_gastos:,.2f}<br>
<br>
<strong>{df_renda.index[0]}:</strong> {porcentagem_maior_renda:.0%} da renda líquida<br>
<strong>{df_renda.index[1]}:</strong> {porcentagem_segunda_renda:.0%} da renda líquida<br>
<strong>{df_renda.index[2]}:</strong> {porcentagem_terceira_renda:.0%} da renda bruta<br>
<strong>{df_renda.index[3]}:</strong> {porcentagem_quarta_renda:.0%} da renda líquida<br>
<br>
{df_renda.to_html(formatters={coluna: '{:,.2f}'.format for coluna in df_renda.columns})}
"""


# In[34]:


# envio do email com as informações desejadas, via gmail
server = smtplib.SMTP('smtp.gmail.com:587')  
corpo_email = f'{analises}'

msg = email.message.Message()
msg['Subject'] = "Informações Financeiras Uber/99"

# Passo a passo na 1ª vez: Clicar na foto do perfil no Gmail -> Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
# Depois, faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha
msg['From'] = 'beprafael@gmail.com'
msg['To'] = 'bep_rafael@hotmail.com'
password = "xxhvvhkyafpjeufo" # senha gerada para envio de e-mail por meio de outros apps
msg.add_header('Content-Type', 'text/html')
msg.set_payload(corpo_email )

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado')


# In[ ]:


#criar alerta de fim da automação
gui.alert('Automação concluida!!!')


# # criar executável
