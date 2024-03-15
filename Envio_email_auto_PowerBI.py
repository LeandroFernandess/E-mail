# Importando as bibliotecas necessárias:

import win32com.client as win32
from datetime import datetime


# Definindo a variável do dia atual:
Hoje = datetime.now().date() # Obtém a data atual
Hoje = Hoje.strftime('%d/%m/%Y')
# Enviando o e-mail:
    # Cria uma instância no Outlook:
outlook = win32.Dispatch('outlook.application')
# Cria um e-mail:
email = outlook.CreateItem(0)
# Configura o e-mail
email.Subject = f'Indicadores Handling - Centro Logístico - {Hoje}'
email.Body = '''Bom dia.\n
Segue anexo referente ao indicador do Handling CL.\n
at.
Leandro Fernandes'''
email.To = 'e-mail 1' 
email.Cc = 'e-mail 2'
# Anexando um arquivo:
attachment = r'C:\Users\F89074d\Desktop\Indicadores - PB\Indicadores Centro Logístico.pbix'
email.Attachments.Add(attachment)
# Enviando o e-mail
email.Send()
print('E-mail enviado com sucesso!')
