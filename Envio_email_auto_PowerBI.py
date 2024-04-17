# Importando as bibliotecas necessárias:

import win32com.client as win32
from datetime import datetime

class Indicadores_Centro_Logístico():
    # Criando a função que envia o indicador com todos os gráficos logísticos:
    def Indicadores_CL():
        # Definindo a variável do dia atual:
        Hoje = datetime.now().date() 
        Hoje = Hoje.strftime('%d/%m/%Y')
        # Enviando o e-mail:
            # Cria uma instância no Outlook:
        outlook = win32.Dispatch('outlook.application')
            # Cria um e-mail:
        email = outlook.CreateItem(0)
            # Configura o e-mail
                # Título
        email.Subject = f'Indicadores Handling - Centro Logístico - {Hoje}'
                # Corpo do e-mail
        email.Body = '''Bom dia.\n
        Segue anexo referente ao indicador do Handling CL.\n
        Att.
        Leandro Fernandes'''
            # Destinatário:
        email.To = 'teste1'
            # Cópia:
        email.Cc = 'teste1'
            # Anexando um arquivo:
        attachment = r'C:\Users\F89074d\Desktop\Indicadores - PB\Indicadores Centro Logístico.pbix'
        email.Attachments.Add(attachment)
            # Enviando o e-mail
        email.Send()
        print('E-mail contendo o indicador dos gráficos logísticos foi enviado com sucesso!')

    # Criando a função que envia o indicador do recebimento:
    def Indicador_Recebimento():
        # Definindo a variável do dia atual:
        Hoje = datetime.now().date() 
        Hoje = Hoje.strftime('%d/%m/%Y')
        # Enviando o e-mail:
            # Cria uma instância no Outlook:
        outlook = win32.Dispatch('outlook.application')
            # Cria um e-mail:
        email = outlook.CreateItem(0)
            # Configura o e-mail
                # Título
        email.Subject = f'Indicador Recebimento - Centro Logístico - {Hoje}'
                # Corpo do e-mail
        email.Body = '''Bom dia.\n
        Segue anexo referente ao indicador de recebimento do CL.\n
        Att.
        Leandro Fernandes'''
            # Destinatário:
        email.To = 'teste1'
            # Cópia:
        email.Cc = 'teste1'
            # Anexando um arquivo:
        attachment = r'C:\Users\F89074d\Desktop\Indicadores - PB\Indicador Recebimento.pbix'
        email.Attachments.Add(attachment)
            # Enviando o e-mail
        email.Send()
        print('E-mail contendo o indicador do recebimento foi enviado com sucesso!')

# Utilizando as funções para testes de performance:

Indicadores_Centro_Logístico.Indicadores_CL()
Indicadores_Centro_Logístico.Indicador_Recebimento()
