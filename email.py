import win32com.client as win32
import pandas as pd

# enviar emails
outlook = win32.Dispatch("Outlook.Application")

# carregando os dados
gerentes_df = pd.read_excel(r'E:\Python\python\Impressionador\Pasta auxiliar\Enviar E-mails.xlsx')


for i, email in enumerate(gerentes_df['E-mail']):
    # usar loc sempre que for trabalhar com for em planilhas
    gerente = gerentes_df.loc[i, 'Gerente']
    relatorio = gerentes_df.loc[i, 'Relatório']

    # criando o email
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = f'Relatorio de {relatorio}'
    mail.Body = f'teste email {gerente} {relatorio}'

    # Incluindo o anexo
    attachment = r'E:\Python\python\Impressionador\Pasta auxiliar\{}.xlsx'.format(relatorio)
    mail.Attachments.Add(attachment)

    mail.Send()
