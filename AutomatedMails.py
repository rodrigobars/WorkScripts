import os
import smtplib
from email.message import EmailMessage
import pandas as pd

email = input('Insira o e-mail \U0001F4E7: ')
senha = input('Insira a senha \U0001F512: ')
path_mails = input('Insira o caminho do arquivo mails: ')
mails = pd.read_table(rf'{path_mails}',
                      delimiter=";", names=["empresa", "e-mail"])

EMAIL_ADRESS = email
EMAIL_PASS = senha

pregao = input("Informe o pregão(Ex: 21/2021): ")
homologacao = input("Informe a data de homologação(Ex: 30/09/2021): ")
objeto = input("Informe o objeto da licitação: ")
ata = input("Informe o caminho da Ata: ")
folder_termo = input(
    "Informe a pasta que contém os Termos de Responsabilidade: ")

for empresa in range(len(mails)):
    print(
        f"\n\n\n\n\n\n\n\n Progresso - {round((empresa/(len(mails)))*100, 2)}% \n")
    msg = EmailMessage()
    msg["Subject"] = f"PR {pregao} - Ata de Registro de Preço"
    msg["From"] = EMAIL_ADRESS
    msg["To"] = mails.iloc[empresa, 1]
    msg.set_content(f"""
    <!DOCTYPE html>
    <html>
    <body>
        <p>Prezados, boa tarde...</p>
        <br>
        <p>No dia <b>{homologacao}</b> foi realizada a homologação do pregão <b>{pregao} - {objeto}.</b></p>
        <br>
        <p>Dessa forma, encaminho a Ata de registro de preços já assinada pelo responsável da PROAD, junto ao Termo de Responsabilidade sobre a Ata, onde somente o Termo de Responsabilidade deve ser devidamente ASSINADO (digitalmente ou não) e CARIMBADO para posterior devolução através deste e-mail. </p>
        <br>
        <p>Informo que não é necessário enviar pelos correios e sim somente digitalizada por email.</p>
        <br>
        <p>Por gentileza confirmar o recebimento deste email.</p>
        <br>
        <blockquote><small><i>Obs: a tag "externa" é devida ao fato da mensagem estar sendo disparada para todas as empresas vencedoras de forma automatizada.</i></small></blockquote>
    </body>
    </html>
    """, subtype='html')

    with open(rf'{ata}', 'rb') as f:
        archiveName = os.path.basename(r'{}'.format(f.name))
        file_data = f.read()
        msg.add_attachment(file_data, maintype='application',
                           subtype='pdf', filename=archiveName)

    path_termo = folder_termo+"\\"+mails.iloc[empresa, 0]+".docx"
    with open(rf'{path_termo}', 'rb') as w:
        archiveName = os.path.basename(r'{}'.format(w.name))
        file_data = w.read()
        msg.add_attachment(file_data, maintype='application',
                           subtype='docx', filename=archiveName)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADRESS, EMAIL_PASS)
        smtp.send_message(msg)

print(
    f"\n\n\n\n\n\n\n\n Progresso - 100.00% \n")
print("Enviado \U0001F7E2")
