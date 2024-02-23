import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def enviar_email_com_anexos(destinatario, assunto, corpo, lista_de_anexos):
    remetente = os.getenv("EMAIL_SENDER")
    senha = os.getenv("EMAIL_PASSWORD")

    # Criar o objeto MIMEMultipart e definir os cabeçalhos
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Adicionar o corpo do email
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar os arquivos da lista_de_anexos
    for anexo in lista_de_anexos:
        part = MIMEBase('application', 'octet-stream')
        with open(anexo, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        nome_anexo = os.path.basename(anexo)
        part.add_header('Content-Disposition', f"attachment; filename= {nome_anexo}")
        msg.attach(part)

    # Configurar o servidor SMTP e enviar o email
    server = smtplib.SMTP('smtp.gmail.com', 587)  # Use o servidor SMTP correto e a porta
    server.starttls()
    server.login(remetente, senha)
    text = msg.as_string()
    server.sendmail(remetente, destinatario, text)
    server.quit()
    print("Email enviado com sucesso!")