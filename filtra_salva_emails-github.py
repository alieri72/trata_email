import os
import win32com.client as win32
import tempfile
import email
from email.utils import formatdate
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# ---------------------
# CONFIGURAÇÃO INICIAL
# ---------------------
# Pasta onde os PDFs serão guardados – o utilizador pode alterar
output_folder = "Output_PDFs"

# Pasta onde o HTML será guardado – normalmente a mesma dos PDFs, mas pode ser alterada
html_output_folder = "Output_PDFs"

# Nome do ficheiro HTML onde o corpo dos emails será guardado
html_filename = "Email_Bodies.html"

# ---------------------
# FUNÇÕES
# ---------------------

def count_emails_in_inbox():
    """Conta o número de emails na caixa de entrada do Outlook"""
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    num_emails = len(inbox.Items)
    return num_emails

def save_emails_as_pdf():
    """Percorre os emails da Inbox e guarda. como PDF, os que têm o assunto definido - ver mais abaixo e substituir XX-some-subject--XX"""
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

    # Criar pastas se não existirem
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    if not os.path.exists(html_output_folder):
        os.makedirs(html_output_folder)

    # Percorre todos os emails na Inbox
    for email_item in inbox.Items:
        # --- AQUI o utilizador deve ajustar o assunto do email que quer processar - substituir XX-some-subject--XX por outra coisa ---
        if email_item.Subject == "XX-some-subject--XX":
            # Extrai um nome personalizado do corpo do email (parte 3)
            subject_parts = email_item.Body.split('|')
            if len(subject_parts) >= 3:
                custom_name = subject_parts[2].strip()
                custom_name = custom_name.replace(":", "")  # Remove ":" do nome do ficheiro
                save_email_as_pdf(email_item, custom_name)
                append_email_body_to_html(email_item)

def save_email_as_pdf(email_item, custom_name):
    """Gera um PDF a partir do conteúdo do email"""
    subject = email_item.Subject
    sender = email_item.SenderEmailAddress
    recipients = ', '.join([recipient.Address for recipient in email_item.Recipients])
    email_date = formatdate(email_item.ReceivedTime.timestamp(), localtime=True)

    # Parse do email
    msg = email.message_from_string(email_item.Body)

    # Criação de ficheiro PDF temporário
    temp_pdf_file = tempfile.NamedTemporaryFile(delete=False)
    temp_pdf_file.close()

    # Configuração do documento PDF
    doc = SimpleDocTemplate(temp_pdf_file.name, pagesize=A4)
    styles = getSampleStyleSheet()

    story = []
    # Adiciona informação do email ao PDF
    story.append(Paragraph(f"Subject: {subject}", styles['Title']))
    story.append(Paragraph(f"From: {sender}", styles['Normal']))
    story.append(Paragraph(f"To: {recipients}", styles['Normal']))
    story.append(Paragraph(f"Date: {email_date}", styles['Normal']))
    story.append(Spacer(1, 12))

    # Adiciona o corpo do email
    for part in msg.walk():
        if part.get_content_type() == 'text/plain':
            body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
            body_lines = body.split("\n| ")
            for line in body_lines:
                story.append(Paragraph(line, styles['Normal']))
                story.append(Spacer(1, 6))  # Quebra de linha

    doc.build(story)

    # Renomeia o ficheiro temporário para o nome final
    output_filename = os.path.join(output_folder, f"{custom_name}.pdf")
    os.rename(temp_pdf_file.name, output_filename)

def append_email_body_to_html(email_item):
    """Adiciona o corpo do email a um ficheiro HTML"""
    html_file_path = os.path.join(html_output_folder, html_filename)

    with open(html_file_path, 'a', encoding='utf-8') as html_file:
        body = email_item.Body
        lines = body.splitlines()

        for line in lines:
            line = line.replace("/nT|", "<li>")
            if "\n" in line:
                bold_red_text = f'<font color="red"><strong>{line.replace("nT|", "<br/>")}</strong></font>'
                html_file.write(bold_red_text + "<hr/>")
            else:
                html_file.write(line + "<br/>")
        html_file.write("<hr/>")

# ---------------------
# FUNÇÃO PRINCIPAL
# ---------------------
def main():
    num_emails = count_emails_in_inbox()
    print(f"Número de emails na Inbox: {num_emails}")
    save_emails_as_pdf()
    print(f"Emails com assunto 'XX-some-subject--XX' foram guardados como PDF na pasta '{output_folder}'.")
    print(f"Corpo dos emails adicionado ao ficheiro HTML '{html_filename}'.")

# Executa a função principal
if __name__ == "__main__":
    main()
