import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

def ler_planilha(nome_arquivo, tipo='excel'):
    """Lê dados de uma planilha Excel ou CSV e retorna um DataFrame."""
    if tipo == 'excel':
        df = pd.read_excel(nome_arquivo)
    elif tipo == 'csv':
        df = pd.read_csv(nome_arquivo)
    else:
        raise ValueError('Tipo de arquivo inválido. Use "excel" ou "csv".')
    return df

def filtrar_dados(df, coluna, valor):
    """Filtra um DataFrame por um valor específico em uma coluna."""
    return df[df[coluna] == valor]

def calcular_estatisticas(df, coluna):
    """Calcula estatísticas básicas para uma coluna específica."""
    return {
        'Soma': df[coluna].sum(),
        'Média': df[coluna].mean(),
        'Máximo': df[coluna].max(),
        'Mínimo': df[coluna].min()
    }

def enviar_email(destinatario, assunto, mensagem, anexo=None):
    """Envia um email com um anexo opcional."""
    msg = MIMEMultipart()
    msg['From'] = 'seu_email@example.com'
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(mensagem, 'plain'))

    if anexo:
        with open(anexo, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{anexo}"')
            msg.attach(part)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login('seu_email@example.com', 'sua_senha')
        server.sendmail('seu_email@example.com', destinatario, msg.as_string())

def gerar_relatorio_pdf(dados, nome_arquivo):
    """Gera um relatório em PDF com os dados fornecidos."""
    c = canvas.Canvas(nome_arquivo, pagesize=letter)
    c.drawString(inch, 10 * inch, 'Relatório de Dados')

    for i, row in enumerate(dados):
        c.drawString(inch, (10 - i * 0.5) * inch, f'{row}')

    c.save()

def main():
    df = ler_planilha('dados.xlsx', tipo='excel')
    clientes_sul = filtrar_dados(df, 'Região', 'Sul')
    estatisticas_vendas = calcular_estatisticas(clientes_sul, 'Vendas')
    gerar_relatorio_pdf(clientes_sul.to_numpy(), 'relatorio_sul.pdf')
    enviar_email('gerente@example.com', 'Relatório de Clientes do Sul', 'Segue o relatório com os dados dos clientes do Sul.', 'relatorio_sul.pdf')

if __name__ == '__main__':
    main()