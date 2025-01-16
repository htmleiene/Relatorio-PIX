import requests
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import smtplib
import base64
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from time import sleep
import os
import base64
import mimetypes
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',    # Para enviar e-mails
    'https://www.googleapis.com/auth/gmail.readonly' # Para ler e-mails
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(BASE_DIR, 'credentials.json')
TOKEN_PATH = os.path.join(BASE_DIR, 'token.json')
SERVICE_ACCOUNT_FILE = CREDENTIALS_PATH

def configurar_proxy():
    # Carregar variáveis do arquivo .env
    load_dotenv()

    # Recuperar usuário e senha
    usuario = os.getenv('PROXY_USERNAME')
    senha = os.getenv('PROXY_PASSWORD')
    if usuario and senha:
        # Configurar proxy com endereço IP
        os.environ['HTTP_PROXY'] = f'http://{usuario}:{senha}@proxy.banese.com.br:8080'
        os.environ['HTTPS_PROXY'] = f'http://{usuario}:{senha}@proxy.banese.com.br:8080'
        print("Proxy configurado.")
    else:
        print("Variáveis de ambiente para proxy não encontradas.")

def authenticate_gmail():
    """Autentica o usuário e retorna o serviço Gmail."""
    creds = None
    # Verifica se já existe um token salvo
    if os.path.exists(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
        except Exception as e:
            print(f"Erro ao carregar token.json: {e}")

    # Se o token não existe ou está inválido, realiza o fluxo de autenticação
    if not creds or not creds.valid:
        try:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open(TOKEN_PATH, 'w') as token:
                token.write(creds.to_json())
        except Exception as e:
            print(f"Erro durante o processo de autenticação: {e}")
            return None

    # Retorna o serviço autenticado
    try:
        service = build('gmail', 'v1', credentials=creds)
        print("Autenticação realizada com sucesso!")
        return service
    except Exception as e:
        print(f"Erro ao construir o serviço Gmail: {e}")
        return None

# Configurações da API
API_URL = r"https://olinda.bcb.gov.br/olinda/servico/Pix_DadosAbertos/versao/v1/odata/EstatisticasTransacoesPix(Database=@Database)?@Database='202501'&$top=100&$format=json&$select=AnoMes,PAG_PFPJ,PAG_REGIAO,PAG_IDADE,VALOR,QUANTIDADE"
PARAMS = {
    "$top": 100,
    "$format": "json",
    "$select": "PAG_PFPJ,PAG_REGIAO,PAG_IDADE,VALOR,QUANTIDADE"
}

# Funções auxiliares
def obter_dados(mes):
    """
    Obtém dados da API para o mês especificado.
    """
    url = API_URL.replace("@Database", f"'{mes}'")
    try:
        response = requests.get(url, params=PARAMS)
        response.raise_for_status()
        return pd.DataFrame(response.json()["value"])
    except Exception as e:
        print(f"Erro ao obter dados da API para o mês {mes}: {e}")
        return pd.DataFrame()

def comparar_dados(df_jan, df_dez):
    """
    Compara os dados de janeiro de 2025 com dezembro de 2024.
    """
    df_comparacao = pd.concat([
        df_jan.assign(mes="Janeiro 2025"),
        df_dez.assign(mes="Dezembro 2024")
    ])
    return df_comparacao

def gerar_graficos(df_comparacao):
    """
    Gera gráficos interativos.
    """
    fig_valor = None
    fig_quantidade = None

    # Gráfico 1: Comparação de valor por região
    if 'PAG_REGIAO' in df_comparacao.columns and 'VALOR' in df_comparacao.columns:
        fig_valor = px.bar(
            df_comparacao,
            x="PAG_REGIAO",
            y="VALOR",
            color="mes",
            title="Comparação de Valor por Região"
        )
    else:
        print("Coluna 'PAG_REGIAO' ou 'VALOR' não encontrada.")

    # Gráfico 2: Quantidade de transações por faixa etária
    if 'PAG_IDADE' in df_comparacao.columns and 'QUANTIDADE' in df_comparacao.columns:
        fig_quantidade = px.bar(
            df_comparacao,
            x="PAG_IDADE",
            y="QUANTIDADE",
            color="mes",
            title="Quantidade de Transações por Faixa Etária"
        )
    else:
        print("Coluna 'PAG_IDADE' ou 'QUANTIDADE' não encontrada.")
    
    return fig_valor, fig_quantidade

def capturar_dashboard(figuras, nome_arquivo):
    """
    Captura as figuras e salva como imagens.
    """
    for i, fig in enumerate(figuras):
        if fig:  # Verificar se a figura é válida
            fig.write_image(f"{nome_arquivo}_fig{i+1}.png")
        else:
            print(f"Figura {i+1} não gerada. Não é possível salvar a imagem.")


def create_message_with_images(sender, to, subject, message_text, image_paths):
    """
    Cria uma mensagem de e-mail com várias imagens no corpo (sem arquivos anexados).
    """
    # Criar a mensagem principal
    message = MIMEMultipart()
    message['to'] = ', '.join(to) if isinstance(to, list) else to  # Suporte para múltiplos destinatários
    message['from'] = sender
    message['subject'] = subject

    # Corpo HTML com as imagens inline
    html_images = ''.join([f'<img src="cid:embedded_image_{i}" alt="Imagem {i}" style="max-width: 600px;"><br>' for i in range(len(image_paths))])
    html_content = f"""
    <html>
    <body>
        <p>{message_text}</p>
        {html_images}
    </body>
    </html>
    """
    html_part = MIMEText(html_content, 'html')
    message.attach(html_part)

    # Adicionar imagens inline
    for i, image_path in enumerate(image_paths):
        with open(image_path, 'rb') as img_file:
            img_data = img_file.read()
            image_part = MIMEImage(img_data)
            image_part.add_header('Content-ID', f'<embedded_image_{i}>')
            image_part.add_header('Content-Disposition', 'inline', filename=image_path.split("/")[-1])
            message.attach(image_part)

    # Codificar a mensagem em base64
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw_message}

def enviar_email(service, user_id, message):
    """
    Envia a mensagem criada usando a API do Gmail.
    """
    try:
        sent_message = service.users().messages().send(userId=user_id, body=message).execute()
        print(f"E-mail enviado! ID da mensagem: {sent_message['id']}")
    except Exception as error:
        print(f"Erro ao enviar e-mail: {error}")

# Script principal
if __name__ == "__main__":
    configurar_proxy()
    service = authenticate_gmail()
    if not service:
        print("Erro ao autenticar o serviço Gmail. Encerrando o script.")
        exit()

    # Obter dados
    print("Obtendo dados...")
    dados_jan = obter_dados("202501")
    dados_dez = obter_dados("202412")

    # Comparar dados
    print("Comparando dados...")
    dados_comparados = comparar_dados(dados_jan, dados_dez)

    # Gerar gráficos
    print("Gerando gráficos...")
    fig_valor, fig_quantidade = gerar_graficos(dados_comparados)

    # Capturar dashboard
    print("Salvando captura de tela...")
    capturar_dashboard([fig_valor, fig_quantidade], "dashboard_pix")

    # Enviar e-mail
    print("Enviando e-mail...")
    sender = "sender@gmail.com"
    to = ["destinario@gmail.com"]
    subject = "Assunto do E-mail"
    message_text = "Corpo do e-mail com imagens"
    image_paths = ["dashboard_pix_fig1.png", "dashboard_pix_fig2.png"]

    # Cria a mensagem
    message = create_message_with_images(sender, to, subject, message_text, image_paths)

    # Envia a mensagem
    enviar_email(service, "me", message)
