import smtplib
import os
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
import pandas as pd
from datetime import datetime


# Configurações de caminho
BASE_DIR = Path(r"C:\Users\rmbotelho\Documents\Plano PM v2")
ORIGINAL_FILE = BASE_DIR / "ih08.xlsx"
FILTERED_FILE = BASE_DIR / "ih08_filtrada.xlsx"
MERGED_FILE = BASE_DIR / "ih08_ip03_merged.xlsx"
SEM_PLANO_FILE = BASE_DIR / "equipamentos_sem_plano.xlsx"
IP03_FILE = BASE_DIR / "excel" / "ip03.xlsx"

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(BASE_DIR / "processamento.log"),
        logging.StreamHandler()
    ]
)


def limpar_arquivos_temporarios():
    """Remove todos os arquivos gerados pelo processo"""
    try:
        arquivos_para_excluir = [
            FILTERED_FILE,
            MERGED_FILE,
            IP03_FILE
        ]
        
        for arquivo in arquivos_para_excluir:
            if arquivo.exists():
                arquivo.unlink()
                logging.info(f"Arquivo excluído: {arquivo.name}")
                
        return True
        
    except Exception as e:
        logging.error(f"Falha ao excluir arquivos: {str(e)}")
        return False

def enviar_email():
    """Envia o arquivo de equipamentos sem plano por e-mail."""
    try:
        if not SEM_PLANO_FILE.exists():
            logging.error("Arquivo de equipamentos sem plano não encontrado!")
            return False

        # Carregar dados para o corpo do e-mail
        df = pd.read_excel(SEM_PLANO_FILE)
        total_equipamentos = len(df)
        data_processamento = datetime.now().strftime("%d/%m/%Y %H:%M")

        # Configurações do e-mail (ALTERAR!)
        smtp_server = "smtp.outlook.com"
        smtp_port = 587
        email_from = ""
        email_password = ""  # Usar senha de aplicativo
        email_to = [""]

        # Criar mensagem
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = ", ".join(email_to)
        msg['Subject'] = f"Equipamentos Sem Plano - {datetime.now().strftime('%d/%m/%Y')}"

        # Corpo do e-mail
        body = f"""
        Relatório de Equipamentos Sem Plano de Manutenção

        Detalhes:
        - Data do processamento: {data_processamento}
        - Total de equipamentos: {total_equipamentos}
        - Lista completa em anexo

        Ação Requerida:
        1. Verificar equipamentos listados
        2. Criar planos de manutenção faltantes
        3. Atualizar sistema SAP

        Att.,
        Sistema de Automação de Manutenção
        """
        msg.attach(MIMEText(body, 'plain'))

        # Anexar arquivo
        with open(SEM_PLANO_FILE, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{SEM_PLANO_FILE.name}"'
            )
            msg.attach(part)

        # Enviar e-mail
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_from, email_password)
            server.sendmail(email_from, email_to, msg.as_string())
        
        logging.info("E-mail com equipamentos sem plano enviado com sucesso!")
        return True

    except Exception as e:
        logging.error(f"Erro ao enviar e-mail: {str(e)}", exc_info=True)
        return False