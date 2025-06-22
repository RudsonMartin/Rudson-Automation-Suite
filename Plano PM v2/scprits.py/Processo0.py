import os
import logging
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import win32com.client as wc
from primarios.envio_planilha import enviar_email,limpar_arquivos_temporarios

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("processamento.log"),
        logging.StreamHandler()
    ]
)

# Constantes
BASE_DIR = Path(r"C:\Users\rmbotelho\Documents\Plano PM v2")
PASTA_EXCEL = BASE_DIR / "excel"
ARQUIVOS = {
    'ih08': PASTA_EXCEL / "ih08.xlsx",
    'ip03': PASTA_EXCEL / "ip03.xlsx",
    'filtrada': BASE_DIR / "ih08_filtrada.xlsx",
    'merged': BASE_DIR / "ih08_ip03_merged.xlsx",
    'sem_plano': BASE_DIR / "equipamentos_sem_plano.xlsx"
}

# Funções de Conexão SAP (mantidas conforme original)
def conectar_sap():
    """Conexão com SAP para gerar relatório IH08"""
    try:
        logging.info("Conectando ao SAP para IH08...")
        sapguiauto = wc.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()

        # Fluxo específico IH08
        session.findById("wnd[0]/tbar[0]/okcd").text = "ih08"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        salvar_planilha(nome_arquivo="ih08.xlsx", pasta_destino=PASTA_EXCEL)
        
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        
        logging.info("Conexão IH08 concluída com sucesso")
        return True

    except Exception as e:
        logging.error(f"Falha na conexão IH08: {str(e)}")
        return False

def conectar_sap2():
    """Conexão com SAP para gerar relatório IP03"""
    try:
        logging.info("Conectando ao SAP para IP03...")
        sapguiauto = wc.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()

        # Fluxo específico IP03
        session.findById("wnd[0]/tbar[0]/okcd").text = "ip03"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[32]").press()
        
        # Repetir processo de seleção
        for _ in range(4):
            session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 4
            session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "4"
            session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell()
        
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        salvar_planilha2(nome_arquivo="ip03.xlsx", pasta_destino=PASTA_EXCEL)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        
        logging.info("Conexão IP03 concluída com sucesso")
        return True

    except Exception as e:
        logging.error(f"Falha na conexão IP03: {str(e)}")
        return False

# Funções de Salvamento (mantidas conforme original)
def salvar_planilha(nome_arquivo="ih08.xlsx", pasta_destino=PASTA_EXCEL):
    try:
        excel = wc.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.ActiveWorkbook
        
        if not workbook:
            logging.error("Nenhuma planilha ativa encontrada para IH08")
            return False

        caminho = pasta_destino / nome_arquivo
        workbook.SaveAs(str(caminho))
        excel.Quit()
        logging.info(f"IH08 salvo em: {caminho}")
        return True

    except Exception as e:
        logging.error(f"Erro ao salvar IH08: {str(e)}")
        return False

def salvar_planilha2(nome_arquivo="ip03.xlsx", pasta_destino=PASTA_EXCEL):
    try:
        excel = wc.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.ActiveWorkbook
        
        if not workbook:
            logging.error("Nenhuma planilha ativa encontrada para IP03")
            return False

        caminho = pasta_destino / nome_arquivo
        workbook.SaveAs(str(caminho))
        excel.Quit()
        logging.info(f"IP03 salvo em: {caminho}")
        return True

    except Exception as e:
        logging.error(f"Erro ao salvar IP03: {str(e)}")
        return False

# Funções de Processamento (otimizadas)
def tratamento_dados():
    try:
        if not ARQUIVOS['ih08'].exists():
            raise FileNotFoundError("Arquivo IH08 não encontrado")

        df = pd.read_excel(ARQUIVOS['ih08'])
        mask = (
            (df['Status sistema'] == 'DEPS ECLI') &
            (~df['Equipamento'].str.startswith('GEBRA', na=False))
        )
        df_filtrado = df[mask].copy()
        
        if df_filtrado.empty:
            logging.warning("Nenhum dado válido após filtragem")
            return False

        df_filtrado.to_excel(ARQUIVOS['filtrada'], index=False)
        ARQUIVOS['ih08'].unlink(missing_ok=True)
        
        logging.info("Dados do IH08 processados")
        return True

    except Exception as e:
        logging.error(f"Falha no tratamento de dados: {str(e)}")
        return False

def tratamento_e_merge():
    try:
        for f in [ARQUIVOS['filtrada'], ARQUIVOS['ip03']]:
            if not f.exists():
                raise FileNotFoundError(f"Arquivo {f.name} não encontrado")

        df_ih08 = pd.read_excel(ARQUIVOS['filtrada'])
        df_ip03 = pd.read_excel(ARQUIVOS['ip03'])
        
        df_ip03['Equipamento_Base'] = df_ip03['Plano manut.'].str[:10]
        df_merged = df_ih08.merge(
            df_ip03,
            how='left',
            left_on='Equipamento',
            right_on='Equipamento_Base',
            suffixes=('_ih08', '_ip03'),
            indicator=True
        )

        df_merged['Dt.criação'] = df_merged['Modificado em_ip03'].combine_first(
            pd.to_datetime(df_merged['Dt.criação']))
        df_merged['Dias_atras'] = (datetime.now() - df_merged['Dt.criação']).dt.days
        
        df_merged.to_excel(ARQUIVOS['merged'], index=False)
        logging.info("Merge concluído com sucesso")
        return True

    except Exception as e:
        logging.error(f"Falha no merge: {str(e)}")
        return False

def gerar_tabela_sem_plano():
    try:
        df_merged = pd.read_excel(ARQUIVOS['merged'])
        sem_plano = df_merged[df_merged['_merge'] == 'left_only']
        
        if not sem_plano.empty:
            sem_plano[['Equipamento', 'Denominação', 'Dt.criação', 'Status sistema']].to_excel(
                ARQUIVOS['sem_plano'], index=False)
            logging.info("Relatório sem plano gerado")
        else:
            logging.info("Todos equipamentos possuem plano")
            
        return True

    except Exception as e:
        logging.error(f"Falha ao gerar relatório: {str(e)}")
        return False    



# Fluxo Principal
def main():
    try:
        logging.info("Iniciando processo completo...")
        
        # Etapa 1: Gerar relatórios SAP
        if not conectar_sap():
            raise RuntimeError("Abortando - Falha no IH08")
        
        if not conectar_sap2():
            raise RuntimeError("Abortando - Falha no IP03")

        # Etapa 2: Processamento de dados
        processamentos = [
            tratamento_dados,
            tratamento_e_merge,
            gerar_tabela_sem_plano
        ]

        for processo in processamentos:
            if not processo():
                raise RuntimeError(f"Processo interrompido - {processo.__name__}")

        # Etapa 3: Finalização e envio
        logging.info("Iniciando procedimentos pós-processamento")
        
        # Envio do e-mail com relatório final
        if not enviar_email():
            raise RuntimeError("Falha no envio do e-mail com relatório")
        
        # Limpeza de arquivos temporários
        limpar_arquivos_temporarios()
        
        logging.info("Processo concluído com sucesso!")
        return 0

    except Exception as e:
        logging.error(f"Erro crítico: {str(e)}")
        logging.info("Preservando arquivos temporários para análise de erros")
        return 1

if __name__ == "__main__":
    sys.exit(main())