import time
from datetime import datetime, timedelta
import os
import win32com.client as wc
from datetime import datetime
import pandas as pd
import win32com.client as wc
from datetime import datetime


def conectar_sap2():
    try:
        # Conectar ao SAP
        sapguiauto = wc.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)  # Obtém a primeira conexão
        session = connection.Children(0)  # Obtém a primeira sessão

        session.findById("wnd[0]").maximize()  # Corrigido com parênteses
        session.findById("wnd[0]/tbar[0]/okcd").text = "ip03"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(4)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[32]").press()
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 4
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "4"
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell()
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "4"
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell()
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "4"
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell()
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "4"
        session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        salvar_planilha2()  # Salva a planilha após o processo no SAP
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        print("SAP 2 automatizado com sucesso.")
        return session
    except Exception as e:
        print(f"Erro ao conectar ou automatizar o SAP: {e}")
        return None
    
def salvar_planilha2(nome_arquivo="ip03.xlsx", pasta_destino=os.getcwd()):
    try:
        # Conectar ao Excel
        excel = wc.Dispatch("Excel.Application")
        excel.Visible = False  # Deixe o Excel visível (opcional)

        # Obtém o workbook ativo
        workbook = excel.ActiveWorkbook

        if workbook is None:
            print("Nenhuma planilha ativa encontrada.")
            return

        # Caminho completo para salvar a planilha
        caminho_salvar = os.path.join(pasta_destino, nome_arquivo)

        # Salva a planilha no local especificado
        workbook.SaveAs(caminho_salvar)
        print(f"Planilha salva em: {caminho_salvar}")
        excel.Quit()  # Fecha o Excel após salvar

    except Exception as e:
        print(f"Erro ao salvar a planilha: {e}")

session = conectar_sap2()
if session:
    salvar_planilha2(nome_arquivo="ip03.xlsx", pasta_destino=r"C:\Users\rmbotelho\Documents\Plano PM v2\excel")
