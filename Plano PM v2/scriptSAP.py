import time
import pygetwindow as gw
import subprocess
import win32com.client as wc
from datetime import datetime
import time
from datetime import datetime, timedelta
import pandas as pd


data_hoje = datetime.today().strftime("%d.%m.%Y")
data_passada = (datetime.today() - timedelta(days=1)).strftime("%d.%m.%Y")


# Conectar ao SAP GUI
sapguiauto = wc.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)
 #inserir dados no SAP
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "zgem"
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]/usr/btn%#REL_001").press
session.findById("wnd[0]/usr/ctxtSO_TPSOL-LOW").text = "trt"
session.findById("wnd[0]/usr/ctxtSO_TPSOL-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSO_TPSOL-LOW").caretPosition = 3
session.findById("wnd[0]/usr/btn%_SO_TPSOL_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,1]").press
session.findById("wnd[2]").close
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 0
session.findById("wnd[1]").sendVKey (4)
session.findById("wnd[2]/usr/lbl[1,5]").setFocus
session.findById("wnd[2]/usr/lbl[1,5]").caretPosition = 4
session.findById("wnd[2]").sendVKey (2)
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(0).width = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]").close
session.findById("wnd[2]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/usr/ctxtSO_ERDAT-LOW").text = "010325"
session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").text = "310325"
session.findById("wnd[0]/usr/ctxtSO_STATU-LOW").text = "blq"
session.findById("wnd[0]/usr/ctxtSO_STATU-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSO_STATU-LOW").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 15
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 31
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 50
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 39
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
session.findById("wnd[1]").sendVKey (4)
#Caminho da pasta onde o arquivo será salvo
session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\Users\rmbotelho\Documents"
session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 28
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").pres
    
def Tratamento_dados():
    # Ler o arquivo Excel
    df = pd.read_excel("C:\Users\rmbotelho\Documents\EXPORT.XLSX", sheet_name="Planilha1")
    
    # Exibir as primeiras linhas do DataFrame
    print(df.head())