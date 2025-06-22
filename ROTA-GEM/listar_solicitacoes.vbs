Dim arg, data_inicio, data_fim
set arg = WScript.Arguments
data_inicio = arg(0)
data_fim = arg(1)


If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/btn%#REL_001").press
session.findById("wnd[0]/usr/ctxtSO_ERDAT-LOW").text = data_inicio
session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").text = data_fim
session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 3
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "3"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\rmbotelho\Downloads\ROTA_DASH (1)\ROTA_DASH\ListarSolic2023.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ListarSolic2023.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
