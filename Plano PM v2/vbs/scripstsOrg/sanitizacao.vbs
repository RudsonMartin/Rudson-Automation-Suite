' sanitizacao.vbs
Dim Arg, equipNumber, saniDias, returnValue
Set Arg = WScript.Arguments

If Arg.Count < 2 Then
    WScript.Echo "Erro: Argumentos insuficientes. Uso: script.vbs [Equipamento] [Dias]"
    WScript.Quit 1
End If

equipNumber = Arg(0)
saniDias = Arg(1)
'mostra se python está passando certo os argumentos
' x=MsgBox("Equipamento: " & equipNumber & " - Dias de Sanitização: " & saniDias, vbOKOnly, "Sanitização")

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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nip01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtRMIPM-WARPL").text = equipNumber
session.findById("wnd[0]/usr/cmbRMIPM-MPTYP").key = "ZT"
session.findById("wnd[0]/usr/cmbRMIPM-MPTYP").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUBSCREEN_HEAD:SAPLIWP3:6000/txtRMIPM-WPTXT").text = equipNumber & " - Sanitizacao Mensal"
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY1:SAPLIWP3:8011/subSUBSCREEN_CYCLE:SAPLIWP3:0205/txtRMIPM-ZYKL1").text = saniDias
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY1:SAPLIWP3:8011/subSUBSCREEN_CYCLE:SAPLIWP3:0205/ctxtRMIPM-ZEIEH").text = "DIA"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_1:SAPLIWO1:0100/ctxtRIWO1-EQUNR").text = equipNumber
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_MAINT_ITEM_TEXT:SAPLIWP3:6005/txtRMIPM-PSTXT").text = equipNumber & " - Sanitizacao Mensal"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-IWERK").text = "0001"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-WPGRP").text = "ZT"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-AUART").text = "ZMTP"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-ILART").text = "Z01"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-GEWERK").text = "MT-SANI"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-WERGW").text = "0001"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNTY").text = "A"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNNR").text = "456"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text = "1"
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").setFocus
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
WScript.Sleep 1000
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
returnValue = "success"
WScript.Echo(returnValue)