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
session.findById("wnd[0]/usr/txtRMIPM-WARPL").text = "GEBRA09470"
WScript.Sleep 5000
session.findById("wnd[0]/usr/cmbRMIPM-MPTYP").key = "ZT"
WScript.Sleep 5000
session.findById("wnd[0]/usr/cmbRMIPM-MPTYP").setFocus
WScript.Sleep 5000
session.findById("wnd[0]").sendVKey 0
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_HEAD:SAPLIWP3:6000/txtRMIPM-WPTXT").text = "GEBRA09470" & " - Sanitizacao Mensal"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY1:SAPLIWP3:8011/subSUBSCREEN_CYCLE:SAPLIWP3:0205/txtRMIPM-ZYKL1").text = "20"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY1:SAPLIWP3:8011/subSUBSCREEN_CYCLE:SAPLIWP3:0205/ctxtRMIPM-ZEIEH").text = "DIA"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_1:SAPLIWO1:0100/ctxtRIWO1-EQUNR").text = "GEBRA09470"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_MAINT_ITEM_TEXT:SAPLIWP3:6005/txtRMIPM-PSTXT").text = "GEBRA09470" & " - Sanitizacao Mensal"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-IWERK").text = "0001"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-WPGRP").text = "ZT"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-AUART").text = "ZMTP"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-ILART").text = "Z01"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-GEWERK").text = "MT-SANI"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/ctxtRMIPM-WERGW").text = "0001"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNTY").text = "A"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNNR").text = "456"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLNAL").text = "1"
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MITEM:SAPLIWP3:8002/tabsTABSTRIP_ITEM/tabpT\11/ssubSUBSCREEN_BODY2:SAPLIWP3:8022/subSUBSCREEN_ITEM_2:SAPLIWP3:0500/txtRMIPM-PLANTEXT").setFocus
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\02").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep 5000
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\02/ssubSUBSCREEN_BODY1:SAPLIWP3:8012/subSUBSCREEN_PARAMETER:SAPLIWP3:0115/chkRMIPM-CALL_CONFIRM").selected = true
WScript.Sleep 5000
session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\02/ssubSUBSCREEN_BODY1:SAPLIWP3:8012/subSUBSCREEN_PARAMETER:SAPLIWP3:0115/chkRMIPM-CALL_CONFIRM").setFocus
WScript.Sleep 5000
session.findById("wnd[0]/tbar[0]/btn[11]").press
WScript.Sleep 5000
session.findById("wnd[1]/tbar[0]/btn[0]").press

returnValue = "success"
