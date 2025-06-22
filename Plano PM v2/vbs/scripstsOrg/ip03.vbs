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
session.findById("wnd[0]/tbar[0]/okcd").text = "ip03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[0]/usr/ctxtERSDT-LOW").text = "010325"
session.findById("wnd[0]/usr/ctxtERSDT-HIGH").text = "010425"
session.findById("wnd[0]/tbar[0]/okcd").text = "ip03"
session.findById("wnd[0]/usr/ctxtERSDT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtERSDT-HIGH").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press
