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
session.findById("wnd[0]/tbar[0]/okcd").text = "iw39"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_MAB").selected = true
session.findById("wnd[0]/usr/chkDY_HIS").selected = true
session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "ZC04"
session.findById("wnd[0]/usr/ctxtDATUV").text = ""
session.findById("wnd[0]/usr/ctxtDATUB").text = ""
session.findById("wnd[0]/usr/chkDY_HIS").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "AUFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
