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
session.findById("wnd[0]/usr/ctxtDATUB").setFocus
session.findById("wnd[0]/usr/ctxtDATUB").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 6505
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 6506
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7658
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7657
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7649
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7635
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7621
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7603
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7584
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7566
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 7595,"AUFNR"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 7577
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
session.findById("wnd[1]/usr/tblSAPLSWUGOBJECT_CONTROL").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").contextMenu
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").contextMenu
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_DESCR"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").contextMenu
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_ICON"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[1]/tbar[0]/btn[0]").press
