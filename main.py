#pip install pywin32
import win32com.client

def get_client():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    if not isinstance(sap_gui_auto, win32com.client.CDispatch):
        return None

    application = sap_gui_auto.GetScriptingEngine
    if not isinstance(application, win32com.client.CDispatch):
        sap_gui_auto = None
        return None

    for conn in range(application.Children.Count):
        connection = application.Children(conn)

        for sess in range(connection.Children.Count):
            session = connection.Children(sess)

            # Check if the session is at the main screen 'SESSION_MANAGER'
            if session.Info.Transaction == 'SESSION_MANAGER':
                return session
    return None

def main():
    # Initialize an object to store the SAP session
    session = get_client()

import win32com.client

def run_sap_script():
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)  # Suponiendo que ya hay una conexión activa
    session = connection.Children(0)  # Suponiendo que es la primera sesión abierta
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "iw39"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkDY_MAB").selected = True 
    session.findById("wnd[0]/usr/chkDY_HIS").selected = True 
    session.findById("wnd[0]/usr/ctxtAUART-LOW").text = "ZC04"
    session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    session.findById("wnd[0]/usr/chkDY_HIS").setFocus
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "AUFNR"
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell

if __name__ == '__main__':
    run_sap_script()


if __name__ == '__main__':
    main()
