import win32com.client
import subprocess
import time
import sys

class sap_login:

    def login():

        try:
            path=r"C:\Users\ ProgramFilesX86\SAP\FrontEnd\SapGui\saplogon.exe"
            subprocess.Popen(path)
            subprocess.check_call([r'C:\ ProgramFilesX86\SAP\FrontEnd\SapGui\sapshcut.exe', '-system=PRD', '-client=100', '-user=SID', '-pw=PASSWORD'])
            time.sleep(10)
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return
                
            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return

            connection = application.OpenConnection("ConnectionName", True)
            if not type(connection) == win32com.client.CDispatch:
                application = None
                SapGuiAuto = None
                return

            session = connection.Children(0)
        
            if not type(session) == win32com.client.CDispatch:
                connection = None
                application = None
                SapGuiAuto = None
                return

            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "ID"
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "PASSWORD"
            session.findById("wnd[0]").sendVKey(0)
        except:
            print(sys.exc_info()[0])
        finally:
            session = None
            connection = None
            application = None
            SapGuiAuto = None
    login()