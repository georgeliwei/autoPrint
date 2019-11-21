import win32serviceutil
import win32service
import win32ts
import win32process
import win32con
import time


class AutoPrintService(win32serviceutil.ServiceFramework):
    _svc_name_ = "autoPrintService"
    _svc_display_name_ = "autoPrintService"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.runFlag = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.runFlag = False
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        self.runFlag = True
        period_time = 5 * 60
        while self.runFlag:
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            self.create_print_process()
            time.sleep(period_time)

    def create_print_process(self):
        session_id = win32ts.WTSGetActiveConsoleSessionId()
        user_token = win32ts.WTSQueryUserToken(session_id)
        startup_info = win32process.STARTUPINFO()
        cmdline = "E:\\python_project\\happyWorkAutoPrint\\dist\\autoPrint\\autoPrint.exe"
        win32process.CreateProcessAsUser(user_token,
                                         None,
                                         cmdline,
                                         None,
                                         None,
                                         False,
                                         win32con.NORMAL_PRIORITY_CLASS,
                                         None,
                                         None,
                                         startup_info)


if __name__ == "__main__":
    win32serviceutil.HandleCommandLine(AutoPrintService)
