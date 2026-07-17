# service.py
import win32serviceutil
import win32service
import win32event
import servicemanager
import time
import sqlite3
import os
import win32file
import pywintypes

DB_PATH = r"C:\Users\ASC\Desktop\Project\commands.db"
PIPE_NAME = r'\\.\pipe\phone_control_pipe'
POLL_SECONDS = 2
CONNECT_RETRY_SECONDS = 2
CONNECT_RETRIES = 5

class PhoneControlService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PhoneControlService"
    _svc_display_name_ = "Phone Control Service"
    _svc_description_ = "Service that reads DB and sends commands to desktop agent via named pipe."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.running = False
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogInfoMsg("PhoneControlService started.")
        self.main()

    def main(self):
        while self.running:
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT id, message FROM commands WHERE status='pending' ORDER BY id ASC")
                rows = c.fetchall()

                for row in rows:
                    id_, message = row
                    message_lower = (message or "").lower()

                    # Parse message into one or more commands
                    commands = []
                    if "open word" in message_lower:
                        commands.append("open_word")
                    if "open powerpoint" in message_lower or "open power point" in message_lower:
                        commands.append("open_ppt")
                    if "shutdown" in message_lower:
                        commands.append("shutdown")

                    if not commands:
                        # unknown command -> mark as done with note
                        c.execute("UPDATE commands SET status='done', note=? WHERE id=?", ("unknown command", id_))
                        conn.commit()
                        continue

                    payload = ";".join(commands).encode('utf-8')

                    sent = False
                    # Try to connect to agent pipe and send
                    for attempt in range(CONNECT_RETRIES):
                        try:
                            handle = win32file.CreateFile(
                                PIPE_NAME,
                                win32file.GENERIC_WRITE,
                                0, None,
                                win32file.OPEN_EXISTING,
                                0, None
                            )
                            win32file.WriteFile(handle, payload)
                            # Close handle
                            try:
                                handle.Close()
                            except:
                                pass
                            sent = True
                            break
                        except pywintypes.error as e:
                            # pipe not ready; wait and retry
                            time.sleep(CONNECT_RETRY_SECONDS)
                            continue
                        except Exception as ex:
                            time.sleep(CONNECT_RETRY_SECONDS)
                            continue

                    if sent:
                        c.execute("UPDATE commands SET status='done' WHERE id=?", (id_,))
                        conn.commit()
                        servicemanager.LogInfoMsg(f"Sent command(s): {commands} for DB id {id_}")
                    else:
                        # couldn't reach agent; leave as pending so retry next loop
                        servicemanager.LogWarningMsg(f"Could not reach agent to send commands for DB id {id_}; will retry.")

                conn.close()
                time.sleep(POLL_SECONDS)
            except Exception as e:
                servicemanager.LogErrorMsg("Service error: %s" % str(e))
                time.sleep(POLL_SECONDS)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(PhoneControlService)
