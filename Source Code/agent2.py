# agent.py
import win32pipe
import win32file
import pywintypes
import os
import time
import psutil
import subprocess
import threading
import re
from ctypes import *
import logging

# Track executing files/commands
executing_commands = {}
executing_processes = {}  # Track process objects

PIPE_NAME = r'\\.\pipe\phone_control_pipe'

WORD_PATH = r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
POWERPOINT_PATH = r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"

def is_sms_recent(sms_date):
    """Check if SMS date is within 10 seconds of current time."""
    current_time_ms = int(time.time() * 1000)  # Convert to milliseconds
    time_diff = current_time_ms - sms_date
    if time_diff <= 7000:
        return True
    return False


def start_app(path, fallback=None):
    try:
        if path and os.path.exists(path):
            subprocess.Popen([path], shell=False)
            return True
        elif fallback:
            subprocess.Popen([fallback], shell=False)
            return True
    except:
        pass
    return False


def get_office_apps_running():
    """Check if Word or PowerPoint applications are currently running with visible windows."""
    office_apps = set()  # Use set to avoid duplicates
    try:
        for proc in psutil.process_iter(['pid', 'name', 'status']):
            try:
                # Only check running processes
                if proc.status() != psutil.STATUS_RUNNING:
                    continue

                proc_name = proc.name()
                if proc_name:
                    proc_name_lower = proc_name.lower()
                    if proc_name_lower in ['winword.exe', 'powerpnt.exe']:
                        # Additional check: ensure the process has a visible window
                        # This helps filter out background processes without UI
                        has_window = False
                        try:
                            # Check if process has any windows (indicates user-facing app)
                            import win32gui
                            import win32process

                            def callback(hwnd, hwnds):
                                if win32gui.IsWindowVisible(hwnd):
                                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                    if pid == proc.pid:
                                        hwnds.append(hwnd)

                            hwnds = []
                            win32gui.EnumWindows(callback, hwnds)

                            if hwnds:  # Process has visible windows
                                has_window = True
                                logging.info(f"📋 Found Office app with window: {proc_name} (PID: {proc.pid})")
                            else:
                                logging.info(f"🚫 Skipping Office process without window: {proc_name} (PID: {proc.pid})")

                        except ImportError:
                            # If win32gui not available, fall back to basic check
                            has_window = True
                            logging.info(f"📋 Found Office app (no window check): {proc_name} (PID: {proc.pid})")
                        except Exception as e:
                            logging.info(f"⚠️ Window check failed for {proc_name}: {e}")
                            has_window = True  # Default to including if check fails

                        if has_window:
                            office_apps.add(proc_name)  # Add to set to avoid duplicates

            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    except Exception as e:
        logging.info(f"❌ Error checking office apps: {e}")
    return list(office_apps)  # Convert back to list for compatibility


def show_shutdown_popup():
    """Show a popup message to close Office applications."""
    try:
        # Using Windows API to show popup
        title = "System Shutdown Warning"
        message = "System is shutting down. Please close Word and PowerPoint applications."
        result = windll.user32.MessageBoxW(0, message, title, 4)  # 4 = Yes/No buttons
        return result
    except:
        logging.info("⚠️ Shutdown warning: Please close Word and PowerPoint applications.")


def wait_for_office_apps_close():
    """Wait until Word and PowerPoint applications are closed."""
    while True:
        running_office_apps = get_office_apps_running()

        if not running_office_apps:
            logging.info("✅ Word and PowerPoint applications closed.")
            shutdown_system()
            return True

        logging.info(f"⏳ Waiting for Office apps to close: {running_office_apps}")
        time.sleep(2)


def shutdown_system():
    """Shutdown the system."""
    logging.info("🔴 Shutting down system in 10 seconds...")
    time.sleep(10)
    try:
        # Windows shutdown command
        os.system("shutdown /s /t 1")
    except Exception as e:
        logging.info(f"❌ Error during shutdown: {e}")


def read_sms():
    try:
        command = 'adb shell "content query --uri content://sms/inbox"'
        
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=True
        )

        output = result.stdout.decode("utf-8", errors="ignore")

        # logging.info("ADB OUTPUT (first 500 chars):\n", output[:500], "\n")

        if not output.strip():
            return "No SMS found or empty output.", 0

        rows = output.split("Row: ")
        latest_sms = None
        latest_date = 0

        for row in rows:
            date_match = re.search(r"date=(\d+)", row)
            body_match = re.search(r"body=(.*?)(?=, \w+=)", row, re.DOTALL)

            if date_match and body_match:
                sms_date = int(date_match.group(1))
                sms_body = body_match.group(1).strip()

                if sms_date > latest_date:
                    latest_date = sms_date
                    latest_sms = sms_body
        
        logging.info("Latest SMS:", latest_sms)
        logging.info("Latest SMS Date:", latest_date)
        return latest_sms if latest_sms else "No SMS body found.", latest_date


    except Exception as e:
        logging.info("ADB SMS read error:", e)
        return "", 0


def sms_poller():
    global executing_commands, executing_processes
    # seen_texts = set()

    while True:
        logging.info("🔎 Polling SMS...")
        sms_text, sms_date = read_sms()    

        if  sms_text == "open word" :
                # Check if WINWORD is already running
                if not is_sms_recent(sms_date):
                    logging.info("❌ SMS is not recent. Skipping command.")
                    time.sleep(5)
                    continue
                else:                
                    logging.info("📌 Command detected: open_word")
                    try:
                        process = subprocess.Popen([WORD_PATH])
                        executing_processes["open word"] = process
                        logging.info("📌 Command executed: open_word")
                    except Exception as e:
                        logging.info(f"❌ Error executing open word: {e}")

        elif sms_text == "open ppt":
                # Check if POWERPNT is already running
                if not is_sms_recent(sms_date):
                    logging.info("❌ SMS is not recent. Skipping command.")
                    time.sleep(5)
                    continue
                else:
                    logging.info("📌 Command detected: open_ppt")
                    try:
                        process = subprocess.Popen([POWERPOINT_PATH])
                        executing_processes["open ppt"] = process
                        logging.info("📌 Command executed: open_ppt")
                    except Exception as e:
                        logging.info(f"❌ Error executing open ppt: {e}")

        elif sms_text == "shutdown":
                if not is_sms_recent(sms_date):
                    logging.info("❌ SMS is not recent. Skipping command.")
                    time.sleep(5)
                    continue
                else:
                    logging.info("🔴 Shutdown command detected")
                    # Check if Word or PowerPoint are running
                    running_office_apps = get_office_apps_running()
                    if running_office_apps:
                        logging.info(f"📋 Office apps running: {running_office_apps}")
                        # Show popup to close Office applications
                        show_shutdown_popup()
                        # Wait for Office applications to close
                        wait_for_office_apps_close()
                    else:
                        logging.info("✅ No Office applications running. Proceeding with shutdown.")
                    # Shutdown system
                    shutdown_system()
        elif sms_text == "open ppt, open word, shutdown" or sms_text == "open word, open ppt, shutdown":
                if not is_sms_recent(sms_date):
                    logging.info("❌ SMS is not recent. Skipping command.")
                    time.sleep(5)
                    continue
                else:
                    logging.info("🔴 Combined command detected")
                    # Open PowerPoint
                    try:
                        process_ppt = subprocess.Popen([POWERPOINT_PATH])
                        executing_processes["open ppt"] = process_ppt
                        logging.info("📌 Command executed: open_ppt")
                    except Exception as e:
                        logging.info(f"❌ Error executing open ppt: {e}")
                    # Open Word
                    try:
                        process_word = subprocess.Popen([WORD_PATH])
                        executing_processes["open word"] = process_word
                        logging.info("📌 Command executed: open_word")
                    except Exception as e:
                        logging.info(f"❌ Error executing open word: {e}")
                    # Check if Word or PowerPoint are running
                    running_office_apps = get_office_apps_running()
                    if running_office_apps:
                        logging.info(f"📋 Office apps running: {running_office_apps}")
                        # Show popup to close Office applications
                        show_shutdown_popup()
                        # Wait for Office applications to close
                        wait_for_office_apps_close()
                    else:
                        logging.info("✅ No Office applications running. Proceeding with shutdown.")
                    # Shutdown system
                    shutdown_system()
        elif sms_text == "closed word":
                if not is_sms_recent(sms_date):
                    logging.info("❌ SMS is not recent. Skipping command.")
                    time.sleep(5)
                    continue
                else:
                    if "open word" in executing_processes:
                        process = executing_processes["open word"]
                        process.terminate()
                        del executing_processes["open word"]
                        logging.info("📌 Command executed: closed_word")
                    else:
                        logging.info("❌ No running process found for closed_word")
        elif sms_text == "closed ppt":
                if not is_sms_recent(sms_date):
                    logging.info("❌ SMS is not recent. Skipping command.")
                    time.sleep(5)
                    continue
                else:
                    if "open ppt" in executing_processes:
                        process = executing_processes["open ppt"]
                        process.terminate()
                        del executing_processes["open ppt"]
                        logging.info("📌 Command executed: closed_ppt")
                    else:
                        logging.info("❌ No running process found for closed_ppt")
        time.sleep(5)

def run_pipe_server():
    while True:
        try:
            pipe = win32pipe.CreateNamedPipe(
                PIPE_NAME,
                win32pipe.PIPE_ACCESS_DUPLEX,
                win32pipe.PIPE_TYPE_MESSAGE |
                win32pipe.PIPE_READMODE_MESSAGE |
                win32pipe.PIPE_WAIT,
                1, 65536, 65536, 0, None
            )

            logging.info("\n🔌 Waiting for service to connect...")
            win32pipe.ConnectNamedPipe(pipe, None)
            logging.info("✅ Service connected.")

            while True:
                try:
                    result, data = win32file.ReadFile(pipe, 64 * 1024)
                    text = data.decode('utf-8')
                    logging.info("📩 Received:", text)

                    cmds = text.split(";")
                    # handle_commands(cmds)

                except pywintypes.error:
                    break
                except Exception as e:
                    logging.info("Pipe read error:", e)
                    break

            try:
                win32file.CloseHandle(pipe)
            except:
                pass

        except Exception as e:
            logging.info("Pipe server error:", e)
            time.sleep(2)



if __name__ == "__main__":
    threading.Thread(target=sms_poller, daemon=True).start()
    run_pipe_server()
