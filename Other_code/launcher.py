import subprocess
import os
import sys
import time
from dotenv import load_dotenv

def get_base_dir():
    if getattr(sys, 'frozen', False):
        current_path = os.path.dirname(sys.executable)
    else:
        current_path = os.path.dirname(os.path.abspath(__file__))

    for _ in range(3):
        if os.path.exists(os.path.join(current_path, ".env")) or \
           os.path.exists(os.path.join(current_path, "updater")):
            return current_path
        
        parent = os.path.dirname(current_path)
        if parent == current_path: 
            break
        current_path = parent
    
    return os.path.dirname(sys.executable)

BASE_DIR = get_base_dir()

env_path = os.path.join(BASE_DIR, ".env")
load_dotenv(env_path)

try:
    from Database.MariaDB import Database_process
except ImportError as e:
    with open(os.path.join(BASE_DIR, "launcher_import_error.log"), "w") as f:
        f.write(str(e))
    sys.exit(1)

class CheckUpdate:
    def __init__(self, DB: Database_process):
        self.DB = DB
        self.current_version = os.getenv("APP_VERSION")

    def check_version(self):
        try:
            sql = "SELECT version,link FROM version_info ORDER BY release_date DESC LIMIT 1;"
            result = self.DB.query(sql)
            
            if not result:
                return "No_update", None
            
            latest_version, download_link = result[0]
            
            if latest_version != self.current_version:
                return latest_version, download_link
            return "No_update", None
            
        except Exception as e:
            return "No_update", None

if __name__ == "__main__":
    app_executable = os.path.join(BASE_DIR, "CMMSApp.dist", "CMMSApp.exe")
    updater_executable = os.path.join(BASE_DIR, "updater", "update.exe")
    target_folder = os.path.join(BASE_DIR, "CMMSApp.dist")

    try:
        db = Database_process()
        updater = CheckUpdate(db)
        latest_version, download_link = updater.check_version()

        if latest_version == "No_update" or latest_version is None:
            if os.path.exists(app_executable):
                subprocess.Popen([app_executable],cwd=os.path.dirname(app_executable))
            else:
                raise FileNotFoundError(f"App not found at: {app_executable}")
        else:
            if os.path.exists(updater_executable):
                subprocess.Popen([
                    updater_executable, 
                    download_link, 
                    target_folder, 
                    latest_version
                ], cwd=os.path.dirname(updater_executable))
            else:
                raise FileNotFoundError(f"Updater not found at: {updater_executable}")
                
        sys.exit(0)

    except Exception as e:
        log_path = os.path.join(BASE_DIR, "launcher_debug.log")
        with open(log_path, "w") as f:
            f.write(f"Error: {str(e)}\n")
            f.write(f"Detected Base Dir: {BASE_DIR}\n")
            f.write(f"Looking for Updater at: {updater_executable}\n")
            f.write(f"Looking for App at: {app_executable}\n")
        
        try:
            if os.path.exists(app_executable):
                subprocess.Popen([app_executable], cwd=os.path.dirname(app_executable))
        except:
            pass