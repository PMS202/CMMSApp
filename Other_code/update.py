import sys
import shutil
import os
import time
import subprocess
from dotenv import set_key
from PyQt5.QtWidgets import QApplication, QMessageBox

def update_application(download_link: str, target_folder: str, latest_version: str):
    time.sleep(2)

    temp_zip_path = os.path.join(os.getcwd(), "update_temp.zip")
    extract_folder = os.path.join(os.getcwd(), "update_temp")

    try:
        source_path = download_link.strip('"').strip("'")
        if os.path.exists(source_path):
            shutil.copy2(source_path, temp_zip_path)
        else:
            raise FileNotFoundError(f"Cannot find update file at: {source_path}")

        shutil.unpack_archive(temp_zip_path, extract_folder)

        extracted_items = [os.path.join(extract_folder, name) for name in os.listdir(extract_folder)]
        extracted_dir = None
        if len(extracted_items) == 1 and os.path.isdir(extracted_items[0]):
            extracted_dir = extracted_items[0]

        if os.path.exists(target_folder):
            shutil.rmtree(target_folder, ignore_errors=True)

        if extracted_dir:
            shutil.move(extracted_dir, target_folder)
        else:
            os.makedirs(target_folder, exist_ok=True)
            for item in extracted_items:
                shutil.move(item, os.path.join(target_folder, os.path.basename(item)))
        for name in os.listdir(target_folder):
            path = os.path.join(target_folder, name)
            if os.path.isdir(path) and name.lower().startswith("update_v"):
                shutil.rmtree(path, ignore_errors=True)
        time.sleep(1)
        
        root_folder = os.path.dirname(target_folder)
        env_path = os.path.join(root_folder, ".env")
        if not os.path.exists(env_path):
            env_path = os.path.join(target_folder, ".env")
        set_key(env_path, "APP_VERSION", latest_version)

        app_executable = os.path.join(target_folder, "CMMSApp.exe")
        if os.path.exists(app_executable):
            subprocess.Popen([app_executable], cwd=target_folder)
        else:
            raise FileNotFoundError(f"App not found at: {app_executable}")

        sys.exit(0)

    except Exception as e:
        app = QApplication(sys.argv)
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle("Update Failed")
        msg_box.setText(f"An error occurred during the update: {e}")
        msg_box.exec_()
        sys.exit(1)
    finally:
        try:
            if os.path.exists(temp_zip_path):
                os.remove(temp_zip_path)
            if os.path.exists(extract_folder):
                shutil.rmtree(extract_folder, ignore_errors=True)
        except:
            pass


if __name__ == "__main__":
    if len(sys.argv) >= 4:
        link = sys.argv[1]
        folder = sys.argv[2]
        ver = sys.argv[3]
        update_application(link, folder, ver)
    else:
        app = QApplication(sys.argv)
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle("Update Failed")
        msg_box.setText(f"Usage: update.exe <download_link> <target_folder> <version>")
        msg_box.exec_()
        sys.exit(1)