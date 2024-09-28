import win32com.client
import xlwings as xw
import os
import shutil
import datetime
import time
import hashlib
import json
from github import Github

DEFAULT_CONFIG = {
    "repo_name": "your_repo_name",
    "github_token": "enter_your_github_token_here",
    "retry_interval": 10,
    "check_interval": 30
}

CONFIG_PATH = 'config.json'
LOG_FILE_PATH = 'log.txt'

def log_message(message, log_file_path=LOG_FILE_PATH):
    if not os.path.exists(log_file_path):
        with open(log_file_path, 'w', encoding='utf-8') as log_file:
            log_file.write('')
    with open(log_file_path, "a", encoding='utf-8') as log_file:
        log_file.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

def validate_config_value(key, value):
    if key in ["retry_interval", "check_interval"]:
        return isinstance(value, int) and value > 0
    elif key in ["repo_name", "github_token"]:
        return isinstance(value, str) and len(value) > 0
    return True

def load_config():
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
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except (json.JSONDecodeError, ValueError) as e:
            log_message(f'Error loading config: {e}. Resetting invalid entries to default values.')
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
    else:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    return config

config = load_config()

def initialize_com_object(app_name):
    while True:
        try:
            app = win32com.client.Dispatch(app_name)
            return app
        except Exception as e:
            print(f"Error initializing {app_name}: {e}. Retrying in {config['retry_interval']} seconds...")
            time.sleep(config['retry_interval'])

ppt = initialize_com_object("PowerPoint.Application")
word = initialize_com_object("Word.Application")

appdata_path = os.getenv('APPDATA')
repo_path = os.path.join(appdata_path, "OfficeDispatch")

os.makedirs(repo_path, exist_ok=True)

processed_files = set()
log_initialized = False

def calculate_md5(file_path):
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def copy_file(file_name, file_size, category, target_folder, log_file_path):
    try:
        dest_folder = os.path.join(target_folder, category)
        dest_file_path = os.path.join(dest_folder, os.path.basename(file_name))
        
        if os.path.exists(dest_file_path) and calculate_md5(file_name) == calculate_md5(dest_file_path):
            log_message(f"File {file_name} already exists and is identical, skipping copy", log_file_path)
            return False
        
        shutil.copy(file_name, dest_folder)
        log_message(f"Copied {file_name} to {dest_folder} successfully", log_file_path)
        return True
    except Exception as e:
        log_message(f"Failed to copy {file_name} due to {e}", log_file_path)
        return False

def process_files():
    log_file_path = None
    new_files = []
    try:
        current_date = datetime.datetime.now().strftime("%Y.%m.%d")
        target_folder = os.path.join(repo_path, current_date)

        os.makedirs(os.path.join(target_folder, "PowerPoint"), exist_ok=True)
        os.makedirs(os.path.join(target_folder, "Excel"), exist_ok=True)
        os.makedirs(os.path.join(target_folder, "Word"), exist_ok=True)

        log_file_path = os.path.join(target_folder, "log.txt")

        if ppt:
            for presentation in ppt.Presentations:
                file_name = presentation.FullName
                if file_name not in processed_files:
                    processed_files.add(file_name)
                    file_size = os.path.getsize(file_name)
                    if copy_file(file_name, file_size, "PowerPoint", target_folder, log_file_path):
                        new_files.append(os.path.join("PowerPoint", os.path.basename(file_name)))

        try:
            for book in xw.books:
                file_name = book.fullname
                if file_name not in processed_files:
                    processed_files.add(file_name)
                    file_size = os.path.getsize(file_name)
                    if copy_file(file_name, file_size, "Excel", target_folder, log_file_path):
                        new_files.append(os.path.join("Excel", os.path.basename(file_name)))
        except Exception as e:
            print(f"No Excel instances found: {e}")

        if word:
            for document in word.Documents:
                file_name = document.FullName
                if file_name not in processed_files:
                    processed_files.add(file_name)
                    file_size = os.path.getsize(file_name)
                    if copy_file(file_name, file_size, "Word", target_folder, log_file_path):
                        new_files.append(os.path.join("Word", os.path.basename(file_name)))

        return new_files, target_folder, log_file_path
    except Exception as e:
        if log_file_path:
            log_message(f"Error in processing files: {e}", log_file_path)
        return [], None, None

def upload_to_github(repo_name, token, target_folder, new_files, log_file_path):
    try:
        g = Github(token)
        user = g.get_user()
        repo = user.get_repo(repo_name)
        
        commit_files = []
        for rel_path in new_files:
            file_path = os.path.join(target_folder, rel_path)
            with open(file_path, "rb") as file:
                content = file.read()
            
            full_rel_path = os.path.join(os.path.basename(target_folder), rel_path).replace("\\", "/")
            commit_message = f"Update {os.path.basename(file_path)}"
            try:
                existing_file = repo.get_contents(full_rel_path)
                if existing_file.sha != hashlib.sha1(content).hexdigest():
                    repo.update_file(
                        path=full_rel_path,
                        message=commit_message,
                        content=content,
                        sha=existing_file.sha,
                        branch="main"
                    )
                    commit_files.append(full_rel_path)
            except:
                repo.create_file(
                    path=full_rel_path,
                    message=commit_message,
                    content=content,
                    branch="main"
                )
                commit_files.append(full_rel_path)
        
        if commit_files:
            log_message(f"Files uploaded to GitHub successfully: {', '.join(commit_files)}", log_file_path)
        else:
            log_message("No changes detected, no files uploaded", log_file_path)
    except Exception as e:
        log_message(f"Failed to upload files to GitHub due to {e}", log_file_path)

while True:
    new_files, target_folder, log_file_path = process_files()
    if new_files:
        upload_to_github(config['repo_name'], config['github_token'], target_folder, new_files, log_file_path)
    time.sleep(config['check_interval'])
