import os
import time
import logger
import queue
import platform
import hashlib
import datetime
from github import Github

upload_queue = queue.Queue()

def get_hwid():
    processor = platform.processor()
    system_version = platform.version()
    machine = platform.machine()
    node = platform.node()
    hwid_source = f"{processor}_{system_version}_{machine}_{node}"
    hwid = hashlib.sha256(hwid_source.encode()).hexdigest()
    return hwid

def upload_files_to_github(repo_name, token, upload_queue, log_file_path, repo_path, retry_interval):
    g = Github(token)
    user = g.get_user()
    repo = user.get_repo(repo_name)


    while not upload_queue.empty():
        file_path = upload_queue.get()
        with open(file_path, "rb") as file:
            content = file.read()
        file_name = os.path.basename(file_path)

        common_path = os.path.commonpath([repo_path, file_path])
        full_rel_path = os.path.relpath(file_path, common_path).replace("\\", "/")
        file_name = os.path.basename(full_rel_path)

        commit_message = f"Upload {file_name}"
        commit_message_upd = f"Update {file_name}"

        attempt = 0
        success = False
        while attempt < retry_interval and not success:
            try:
                existing_file = repo.get_contents(full_rel_path)
                repo.update_file(
                    path=full_rel_path,
                    message=commit_message_upd,
                    content=content,
                    sha=existing_file.sha,
                    branch="main"
                )
                logger.log_message(f"Updated {full_rel_path}", log_file_path)
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
                        logger.log_message(f"Uploaded {full_rel_path}", log_file_path)
                        success = True
                    except Exception as create_e:
                        logger.log_message(f"Error uploading {full_rel_path}: {create_e}. Retrying...")
                        time.sleep(retry_interval)
                        attempt += 1
                else:
                    logger.log_message(f"Error updating {full_rel_path}: {e}. Retrying...")
                    time.sleep(retry_interval)
                    attempt += 1

        if not success:
            logger.log_message(f"Failed to process {full_rel_path} after multiple attempts.")
        upload_queue.task_done()

def check_and_rename_previous_logs(repo_path):
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