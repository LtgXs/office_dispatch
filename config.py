import os
import sys
import json
import base64
import hashlib
import logger
from cryptography.fernet import Fernet

DEFAULT_CONFIG = {
    "fetch_interval": 60,
    "json_file_path": "path/to/your/command.json",
    "default_screenshot_path": "C:/path/to/default_screenshot_directory",
    "repo_name": "your_repo_name",
    "github_token": "enter_your_github_token_here",
    "retry_interval": 10,
    "check_interval": 30,
    "RemoteExc": {"enabled": False, "interval": 120}
}

if getattr(sys, 'frozen', False): # For pyinstaller
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(script_dir, "config.json")
PASSWORD = "enter_your_password_here"
key = base64.urlsafe_b64encode(hashlib.sha256(PASSWORD.encode()).digest())
cipher_suite = Fernet(key)

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
    logger.log_message('Loading Config...')
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
                        logger.log_message(f'Invalid value for {key}: {user_config[key]}. Resetting to default value.')
                else:
                    logger.log_message(f'Missing key {key}. Resetting to default value.')
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
                            logger.log_message("Decrypted token does not contain 'ghp_'. Terminating script.")
                            raise ValueError("Decryption failed")
                    except Exception as e:
                        logger.log_message(f"Failed to decrypt GitHub token: {e}. Terminating script.")
                        raise ValueError("Decryption failed")
        except (json.JSONDecodeError, ValueError) as e:
            logger.log_message(f'Error loading config: {e}. Resetting invalid entries to default values.')
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
    else:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            logger.log_message(f'Creating new config file: {CONFIG_PATH}')
            json.dump(config, f, ensure_ascii=False, indent=4)
    logger.log_message('Config loaded successfully.')
    return config