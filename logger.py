import os
import datetime
import inspect

log_initialized = False
current_date = datetime.datetime.now().strftime("%Y.%m.%d")
LOG_FILE_PATH = os.path.join(os.getenv('APPDATA'), "OfficeDispatch", current_date, f'{current_date}.log')

def log_message(message, log_file_path=LOG_FILE_PATH):
    global log_initialized
    log_dir = os.path.dirname(log_file_path)
    os.makedirs(log_dir, exist_ok=True)

    if not log_initialized:
        if os.path.exists(log_file_path) and os.path.getsize(log_file_path) > 0:
            with open(log_file_path, 'a', encoding='utf-8') as log:
                log.write('\n')
        log_initialized = True

    caller_frame = inspect.stack()[1]
    caller_filename = os.path.basename(caller_frame.filename)
    
    formatted_message = f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [od.module.{caller_filename}] {message}"
    print(formatted_message)
    with open(log_file_path, "a", encoding='utf-8') as log_file:
        log_file.write(f"{formatted_message}\n")
