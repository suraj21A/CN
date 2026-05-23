import os
import time
import shutil
import traceback
from datetime import datetime

from MARUTI_GUJARAT_CN import process_maruti_cn

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR =os.path.dirname(TOOL_DIR)
RUN_DIR = os.path.join(ROOT_DIR,"RUN")
INPUT_DIR = os.path.join(RUN_DIR, "INPUT")
PROCESSED_DIR = os.path.join(RUN_DIR, "ProcessedInput")
ERROR_DIR = os.path.join(TOOL_DIR, "ErrorInput")
LOG_FILE = os.path.join(TOOL_DIR, "watch_log.txt")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

ALLOWED_EXT = (".xlsx", ".xls")


def log(msg):
    text = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(text)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def is_file_ready(file_path):
    try:
        size1 = os.path.getsize(file_path)
        time.sleep(2)
        size2 = os.path.getsize(file_path)
        return size1 == size2 and size1 > 0
    except Exception:
        return False


def move_file(src, folder):
    os.makedirs(folder, exist_ok=True)

    base = os.path.basename(src)
    name, ext = os.path.splitext(base)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(folder, f"{name}_{timestamp}{ext}")

    shutil.move(src, dest)
    return dest


def process_file(file_path):
    log("=" * 60)
    log(f"New input file detected: {file_path}")

    if not is_file_ready(file_path):
        log("File is not ready yet. Skipping for now.")
        return

    try:
        result = process_maruti_cn(file_path)

        if result:
            moved_path = move_file(file_path, PROCESSED_DIR)
            log(f"Output generated: {result}")
            log(f"Processed input moved to: {moved_path}")
        else:
            moved_path = move_file(file_path, ERROR_DIR)
            log(f"No output generated. File moved to error folder: {moved_path}")

    except Exception as e:
        log(f"Error processing file: {file_path}")
        log(str(e))
        log(traceback.format_exc())

        try:
            moved_path = move_file(file_path, ERROR_DIR)
            log(f"Error input moved to: {moved_path}")
        except Exception as move_err:
            log(f"Could not move error file: {move_err}")


def main():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(PROCESSED_DIR, exist_ok=True)
    os.makedirs(ERROR_DIR, exist_ok=True)

    log("TATA watch_input started")
    log(f"Watching folder: {INPUT_DIR}")

    processed_seen = set()

    while True:
        try:
            files = [
                os.path.join(INPUT_DIR, f)
                for f in os.listdir(INPUT_DIR)
                if f.lower().endswith(ALLOWED_EXT)
                and not f.startswith("~$")
            ]

            for file_path in files:
                if file_path not in processed_seen:
                    processed_seen.add(file_path)
                    process_file(file_path)

            time.sleep(5)

        except KeyboardInterrupt:
            log("Watcher stopped by user.")
            break

        except Exception as e:
            log(f"Watcher error: {e}")
            log(traceback.format_exc())
            time.sleep(5)


if __name__ == "__main__":
    main()