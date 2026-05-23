import os
import time
import shutil
from Hundai_export import process_hyundai_cn

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR =os.path.dirname(TOOL_DIR)
RUN_DIR = os.path.join(ROOT_DIR,"RUN")
INPUT_DIR = os.path.join(RUN_DIR, "INPUT")
PROCESSED_INPUT_DIR = os.path.join(RUN_DIR, "ProcessedInput")
ERROR_DIR = os.path.join(TOOL_DIR, "ErrorInput")
LOG_FILE = os.path.join(TOOL_DIR, "watch_log.txt")


os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(PROCESSED_INPUT_DIR, exist_ok=True)

SUPPORTED_EXTENSIONS = (".xlsx", ".xls")


def is_excel_file(filename):
    return filename.lower().endswith(SUPPORTED_EXTENSIONS)


def get_unprocessed_files():
    files = []
    for f in os.listdir(INPUT_DIR):
        full_path = os.path.join(INPUT_DIR, f)
        if os.path.isfile(full_path) and is_excel_file(f):
            files.append(full_path)
    return files


def move_input_file(input_path):
    filename = os.path.basename(input_path)
    dest_path = os.path.join(PROCESSED_INPUT_DIR, filename)

    if os.path.exists(dest_path):
        base, ext = os.path.splitext(filename)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        dest_path = os.path.join(PROCESSED_INPUT_DIR, f"{base}_{timestamp}{ext}")

    shutil.move(input_path, dest_path)


def main():
    print("=" * 60)
    print("Watching Hyundai INPUT folder for new Excel files...")
    print(INPUT_DIR)
    print("=" * 60)

    processed_in_this_session = set()

    while True:
        try:
            files = get_unprocessed_files()

            for file_path in files:
                if file_path in processed_in_this_session:
                    continue

                print(f"New input file detected: {file_path}")

                try:
                    output_file = process_hyundai_cn(file_path)

                    if output_file:
                        print(f"Generated file: {output_file}")
                        move_input_file(file_path)
                        print(f"Moved processed input: {file_path}")
                    else:
                        print(f"No output generated for: {file_path}")

                    processed_in_this_session.add(file_path)

                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")

            time.sleep(5)

        except KeyboardInterrupt:
            print("Stopped watching.")
            break

        except Exception as e:
            print(f"Watcher error: {e}")
            time.sleep(5)


if __name__ == "__main__":
    main()