import os
import time
import shutil
from hyundai_pune_export import process_hyundai_cn

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


def get_excel_files():
    files = []
    for f in os.listdir(INPUT_DIR):
        full_path = os.path.join(INPUT_DIR, f)
        if os.path.isfile(full_path) and is_excel_file(f):
            files.append(full_path)
    return files


def find_required_files():
    trpspn_file = None
    truck_file = None

    for file_path in get_excel_files():
        name = os.path.basename(file_path).lower()

        if "trpspn" in name and trpspn_file is None:
            trpspn_file = file_path
        elif "truck" in name and truck_file is None:
            truck_file = file_path

    return trpspn_file, truck_file


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
    print("Watching Hyundai INPUT folder for TRPSPN + TRUCK files...")
    print(INPUT_DIR)
    print("=" * 60)

    processed_pairs = set()

    while True:
        try:
            trpspn_file, truck_file = find_required_files()

            if trpspn_file and truck_file:
                pair_key = (
                    os.path.basename(trpspn_file).lower(),
                    os.path.basename(truck_file).lower()
                )

                if pair_key not in processed_pairs:
                    print("Required files found:")
                    print("TRPSPN:", trpspn_file)
                    print("TRUCK :", truck_file)

                    try:
                        output_file = process_hyundai_cn(trpspn_file, truck_file)

                        if output_file:
                            print(f"Generated file: {output_file}")
                            move_input_file(trpspn_file)
                            move_input_file(truck_file)
                            print("Moved processed input files.")
                        else:
                            print("No output generated.")

                        processed_pairs.add(pair_key)

                    except Exception as e:
                        print(f"Error processing files: {e}")

            else:
                if not trpspn_file:
                    print("Waiting for TRPSPN file...")
                if not truck_file:
                    print("Waiting for TRUCK file...")

            time.sleep(5)

        except KeyboardInterrupt:
            print("Stopped watching.")
            break

        except Exception as e:
            print(f"Watcher error: {e}")
            time.sleep(5)


if __name__ == "__main__":
    main()