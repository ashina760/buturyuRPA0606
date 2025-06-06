from datetime import datetime
import csv
import os

def log_process_result(log_path, new_file_path, new_folder_path, save_path=None, name=None, errors=None, result=None, new_csv_path=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 安全处理各参数
    save_path = save_path or ""
    name = name or ""
    errors = errors or {}
    result = result or {}
    new_csv_path = new_csv_path or ""

    # 把 errors 中的内容拉平
    if isinstance(errors, dict):
        error_messages = []
        for key, value in errors.items():
            if isinstance(value, dict):
                for sub_key, sub_value in value.items():
                    error_messages.append(f"{key}:{sub_key}={sub_value}")
            else:
                error_messages.append(f"{key}={value}")
        error_string = "; ".join(error_messages)
    else:
        error_string = str(errors)

    validation_success = "No" if errors else "Yes"

    log_data = {
        "Timestamp": timestamp,
        "FilePath": new_file_path,
        "FolderPath": new_folder_path,
        "SavePath": save_path,
        "Name": name,
        "ValidationSuccess": validation_success,
        "ValidationErrors": error_string,
        "UploadSuccess": "" if validation_success == "No" else result.get("success", ""),
        "UploadError": "" if validation_success == "No" else result.get("error", "") if not result.get("success", True) else "",
        "NewCsvPath": "" if validation_success == "No" else new_csv_path,
    }

    # 替换所有值中的逗号为一个空格
    log_data_cleaned = {k: (str(v).replace(",", " ") if v is not None else "") for k, v in log_data.items()}

    file_exists = os.path.isfile(log_path)
    with open(log_path, "a", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=log_data_cleaned.keys())
        if not file_exists:
            writer.writeheader()
        writer.writerow(log_data_cleaned)
