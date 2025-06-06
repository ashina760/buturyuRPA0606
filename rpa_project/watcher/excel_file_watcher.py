import os
import time
from shutil import move
from settings import TIMESTAMP
from settings import WARCH_DIR,INTERVAL
class ExcelFileWatcher:
    def __init__(self, watch_dir = WARCH_DIR, interval = INTERVAL):
        self.watch_dir = watch_dir
        self.interval = interval
        self.processed_files = set()

    def wait_for_new_file(self):
        """
        等待新 Excel 文件并将其移动到带时间戳的新文件夹中。
        返回处理后的文件完整路径和新文件夹路径。
        """
        while True:
            current_files = {
                os.path.join(self.watch_dir, f)
                for f in os.listdir(self.watch_dir)
                if os.path.isfile(os.path.join(self.watch_dir, f)) and f.endswith(".xlsx")
            }

            new_files = current_files - self.processed_files

            if new_files:
                original_file_path = new_files.pop()
                new_file = os.path.basename(original_file_path)

                # 创建带时间戳的新文件夹
               
                folder_name = f"{TIMESTAMP} {new_file}"
                new_folder_path = os.path.join(self.watch_dir, folder_name)
                os.makedirs(new_folder_path, exist_ok=True)

                # 移动文件到新文件夹中
                new_file_path = os.path.join(new_folder_path, new_file)
                move(original_file_path, new_file_path)

                # print(f"已移动新文件: {new_file_path}")
                self.processed_files.add(new_file_path)

                return new_file_path, new_folder_path  # 返回文件路径和目录路径

            time.sleep(self.interval)

