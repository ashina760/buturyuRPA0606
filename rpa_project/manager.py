from watcher.excel_file_watcher import ExcelFileWatcher
from excel_handler.processor import ExcelProcessor
from web_automation.automator import AeonUploader
from excel_handler.workflow import validate_excel_data, generate_upload_data, match_and_fill_from_csv,move_csv_to_folder,get_latest_file
from ledger.log import log_process_result
from settings import DOWNLOADS_PATH
# 第一步：监视文件夹
watcher = ExcelFileWatcher()
print("📂 正在持续监听文件夹...")

while True:
    new_file_path, new_folder_path = watcher.wait_for_new_file()
    print("✅ 检测到并移动了文件")

    a = ExcelProcessor(new_file_path)

    try:
        # 第二步：校验excel数据
        errors = validate_excel_data(a)
        cells = ["L6"]
        name = a.get_cell_values_from_workbook(cells)
        # print("📄 名称:", name)
        if errors:
            print("❌ 校验失败，原因：", errors)
            result["success"] = False
        else:
            # 第三步：生成nagashikomi数据

            save_path = generate_upload_data(a, new_folder_path)
            print("✅ 流しデータ生成完毕")

            # 第四步：上传数据到 Web
            uploader = AeonUploader()
            result = uploader.run(save_path)
            if result["success"]:
                print("✅开始填充发注番号")
                # 第五步：从 CSV 匹配并填充
                csv_path = match_and_fill_from_csv(processor=a)
                a.save()
                print(f"💾 文件已保存")
                new_csv_path = move_csv_to_folder(csv_path, new_folder_path)
                # print(f"{csv_path} 已成功移动到 {new_folder_path}")
            if result["inputEl"]:
                print("❌ 投入ERR")
                csv_path = get_latest_file(DOWNLOADS_PATH)
                new_csv_path = move_csv_to_folder(csv_path, new_folder_path)

            else:
                print("❌ 上传失败，原因：", result["error"])

    except Exception as e:
        print("⚠️ 处理流程出错:", str(e))

    finally:
        a.close()
        log_process_result(
            log_path=r"C:\Users\rp4-bpo\Box\70.（BPO）本社効率化PT\Wave1　(0605本番稼働)\01　資材チーム\◆07.物流G\log.csv",
            new_file_path=new_file_path,
            new_folder_path=new_folder_path,
            save_path=locals().get("save_path"),
            name=locals().get("name"),
            errors=locals().get("errors"),
            result=locals().get("result"),
            new_csv_path = locals().get("new_csv_path"),
        )
        print("📄 文件处理完毕，继续监听中...\n")
