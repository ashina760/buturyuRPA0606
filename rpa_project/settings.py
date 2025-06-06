# settings.py
from datetime import datetime

INTERVAL = 20  # 秒
WARCH_DIR = r"C:\Users◆"

#配送可能日期excel路径
REFERENCE_PATH = r"C:\myenv\NPFKB.xlsx"
DOWNLOADS_PATH = r"D:\DATA\Downloads"

# 标题校验配置
TITLE_COLUMNS = ["C", "D", "E", "F", "G", "H", "I", "J", "K"]
EXPECTED_TITLES = [
    "仕入先コード", "入荷倉庫コード", "商品コード", "商品名（伝票用）",
    "発注数量", "納期", "発注単価", "発注金額", "伝票摘要"
]

# 空值检查配置
MANDATORY_CELLS = ["L6"]

# 用于删除空行检查的列
MANDATORY_COLUMN = "G"  

# 日期校验列配置
DATE_COLUMN = "H"
ID_COLUMN = "D"

# 上传excel配置
FILL_VALUES = {
    1: "D",
    5: "",
    6: "2",
    7: "9",
    8: "99"
}

#找出数据区域中 min_col~ max_col 列所有空单元格
MIN_COL = 3
MAX_COL = 11

# 创建带时间戳
TIMESTAMP = datetime.now().strftime("%Y%m%d-%H%M")


KEY_COLUMNS_IN_A = ["C", "D", "E", "G", "H", "K"]
KEY_COLUMNS_IN_B = ["仕入先コード", "センターコード", "商品コード", "発注数量", "指定納期", "伝票備考"]
VALUE_COLUMN_IN_B = "発注番号"
TARGET_COLUMN_IN_A = "M"


# Chrome 相关配置
CHROME_PATH = r"C:\chrome-win64\chrome.exe"
CHROMEDRIVER_PATH = r"C:\chromedriver-win64\chromedriver.exe"

# 登录信息
AEON_OPCD =  11
AEON_PASSWORD =11 