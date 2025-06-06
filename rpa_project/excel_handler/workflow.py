# workflow.py
from excel_handler.processor import ExcelProcessor
from excel_handler.utils import check_dates_in_dict, check_past_dates
from settings import (TITLE_COLUMNS, EXPECTED_TITLES,MANDATORY_CELLS, MANDATORY_COLUMN,DATE_COLUMN, ID_COLUMN,FILL_VALUES,MIN_COL, MAX_COL)
from settings import REFERENCE_PATH,KEY_COLUMNS_IN_A, KEY_COLUMNS_IN_B, VALUE_COLUMN_IN_B, TARGET_COLUMN_IN_A,DOWNLOADS_PATH
import pandas as pd
from openpyxl.utils import column_index_from_string
from excel_handler.utils import get_latest_file

def validate_excel_data(processor: ExcelProcessor) -> dict:
    """
    执行一系列校验，如标题、空单元格、历史日期等。
    如果发现错误，返回包含错误信息及详情的字典。
    """
    errors = {}

    if processor.has_multiple_sheets():
        errors["sheets"] = {"存在多个表"}

    title_result = processor.is_title_valid(TITLE_COLUMNS, expected_values=EXPECTED_TITLES)
    if title_result:
        errors["title"] = {title_result}

    processor.delete_empty_rows(MANDATORY_COLUMN)

    if processor.is_cell_empty(MANDATORY_CELLS):
        errors["cell_check"] = {"L6 为空白"}

    empty_cells = processor.find_empty_cells(MIN_COL, MAX_COL)
    if empty_cells:
        errors["empty_cells"] = empty_cells

    
    id_date_tuple = processor.get_column_dates_with_colD(DATE_COLUMN, ID_COLUMN)

    b = ExcelProcessor(REFERENCE_PATH)

    delivery_date_dict = b.get_column_based_dict()

    unmatched = check_dates_in_dict(id_date_tuple, delivery_date_dict)
    if unmatched:
        errors["找不到日期"] = {unmatched}

    past_dates = check_past_dates(id_date_tuple)
    if past_dates:
        errors["纳品日为过去日"] = {past_dates}

    b.close()
    
    return errors

def generate_upload_data(processor: ExcelProcessor, save_dir: str) -> str:
    """
    调用生成上传数据的函数，返回保存路径。
    """
    save_path = processor.create_upload_data(save_dir, FILL_VALUES)
    return save_path


def match_and_fill_from_csv(processor: ExcelProcessor):
    """
    在 A 表中，根据指定列组合 key，在 B (CSV) 表中查找匹配项，如果找到则将指定列的值写入 A 表目标列。
    
    参数:
        processor: ExcelProcessor 实例 (处理 A 表)
        csv_path: CSV 文件路径 (B 表)
        key_columns_in_a: List[str]，A 表中参与 key 的列名，如 ['C', 'D', 'E']
        key_columns_in_b: List[str]，CSV 中对应的列名，如 ['col1', 'col2', 'col3']
        value_column_in_b: str，CSV 中要写入 A 表的值所在的列名，如 'F'
        target_column_in_a: str，写入 A 表的目标列名，如 'M'
    """
    csv_path = get_latest_file(DOWNLOADS_PATH)

    df_b = pd.read_csv(csv_path, encoding="cp932", dtype=str).fillna("")  # 读取并填空字符串，避免 NaN 干扰
    df_b["key"] = df_b.apply(lambda row: build_clean_key(row, KEY_COLUMNS_IN_B), axis=1)
    key_value_dict = dict(zip(df_b["key"], df_b[VALUE_COLUMN_IN_B]))  # 假设你要记录的是 F 列的值
    processor.get_min_max_row()
    max_row = processor.max_row
    processor.convert_column_to_yyyymmdd("H")
    target_col_idx = column_index_from_string(TARGET_COLUMN_IN_A)
    # print(key_value_dict)
    # 遍历 A 表的每一行，构造 key 并匹配写入值
    for row in range(2, max_row + 1):  # 从第2行开始跳过表头
        key_parts = [str(processor.sheet[f"{col}{row}"].value).strip() for col in KEY_COLUMNS_IN_A]
        full_key = ''.join(key_parts)
        full_key = full_key.replace(" ", "").replace("\u3000", "").replace("\n", "")  # ✨ 清理空格
        # print(f"构造的 key: {full_key}")

        if full_key in key_value_dict:
            processor.sheet.cell(row=row, column=target_col_idx, value=key_value_dict[full_key])
    return csv_path
def build_clean_key(row, key_columns):
    def normalize(val):
        if pd.isna(val):
            return ""
        val = str(val).strip().replace('\u3000', '').replace('\n', '')
        try:
            if 'e' in val.lower():
                val = format(float(val), '.0f')  # 去除科学计数法
        except:
            pass
        return val
    return ''.join([normalize(row[col]) for col in key_columns])

import os
import shutil

def move_csv_to_folder(csv_path, new_folder_path):
    # 确保目标文件夹存在，不存在则创建
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)
    
    # 获取CSV文件的文件名
    file_name = os.path.basename(csv_path)
    
    # 构建目标路径
    new_path = os.path.join(new_folder_path, file_name)
    
    # 移动文件
    shutil.move(csv_path, new_path)

    return new_path
    
    