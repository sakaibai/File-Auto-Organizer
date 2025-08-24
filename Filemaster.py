import os
import shutil
import datetime
import re
import pandas as pd
from PyPDF2 import PdfReader
import csv

# === 設定 ===
SOURCE_DIR = "C:/Users/Yusuke Kori/Desktop/整理対象"
DEST_DIR = "C:/Users/Yusuke Kori/Desktop/整理済み"
RENAME_RULE = "{project}_{date}_{type}_{original}"
LOG_FILE = os.path.join(DEST_DIR, "整理ログ.csv")

# フォルダが存在しない場合は作成
os.makedirs(SOURCE_DIR, exist_ok=True)
os.makedirs(DEST_DIR, exist_ok=True)

# === プロジェクト名抽出関数 ===
def extract_project_name(file_path):
    if file_path.endswith(".pdf"):
        try:
            reader = PdfReader(file_path)
            text = reader.pages[0].extract_text()
            match = re.search(r"Project[:\s]+(\w+)", text)
            return match.group(1) if match else "UnknownProject"
        except:
            return "UnreadablePDF"
    elif file_path.endswith(".xlsx"):
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            for col in df.columns:
                if "Project" in col:
                    return str(df[col].iloc[0])
            return "UnknownProject"
        except:
            return "UnreadableExcel"
    else:
        return "Other"

# === ログ記録関数 ===
def log_result(original, new_name, project, date, ext, target_folder, processed_datetime):
    with open(LOG_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow([original, new_name, project, date, ext, target_folder, processed_datetime])
        
# === ログファイル初期化（ヘッダー追加） ===
if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(["Original Name", "New Name", "Project", "File Date", "Extension", "Target Folder", "Processed Datetime"])

# === ログ記録関数 ===
def log_result(original, new_name, project, date, ext, target_folder, processed_datetime):
    with open(LOG_FILE, mode='a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow([original, new_name, project, date, ext, target_folder, processed_datetime])
        
# === ファイル整理処理 ===
def organize_files():
    for filename in os.listdir(SOURCE_DIR):
        file_path = os.path.join(SOURCE_DIR, filename)
        if not os.path.isfile(file_path):
            continue

        ext = os.path.splitext(filename)[1][1:].lower()
        date = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y%m%d")
        project = extract_project_name(file_path)
        new_name = RENAME_RULE.format(
            project=project,
            date=date,
            type=ext,
            original=filename
        )

        # 保存先フォルダ作成
        target_folder = os.path.join(DEST_DIR, project, date)
        os.makedirs(target_folder, exist_ok=True)

        # ファイル移動＆リネーム
        shutil.move(file_path, os.path.join(target_folder, new_name))

        # 見える化ログ出力
        print(f"📦 処理対象: {filename}")
        print(f"📁 種類: {ext} | 📅 日付: {date} | 🏷 プロジェクト: {project}")
        print(f"➡ 新しいファイル名: {new_name}")
        print(f"✅ 移動先: {target_folder}")
        print("-" * 50)

        # 整理日時（現在時刻）を取得
        processed_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # CSVログ記録（引数を7つにする）
        log_result(filename, new_name, project, date, ext, target_folder, processed_datetime)

organize_files()
