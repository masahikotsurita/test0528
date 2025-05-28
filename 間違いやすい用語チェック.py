import os
import sys
import datetime
from docx import Document
from docx.oxml.ns import qn
from openpyxl import load_workbook
from pptx import Presentation
from pptx.table import Table
import configparser

# iniファイルからkeywordsとreplacementsを読み込む
config = configparser.ConfigParser()
# exe内かどうかを判断してiniの場所を動的に決定
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # PyInstaller実行環境
    ini_path = os.path.join(os.path.dirname(sys.executable), '間違いやすい用語チェック.ini')
else:
    base_path = os.path.dirname(__file__)
    ini_path = os.path.join(base_path, '間違いやすい用語チェック.ini')
config.read(ini_path, encoding='utf-8')

keywords = []
replacements = []

if 'Replacements' in config:
    for k, r in config['Replacements'].items():
        keywords.append(k)
        replacements.append(r)

log_path = os.path.join(os.environ['TEMP'], f"ReplaceLog_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

def log(message):
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"{message}\n")

def search_text_in_docx(path):
    log(f"――――　ファイル: {os.path.basename(path)}　――――")
    doc = Document(path)
    current_heading = "章番号不明"
    for para in doc.paragraphs:
        if para.style.name.startswith("Heading"):
            # 見出し段落の先頭番号を見出し番号として扱う（例: "1.1 概要" → "1.1"）
            split_text = para.text.strip().split()
            if split_text and any(char.isdigit() for char in split_text[0]):
                current_heading = split_text[0]
            else:
                current_heading = para.text.strip()
        for k, r in zip(keywords, replacements):
            if k in para.text:
                log(f"{current_heading}: '{k}' → '{r}'")

def search_text_in_xlsx(path):
    log(f"――――　ファイル: {os.path.basename(path)}　――――")
    wb = load_workbook(path)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for k, r in zip(keywords, replacements):
                        if k in cell.value:
                            log(f"シート'{sheet.title}' セル{cell.coordinate}: '{k}' → '{r}'")

def search_text_in_pptx(path):
    log(f"――――　ファイル: {os.path.basename(path)}　――――")
    prs = Presentation(path)
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for k, r in zip(keywords, replacements):
                    if k in shape.text:
                        log(f"スライド{i+1}: '{k}' → '{r}'")
            elif shape.has_table:
                table: Table = shape.table
                for row in table.rows:
                    for cell in row.cells:
                        for k, r in zip(keywords, replacements):
                            if k in cell.text:
                                log(f"スライド{i+1}: '{k}' → '{r}'")

def process_files(filepaths):
    for path in filepaths:
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".docx":
                search_text_in_docx(path)
                log("")
                log("")
            elif ext == ".xlsx":
                search_text_in_xlsx(path)
                log("")
                log("")
            elif ext == ".pptx":
                search_text_in_pptx(path)
                log("")
                log("")
            else:
                log(f"[SKIP] 未対応の拡張子: {path}")
        except Exception as e:
            log(f"[ERROR] {path} 処理中にエラー: {str(e)}")

    os.system(f"notepad.exe {log_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        input_files = sys.argv[1:]
        process_files(input_files)
    else:
        print("Office ファイル（docx / xlsx / pptx）をドロップして入力してください：")
        input_str = input("> ")
        files = [f.strip() for f in input_str.split(',') if os.path.isfile(f.strip())]
        if files:
            process_files(files)
        else:
            print("有効なファイルが指定されていません。プログラムを終了します。")
