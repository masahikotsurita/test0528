import os
import sys
import datetime
import configparser
import tkinter as tk

if os.name == 'nt':
    try:
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    except Exception:
        pass

# iniファイルからkeywordsとreplacementsを読み込む
config = configparser.ConfigParser()

# 実行ファイルの配置ディレクトリを基準とする
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)  # PyInstaller実行環境
else:
    base_path = os.path.dirname(__file__)

# PyInstallerの --add-data でバンドルされた ini も探索する
ini_candidates = [os.path.join(base_path, '間違いやすい用語チェック.ini')]
if getattr(sys, '_MEIPASS', None):
    ini_candidates.append(os.path.join(sys._MEIPASS, '間違いやすい用語チェック.ini'))

ini_path = None
for p in ini_candidates:
    if os.path.exists(p):
        ini_path = p
        break
if ini_path is None:
    ini_path = ini_candidates[0]

def load_replacements():
    config.read(ini_path, encoding='utf-8')
    keywords = []
    replacements = []
    if 'Replacements' in config:
        for k, r in config['Replacements'].items():
            keywords.append(k)
            replacements.append(r)
    return keywords, replacements

# ログファイルは exe と同じディレクトリに保存する
log_filename = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_ReplaceLog.txt"
log_path = os.path.join(base_path, log_filename)

def log(message):
    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f"{message}\n")

def edit_ini(path):
    if os.path.exists(path):
        config.read(path, encoding='utf-8')
    if 'Replacements' not in config:
        config['Replacements'] = {}

    root = tk.Tk()
    root.title('間違いやすい用語チェック.ini 編集')

    frame = tk.Frame(root)
    frame.pack(padx=10, pady=10)

    listbox = tk.Listbox(frame, width=50, height=15)
    listbox.grid(row=0, column=0, columnspan=3, sticky='nsew')
    scrollbar = tk.Scrollbar(frame, orient='vertical', command=listbox.yview)
    scrollbar.grid(row=0, column=3, sticky='ns')
    listbox.config(yscrollcommand=scrollbar.set)

    entry_key = tk.Entry(frame)
    entry_val = tk.Entry(frame)
    entry_key.grid(row=1, column=0, padx=5, pady=5)
    entry_val.grid(row=1, column=1, padx=5, pady=5)

    def refresh():
        listbox.delete(0, tk.END)
        for k, v in config['Replacements'].items():
            listbox.insert(tk.END, f'{k} = {v}')

    def on_select(event=None):
        if not listbox.curselection():
            return
        item = listbox.get(listbox.curselection()[0])
        k, v = item.split(' = ', 1)
        entry_key.delete(0, tk.END)
        entry_key.insert(0, k)
        entry_val.delete(0, tk.END)
        entry_val.insert(0, v)

    def on_add_update(event=None):
        k = entry_key.get().strip()
        v = entry_val.get().strip()
        if not k or not v:
            return
        config['Replacements'][k] = v
        refresh()
        entry_key.delete(0, tk.END)
        entry_val.delete(0, tk.END)

    def on_delete(event=None):
        if not listbox.curselection():
            return
        item = listbox.get(listbox.curselection()[0])
        k = item.split(' = ', 1)[0]
        config['Replacements'].pop(k, None)
        refresh()

    def on_save(event=None):
        with open(path, 'w', encoding='utf-8') as f:
            config.write(f)
        root.destroy()

    listbox.bind('<<ListboxSelect>>', on_select)
    root.bind('<Return>', on_add_update)
    root.bind('<Delete>', on_delete)
    root.bind('<Control-s>', on_save)

    tk.Button(frame, text='追加/更新', command=on_add_update).grid(row=1, column=2, padx=5)
    tk.Button(frame, text='削除', command=on_delete).grid(row=2, column=2, padx=5)
    tk.Button(frame, text='保存して終了', command=on_save).grid(row=3, column=2, padx=5)

    refresh()
    root.mainloop()

def search_text_in_docx(path, keywords, replacements):
    from docx import Document
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

def search_text_in_xlsx(path, keywords, replacements):
    from openpyxl import load_workbook
    log(f"――――　ファイル: {os.path.basename(path)}　――――")
    wb = load_workbook(path)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for k, r in zip(keywords, replacements):
                        if k in cell.value:
                            log(f"シート'{sheet.title}' セル{cell.coordinate}: '{k}' → '{r}'")

def search_text_in_pptx(path, keywords, replacements):
    from pptx import Presentation
    from pptx.table import Table
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
    keywords, replacements = load_replacements()
    # 処理対象のファイル一覧を冒頭に記録する
    names = ', '.join(os.path.basename(p) for p in filepaths)
    log(f"比較ファイル: {names}")
    log("")
    for path in filepaths:
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".docx":
                search_text_in_docx(path, keywords, replacements)
                log("")
                log("")
            elif ext == ".xlsx":
                search_text_in_xlsx(path, keywords, replacements)
                log("")
                log("")
            elif ext == ".pptx":
                search_text_in_pptx(path, keywords, replacements)
                log("")
                log("")
            else:
                log(f"[SKIP] 未対応の拡張子: {path}")
        except Exception as e:
            log(f"[ERROR] {path} 処理中にエラー: {str(e)}")

    os.system(f"notepad.exe {log_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and os.path.exists(ini_path):
        input_files = sys.argv[1:]
        process_files(input_files)
    else:
        edit_ini(ini_path)
        sys.exit(0)

