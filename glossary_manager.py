import os
import json
import datetime
import pandas as pd
from docx import Document
from typing import Dict, List, Optional, Tuple

class GlossaryManager:
    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.glossary_path = os.path.join(base_dir, "glossary", "glossary.json")
        self.data = {
            "meta": {
                "update_time": "",
                "count": 0,
                "version": "1.0.0"
            },
            "zh2en": {},
            "en2zh": {}
        }
        self.ensure_dirs()
        self.load()

    def ensure_dirs(self):
        dirs = [
            "glossary", "glossary/imports", "assets", "logs",
            "source_cn", "source_en", "output_zh", "output_en"
        ]
        for d in dirs:
            os.makedirs(os.path.join(self.base_dir, d), exist_ok=True)

    def load(self):
        if os.path.exists(self.glossary_path):
            try:
                with open(self.glossary_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except Exception as e:
                print(f"加载术语库失败: {e}")

    def save(self):
        self.data["meta"]["update_time"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.data["meta"]["count"] = len(self.data["zh2en"])
        with open(self.glossary_path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=4)

    def add_term(self, zh: str, en: str, overwrite: bool = True):
        zh, en = zh.strip(), en.strip()
        if not zh or not en: return
        if overwrite or zh not in self.data["zh2en"]:
            self.data["zh2en"][zh] = en
            self.data["en2zh"][en] = zh

    def delete_term(self, zh: str):
        if zh in self.data["zh2en"]:
            en = self.data["zh2en"].pop(zh)
            self.data["en2zh"].pop(en, None)

    def import_excel(self, file_path: str, overwrite: bool = True):
        df = pd.read_excel(file_path)
        # 自动识别列
        zh_col, en_col = None, None
        for col in df.columns:
            col_lower = str(col).lower()
            if any(k in col_lower for k in ["中", "zh", "chinese"]): zh_col = col
            if any(k in col_lower for k in ["英", "en", "english"]): en_col = col
        
        if zh_col is None or en_col is None:
            zh_col, en_col = df.columns[0], df.columns[1]
            
        for _, row in df.iterrows():
            self.add_term(str(row[zh_col]), str(row[en_col]), overwrite)
        self.save()

    def import_docx(self, file_path: str, overwrite: bool = True):
        doc = Document(file_path)
        for table in doc.tables:
            for i, row in enumerate(table.rows):
                if i == 0: continue # 忽略表头
                if len(row.cells) >= 2:
                    self.add_term(row.cells[0].text, row.cells[1].text, overwrite)
        self.save()

    def export_json(self, path: str):
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=4)

    def get_terms(self, direction: str) -> Dict[str, str]:
        return self.data["zh2en"] if direction == "zh-en" else self.data["en2zh"]
