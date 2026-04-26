#!/usr/bin/env python3
"""
混凝土试件记录筛选工具 - 可视化界面版
支持单文件和文件夹批量处理
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime


# ========== 默认筛选参数 ==========
DEFAULT_AGE_MIN = 40
DEFAULT_AGE_MAX = 50
DEFAULT_DATE_START = ""
DEFAULT_DATE_END = ""
DEFAULT_SITE_NAME = ""
DEFAULT_GRADE = ""
DEFAULT_POUR_PART = ""


def parse_date(s):
    if not s or str(s).strip() == "":
        return None
    s = str(s).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%m/%d/%Y", "%m.%d", "%m/%d"]:
        try:
            dt = datetime.strptime(s, fmt)
            # 如果只解析出月日，年份丢失，用当前年份或从上下文推断
            if dt.year == 1900:
                dt = dt.replace(year=datetime.now().year)
            return dt
        except:
            continue
    return None


def calc_age(make_str):
    make_date = parse_date(make_str)
    if not make_date:
        return None
    today = datetime.now()
    age = (today - make_date).days
    return age if age > 0 else None


def read_excel(path):
    import pandas as pd
    return pd.read_excel(path)


def read_csv(path):
    import pandas as pd
    return pd.read_csv(path)



def read_doc(path):
    """读取旧版 .doc 文件：先用 LibreOffice 转为 docx，再按 docx 处理"""
    import subprocess, tempfile, os, shutil
    
    tmp_dir = tempfile.mkdtemp()
    try:
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "docx",
             "--outdir", tmp_dir, os.path.abspath(path)],
            capture_output=True, timeout=60
        )
        base = os.path.splitext(os.path.basename(path))[0]
        docx_path = os.path.join(tmp_dir, base + ".docx")
        if os.path.exists(docx_path):
            return read_docx(docx_path)
    except Exception:
        pass
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return None



