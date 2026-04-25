#!/usr/bin/env python3
"""
混凝土试件记录筛选脚本
用法：把数据文件拖进来，修改下面的筛选条件，运行即可
支持 .xlsx .xls .csv .docx
"""

import os
import sys
import pandas as pd
from datetime import datetime, timedelta

# ========== 筛选条件（自己改这里） ==========
# 日期范围：制作日期从 X 到 Y
DATE_START = "2026-01-01"   # 留空用 ""
DATE_END = "2026-03-31"

# 龄期范围（天）
AGE_MIN = 50
AGE_MAX = 60

# 其他条件（留空用 "" 表示不限）
SITE_NAME = ""          # 工地名称关键词
GRADE = ""              # 混凝土等级，如 "C50"
POUR_PART = ""          # 浇筑部位关键词

# 输入输出文件
INPUT_FILE = "input.xlsx"   # 拖进来改这里
OUTPUT_FILE = "output.docx"
# ============================================


def read_excel(path):
    import pandas as pd
    df = pd.read_excel(path)
    return df

def read_csv(path):
    import pandas as pd
    df = pd.read_csv(path)
    return df

def read_docx(path):
    import pandas as pd
    from docx import Document
    doc = Document(path)
    rows = []
    for table in doc.tables:
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows[1:], columns=rows[0])
    return df

def parse_date(s):
    if pd.isna(s) or str(s).strip() == "":
        return None
    s = str(s).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%m/%d/%Y"]:
        try:
            return datetime.strptime(s, fmt)
        except:
            continue
    return None

def calc_age_days(make_date, check_date=None):
    """计算龄期天数"""
    d1 = parse_date(make_date)
    if d1 is None:
        return None
    if check_date:
        d2 = parse_date(check_date)
        if d2:
            return (d2 - d1).days
    # 按当前日期推算检测日期（大约龄期）
    today = datetime.now()
    age = (today - d1).days
    return age if age > 0 else None

def filter_data(df):
    # 标准化列名
    col_map = {}
    for col in df.columns:
        c = col.strip().replace(" ", "").replace("\n", "")
        col_map[col] = c
    df = df.rename(columns=col_map)
    
    # 找关键列
    date_col = None
    age_col = None
    site_col = None
    grade_col = None
    pour_col = None
    seq_col = None
    
    for col in df.columns:
        c = col.lower()
        if "日期" in col or "date" in c:
            if "制作" in col or "make" in c:
                date_col = col
            elif "检测" in col or "check" in c or "test" in c:
                if age_col is None:
                    age_col = col
        if "龄期" in col or "age" in c or "天" in col:
            age_col = col
        if "工地" in col or "site" in c or "项目" in col:
            site_col = col
        if "等级" in col or "grade" in c or "强度" in col:
            grade_col = col
        if "浇筑" in col or "pour" in c or "部位" in col:
            pour_col = col
        if "序号" in col or "seq" in c or "no" in c:
            seq_col = col
    
    results = []
    
    for idx, row in df.iterrows():
        # 跳过表头
        if idx == 0 and str(row.iloc[0]).strip() in ["序号", "序列", "NO", "No"]:
            continue
        
        make_date_str = str(row[date_col]).strip() if date_col and date_col in row.index else ""
        make_date = parse_date(make_date_str)
        
        if make_date is None:
            continue
        
        # 日期范围筛选
        start_dt = parse_date(DATE_START) if DATE_START else None
        end_dt = parse_date(DATE_END) if DATE_END else None
        
        if start_dt and make_date < start_dt:
            continue
        if end_dt and make_date > end_dt:
            continue
        
        # 龄期计算
        check_date_str = str(row[age_col]).strip() if age_col and age_col in row.index else ""
        age_days = calc_age_days(make_date_str, check_date_str)
        
        if age_days is None:
            # 尝试从龄期列直接读
            age_str = str(row[age_col]).strip() if age_col and age_col in row.index else ""
            if "d" in age_str.lower():
                try:
                    age_days = int(age_str.lower().replace("d", "").strip())
                except:
                    age_days = 0
        
        if age_days < AGE_MIN or age_days > AGE_MAX:
            continue
        
        # 工地名称筛选
        if SITE_NAME:
            site_val = str(row[site_col]).strip() if site_col and site_col in row.index else ""
            if SITE_NAME not in site_val:
                continue
        
        # 等级筛选
        if GRADE:
            grade_val = str(row[grade_col]).strip() if grade_col and grade_col in row.index else ""
            if GRADE not in grade_val:
                continue
        
        # 浇筑部位筛选
        if POUR_PART:
            pour_val = str(row[pour_col]).strip() if pour_col and pour_col in row.index else ""
            if POUR_PART not in pour_val:
                continue
        
        results.append(row)
    
    return pd.DataFrame(results)

def to_docx(df, output_path):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
    doc = Document()
    
    # 标题
    title = doc.add_heading("混凝土试件记录筛选表", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 条件说明
    p = doc.add_paragraph()
    p.add_run(f"筛选条件：").bold = True
    conditions = []
    if DATE_START or DATE_END:
        conditions.append(f"制作日期：{DATE_START or '开始'} ~ {DATE_END or '至今'}")
    conditions.append(f"龄期：{AGE_MIN} ~ {AGE_MAX} 天")
    if SITE_NAME:
        conditions.append(f"工地：{SITE_NAME}")
    if GRADE:
        conditions.append(f"等级：{GRADE}")
    if POUR_PART:
        conditions.append(f"浇筑部位：{POUR_PART}")
    p.add_run(" | ".join(conditions) if conditions else "全部")
    
    p = doc.add_paragraph()
    p.add_run(f"筛选结果：共 {len(df)} 条记录").bold = True
    p = doc.add_paragraph()
    p.add_run(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    
    doc.add_paragraph()
    
    # 表格
    if df.empty:
        doc.add_paragraph("无符合条件的数据")
        doc.save(output_path)
        return
    
    # 表头
    headers = ["序号", "制作日期", "龄期(天)", "检测日期", "等级", "工地名称", "浇筑部位", "备注"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].bold = True
    
    # 数据行
    for i, (_, row) in enumerate(df.iterrows(), 1):
        cells = table.add_row().cells
        cells[0].text = str(i)
        
        # 制作日期
        date_col = [c for c in df.columns if "日期" in c and "制作" in c]
        make_date = str(row[date_col[0]]).strip() if date_col else ""
        cells[1].text = make_date
        
        # 龄期
        age_col = [c for c in df.columns if "龄期" in c or "天" in c]
        if age_col:
            cells[2].text = str(row[age_col[0]]).strip()
        else:
            cells[2].text = ""
        
        # 检测日期
        check_col = [c for c in df.columns if "检测" in c or "试压" in c or "日期" in c]
        check_col = [c for c in check_col if "制作" not in c]
        if check_col:
            cells[3].text = str(row[check_col[0]]).strip()
        else:
            cells[3].text = ""
        
        # 等级
        grade_col = [c for c in df.columns if "等级" in c or "强度" in c]
        cells[4].text = str(row[grade_col[0]]).strip() if grade_col else ""
        
        # 工地名称
        site_col = [c for c in df.columns if "工地" in c or "项目" in c or "名称" in c]
        cells[5].text = str(row[site_col[0]]).strip() if site_col else ""
        
        # 浇筑部位
        pour_col = [c for c in df.columns if "浇筑" in c or "部位" in c]
        cells[6].text = str(row[pour_col[0]]).strip() if pour_col else ""
        
        # 备注（龄期天数）
        cells[7].text = f"{AGE_MIN}-{AGE_MAX}天"
    
    doc.save(output_path)
    print(f"✅ 已生成：{output_path}")

def main():
    if len(sys.argv) > 1:
        INPUT_FILE = sys.argv[1]
    if len(sys.argv) > 2:
        OUTPUT_FILE = sys.argv[2]
    
    print(f"\n{'='*50}")
    print("🏗️  混凝土试件记录筛选脚本")
    print(f"{'='*50}")
    print(f"输入文件：{INPUT_FILE}")
    print(f"输出文件：{OUTPUT_FILE}")
    print(f"{'='*50}\n")
    
    if not os.path.exists(INPUT_FILE):
        print(f"❌ 文件不存在：{INPUT_FILE}")
        print("\n用法：")
        print("  python concrete_filter.py                      # 使用默认文件名")
        print("  python concrete_filter.py mydata.xlsx          # 指定输入文件")
        print("  python concrete_filter.py mydata.xlsx out.docx # 指定输入输出")
        return
    
    # 读取数据
    ext = os.path.splitext(INPUT_FILE)[1].lower()
    print(f"📖 读取 {ext} 文件...")
    
    if ext in [".xlsx", ".xls"]:
        df = read_excel(INPUT_FILE)
    elif ext == ".csv":
        df = read_csv(INPUT_FILE)
    elif ext == ".docx":
        df = read_docx(INPUT_FILE)
    else:
        print(f"❌ 不支持的文件格式：{ext}")
        return
    
    print(f"   共 {len(df)} 行数据")
    print(f"   列名：{list(df.columns)}\n")
    
    # 筛选
    print("🔍 筛选中...")
    result = filter_data(df)
    print(f"   符合条件：{len(result)} 条\n")
    
    # 输出
    to_docx(result, OUTPUT_FILE)
    print(f"\n📄 Word文档已生成：{OUTPUT_FILE}")
    print("\n按回车键退出...")
    input()


if __name__ == "__main__":
    main()
