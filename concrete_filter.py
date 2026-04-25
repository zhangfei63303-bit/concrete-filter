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
DEFAULT_AGE_MIN = 50
DEFAULT_AGE_MAX = 60
DEFAULT_DATE_START = ""
DEFAULT_DATE_END = ""
DEFAULT_SITE_NAME = ""
DEFAULT_GRADE = ""
DEFAULT_POUR_PART = ""


def parse_date(s):
    if not s or str(s).strip() == "":
        return None
    s = str(s).strip()
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%m/%d/%Y"]:
        try:
            return datetime.strptime(s, fmt)
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
    """手动解析旧版.doc文件（无需额外依赖）"""
    try:
        import pandas as pd
        with open(path, "rb") as f:
            raw = f.read()
        
        # 尝试用多种编码解码
        text = None
        for enc in ["gbk", "gb2312", "utf-8", "latin1"]:
            try:
                text = raw.decode(enc, errors="ignore")
                if "\x00" in text:
                    text = raw.decode(enc).replace("\x00", "")
                break
            except:
                continue
        
        if text is None:
            return None
        
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        header_idx = -1
        for i, line in enumerate(lines):
            if "序号" in line and ("制作日期" in line or "检测日期" in line):
                header_idx = i
                break
        
        if header_idx < 0:
            return None
        
        data_rows = []
        for line in lines[header_idx:]:
            cols = [c.strip() for c in line.split("|")]
            if len(cols) < 3:
                cols = [c.strip() for c in line.split("\t")]
            if len(cols) >= 3:
                first = cols[0].strip()
                if first.isdigit() or (first and first[0] in "123456789"):
                    data_rows.append(cols)
        
        if not data_rows:
            return None
        
        max_cols = max(len(row) for row in data_rows)
        normalized = []
        for row in data_rows:
            if len(row) < max_cols:
                row = row + [""] * (max_cols - len(row))
            normalized.append(row[:max_cols])
        
        headers = normalized[0]
        df = pd.DataFrame(normalized[1:], columns=headers)
        
        col_map = {}
        for col in df.columns:
            if "制作日期" in col: col_map[col] = "制作日期"
            elif "检测日期" in col: col_map[col] = "检测日期"
            elif "龄期" in col: col_map[col] = "龄期"
            elif "等级" in col: col_map[col] = "等级"
            elif "工地" in col or "项目" in col: col_map[col] = "工地名称"
            elif "浇筑" in col or "部位" in col: col_map[col] = "浇筑部位"
        df.rename(columns=col_map, inplace=True)
        
        return df
    except Exception as e:
        return None


def read_docx(path):
    from docx import Document
    doc = Document(path)
    rows = []
    for table in doc.tables:
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
    if not rows:
        return None
    import pandas as pd
    df = pd.DataFrame(rows[1:], columns=rows[0])
    return df


def filter_data(df, date_start, date_end, age_min, age_max, site_name, grade, pour_part):
    col_map = {col: col.strip().replace(" ", "").replace("\n", "") for col in df.columns}
    df = df.rename(columns=col_map)

    date_col = next((c for c in df.columns if "制作" in c and "日期" in c), None)
    age_col = next((c for c in df.columns if "龄期" in c), None)
    site_col = next((c for c in df.columns if "工地" in c), None)
    grade_col = next((c for c in df.columns if "等级" in c), None)
    pour_col = next((c for c in df.columns if "浇筑" in c or "部位" in c), None)

    results = []
    for idx, row in df.iterrows():
        if idx == 0 and str(row.iloc[0]).strip() in ["序号", "序列", "NO"]:
            continue

        make_str = str(row[date_col]).strip() if date_col else ""
        make_date = parse_date(make_str)
        if not make_date:
            continue

        start_dt = parse_date(date_start) if date_start else None
        end_dt = parse_date(date_end) if date_end else None
        if start_dt and make_date < start_dt:
            continue
        if end_dt and make_date > end_dt:
            continue

        age_days = calc_age(make_str)
        if age_days is None or age_days < age_min or age_days > age_max:
            continue

        if site_name and site_name not in str(row[site_col]).strip() if site_col else "":
            continue
        if grade and grade not in str(row[grade_col]).strip() if grade_col else "":
            continue
        if pour_part and pour_part not in str(row[pour_col]).strip() if pour_col else "":
            continue

        results.append((row, age_days))

    return results


def to_docx(results, output_path, all_columns, date_col, age_col, site_col, grade_col, pour_col):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    title = doc.add_heading("混凝土试件记录筛选表", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.add_run(f"筛选结果：共 {len(results)} 条记录").bold = True
    p = doc.add_paragraph()
    p.add_run(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")

    doc.add_paragraph()

    if not results:
        doc.add_paragraph("无符合条件的数据")
        doc.save(output_path)
        return

    headers = ["序号", "制作日期", "龄期(天)", "等级", "工地名称", "浇筑部位"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"

    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
        table.rows[0].cells[i].paragraphs[0].runs[0].bold = True

    for i, (row, age_days) in enumerate(results, 1):
        cells = table.add_row().cells
        cells[0].text = str(i)
        cells[1].text = str(row[date_col]).strip() if date_col else ""
        cells[2].text = str(age_days)
        cells[3].text = str(row[grade_col]).strip() if grade_col else ""
        cells[4].text = str(row[site_col]).strip() if site_col else ""
        cells[5].text = str(row[pour_col]).strip() if pour_col else ""

    doc.save(output_path)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("混凝土试件记录筛选工具")
        self.root.geometry("720x520")
        self.root.resizable(True, True)

        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar(value="output.docx")
        self.date_start = tk.StringVar(value=DEFAULT_DATE_START)
        self.date_end = tk.StringVar(value=DEFAULT_DATE_END)
        self.age_min = tk.IntVar(value=DEFAULT_AGE_MIN)
        self.age_max = tk.IntVar(value=DEFAULT_AGE_MAX)
        self.site_name = tk.StringVar(value=DEFAULT_SITE_NAME)
        self.grade = tk.StringVar(value=DEFAULT_GRADE)
        self.pour_part = tk.StringVar(value=DEFAULT_POUR_PART)

        self.build_ui()

    def build_ui(self):
        # 文件选择区
        frame_file = ttk.LabelFrame(self.root, text="文件选择", padding=10)
        frame_file.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_file, text="数据文件/文件夹：").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_file, textvariable=self.input_file, width=55).grid(row=0, column=1, padx=5)
        ttk.Button(frame_file, text="浏览...", command=self.browse_input).grid(row=0, column=2)

        ttk.Label(frame_file, text="输出文件：").grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Entry(frame_file, textvariable=self.output_file, width=55).grid(row=1, column=1, padx=5, pady=(5,0))
        ttk.Button(frame_file, text="浏览...", command=self.browse_output).grid(row=1, column=2, pady=(5,0))

        ttk.Label(frame_file, text="(可选择文件或整个文件夹，文件夹会自动扫描所有数据文件)").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(3,0), padx=2
        )

        # 筛选条件区
        frame_cond = ttk.LabelFrame(self.root, text="筛选条件", padding=10)
        frame_cond.pack(fill="x", padx=10, pady=5)

        # 第一行：日期范围
        ttk.Label(frame_cond, text="制作日期：").grid(row=0, column=0, sticky="w")
        ttk.Label(frame_cond, text="从").grid(row=0, column=1)
        ttk.Entry(frame_cond, textvariable=self.date_start, width=12).grid(row=0, column=2, padx=5)
        ttk.Label(frame_cond, text="到").grid(row=0, column=3)
        ttk.Entry(frame_cond, textvariable=self.date_end, width=12).grid(row=0, column=4, padx=5)
        ttk.Label(frame_cond, text="(格式如 2026-01-01，留空不限)").grid(row=0, column=5, sticky="w", padx=5)

        # 第二行：龄期
        ttk.Label(frame_cond, text="龄期(天)：").grid(row=1, column=0, sticky="w", pady=(5,0))
        ttk.Label(frame_cond, text="从").grid(row=1, column=1, pady=(5,0))
        ttk.Entry(frame_cond, textvariable=self.age_min, width=12).grid(row=1, column=2, padx=5, pady=(5,0))
        ttk.Label(frame_cond, text="到").grid(row=1, column=3, pady=(5,0))
        ttk.Entry(frame_cond, textvariable=self.age_max, width=12).grid(row=1, column=4, padx=5, pady=(5,0))
        ttk.Label(frame_cond, text="天").grid(row=1, column=5, sticky="w", padx=5, pady=(5,0))

        # 第三行：工地名称
        ttk.Label(frame_cond, text="工地名称：").grid(row=2, column=0, sticky="w", pady=(5,0))
        ttk.Entry(frame_cond, textvariable=self.site_name, width=30).grid(row=2, column=1, columnspan=4, sticky="w", padx=5, pady=(5,0))
        ttk.Label(frame_cond, text="(留空不限)").grid(row=2, column=5, sticky="w", padx=5, pady=(5,0))

        # 第四行：等级
        ttk.Label(frame_cond, text="混凝土等级：").grid(row=3, column=0, sticky="w", pady=(5,0))
        ttk.Entry(frame_cond, textvariable=self.grade, width=30).grid(row=3, column=1, columnspan=4, sticky="w", padx=5, pady=(5,0))
        ttk.Label(frame_cond, text="(如 C50，留空不限)").grid(row=3, column=5, sticky="w", padx=5, pady=(5,0))

        # 第五行：浇筑部位
        ttk.Label(frame_cond, text="浇筑部位：").grid(row=4, column=0, sticky="w", pady=(5,0))
        ttk.Entry(frame_cond, textvariable=self.pour_part, width=30).grid(row=4, column=1, columnspan=4, sticky="w", padx=5, pady=(5,0))
        ttk.Label(frame_cond, text="(留空不限)").grid(row=4, column=5, sticky="w", padx=5, pady=(5,0))

        # 运行按钮
        frame_btn = ttk.Frame(self.root)
        frame_btn.pack(pady=15)

        self.btn_run = ttk.Button(frame_btn, text="开始筛选", command=self.run_filter, width=20)
        self.btn_run.pack()

        # 状态区
        frame_status = ttk.LabelFrame(self.root, text="状态", padding=10)
        frame_status.pack(fill="both", expand=True, padx=10, pady=5)

        self.status_text = tk.Text(frame_status, height=8, state="disabled", wrap="word")
        self.status_text.pack(fill="both", expand=True)

    def browse_input(self):
        fname = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("Word文件", "*.docx"),
                ("所有文件", "*.*"),
                ("Word旧版", "*.doc")
            ]
        )
        if not fname:
            dname = filedialog.askdirectory(title="选择包含数据文件的文件夹")
            if dname:
                self.input_file.set(dname)
                if not self.output_file.get() or self.output_file.get() == "output.docx":
                    self.output_file.set("混凝土试件筛选结果.docx")
            return

        self.input_file.set(fname)
        if not self.output_file.get() or self.output_file.get() == "output.docx":
            base = os.path.splitext(os.path.basename(fname))[0]
            self.output_file.set(f"{base}_筛选结果.docx")

    def browse_output(self):
        fname = filedialog.asksaveasfilename(
            title="保存结果文件",
            defaultextension=".docx",
            filetypes=[("Word文件", "*.docx")]
        )
        if fname:
            self.output_file.set(fname)

    def log(self, msg):
        self.status_text.config(state="normal")
        self.status_text.insert("end", msg + "\n")
        self.status_text.see("end")
        self.status_text.config(state="disabled")
        self.root.update()

    def process_file(self, fpath, date_start, date_end, age_min, age_max, site_name, grade, pour_part):
        """处理单个文件，返回 (results, all_columns)"""
        import pandas as pd

        ext = os.path.splitext(fpath)[1].lower()
        if ext in [".xlsx", ".xls"]:
            df = read_excel(fpath)
        elif ext == ".csv":
            df = read_csv(fpath)
        elif ext == ".docx":
            df = read_docx(fpath)
            if df is None:
                return None, None
        else:
            return None, None

        col_map = {col: col.strip().replace(" ", "").replace("\n", "") for col in df.columns}
        df = df.rename(columns=col_map)

        results = filter_data(df, date_start, date_end, age_min, age_max, site_name, grade, pour_part)
        return results, df.columns

    def run_filter(self):
        input_path = self.input_file.get().strip()
        output_path = self.output_file.get().strip()

        if not input_path:
            messagebox.showerror("错误", "请选择数据文件或文件夹！")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"路径不存在：{input_path}")
            return

        self.btn_run.config(state="disabled", text="处理中...")
        self.status_text.config(state="normal")
        self.status_text.delete("1.0", "end")
        self.status_text.config(state="disabled")

        try:
            date_start = self.date_start.get().strip()
            date_end = self.date_end.get().strip()
            age_min = self.age_min.get()
            age_max = self.age_max.get()
            site_name = self.site_name.get().strip()
            grade = self.grade.get().strip()
            pour_part = self.pour_part.get().strip()

            # 判断是文件还是文件夹
            if os.path.isdir(input_path):
                # 文件夹批量处理
                self.log(f"📁 检测到文件夹：{input_path}")
                supported_ext = ['.xlsx', '.xls', '.csv', '.docx', '.doc']
                files = []
                for root, dirs, files_in_dir in os.walk(input_path):
                    for f in files_in_dir:
                        if any(f.lower().endswith(ext) for ext in supported_ext):
                            files.append(os.path.join(root, f))

                if not files:
                    raise Exception(f"文件夹及子文件夹中未找到支持的数据文件(.xlsx/.xls/.csv/.docx/.doc)")

                self.log(f"   找到 {len(files)} 个数据文件\n")

                all_results = []
                all_columns = None

                for i, fpath in enumerate(files, 1):
                    self.log(f"[{i}/{len(files)}] 处理：{os.path.basename(fpath)}")
                    results, cols = self.process_file(fpath, date_start, date_end, age_min, age_max, site_name, grade, pour_part)

                    if results is None:
                        self.log(f"   ⚠️ 跳过")
                        continue

                    self.log(f"   共 {len(results)} 条记录")
                    all_results.extend(results)
                    if cols is not None and all_columns is None:
                        all_columns = cols

                self.log(f"\n📊 共筛选出 {len(all_results)} 条记录")

                if all_results:
                    date_col = next((c for c in all_columns if "制作" in c and "日期" in c), None)
                    age_col = next((c for c in all_columns if "龄期" in c), None)
                    site_col = next((c for c in all_columns if "工地" in c), None)
                    grade_col = next((c for c in all_columns if "等级" in c), None)
                    pour_col = next((c for c in all_columns if "浇筑" in c or "部位" in c), None)

                    to_docx(all_results, output_path, all_columns, date_col, age_col, site_col, grade_col, pour_col)
                    self.log(f"✅ 完成！")
                    self.log(f"📄 结果已保存：{output_path}")
                    messagebox.showinfo("完成", f"筛选完成！\n\n共处理 {len(files)} 个文件\n合计找到 {len(all_results)} 条记录\n\n已保存到：{output_path}")
                else:
                    to_docx([], output_path, all_columns or [], None, None, None, None, None)
                    self.log(f"⚠️ 无符合条件的数据")
                    self.log(f"📄 结果已保存：{output_path}")
                    messagebox.showinfo("完成", f"筛选完成，但无符合条件的数据\n\n已保存到：{output_path}")

            else:
                # 单文件处理
                self.log(f"📖 读取文件：{input_path}")
                results, cols = self.process_file(input_path, date_start, date_end, age_min, age_max, site_name, grade, pour_part)

                if results is None:
                    raise Exception("无法读取文件，请确认文件格式正确")

                self.log(f"   共 {len(results)} 条记录\n")

                if results:
                    date_col = next((c for c in cols if "制作" in c and "日期" in c), None)
                    age_col = next((c for c in cols if "龄期" in c), None)
                    site_col = next((c for c in cols if "工地" in c), None)
                    grade_col = next((c for c in cols if "等级" in c), None)
                    pour_col = next((c for c in cols if "浇筑" in c or "部位" in c), None)

                    to_docx(results, output_path, cols, date_col, age_col, site_col, grade_col, pour_col)
                    self.log(f"✅ 完成！")
                    self.log(f"📄 结果已保存：{output_path}")
                    messagebox.showinfo("完成", f"筛选完成！\n\n共找到 {len(results)} 条记录\n\n已保存到：{output_path}")
                else:
                    to_docx([], output_path, cols, None, None, None, None, None)
                    self.log(f"⚠️ 无符合条件的数据")
                    messagebox.showinfo("完成", "无符合条件的数据")

        except Exception as e:
            self.log(f"❌ 错误：{str(e)}")
            messagebox.showerror("错误", str(e))

        finally:
            self.btn_run.config(state="normal", text="开始筛选")


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
