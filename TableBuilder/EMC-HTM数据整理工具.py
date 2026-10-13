# -*- coding: utf-8 -*-
"""
EMC HTM 数据整理工具（GUI 版）
功能：
  - 选择 HTM 文件夹路径
  - 自动解析 graph_0（最差Margin汇总）+ graph_3~n（频点明细）
  - 自定义选择要输出的列
  - 结果预览（可自由调整行的上下顺序）
  - 保存/加载列模板
  - 选择输出路径和文件名
  - 一键导出 xlsx

依赖：pip install pandas beautifulsoup4 lxml openpyxl
Python 3.7+
"""

import os
import re
import sys
import json
import glob

import pandas as pd
from bs4 import BeautifulSoup, NavigableString

# ----------------------------------------------------------------------
# 跨平台导入 tkinter
# ----------------------------------------------------------------------
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
except ImportError:
    print("当前 Python 缺少 tkinter，请安装带 tkinter 的 Python。")
    sys.exit(1)


# ======================================================================
# 第一部分：HTM 数据解析（核心逻辑，与之前保持一致）
# ======================================================================

def _text(tag):
    return re.sub(r"\s+", " ", tag.get_text(" ", strip=True)).strip() if tag else ""


def _to_float(s):
    m = re.search(r"-?\d+\.?\d*", str(s))
    return float(m.group()) if m else None


def parse_index(index_path):
    """解析 index 页，返回 {graph_3: {detector, sn, test_case, unit_config, config_desc}}"""
    soup = BeautifulSoup(open(index_path, encoding="utf-8").read(), "lxml")
    meta = {}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        m = re.match(r"graph_(\d+)\.htm", href, re.I)
        if not m:
            continue
        gname = "graph_" + m.group(1)
        line = _text(a)
        rec = {}
        sn_m = re.search(r"(C-#\d+\([^)]*\))", line)
        tc_m = re.search(r"((?:H\d+C|Test)[^,]*,\s*[^,]*,\s*[^,]*,\s*\d+TB\s+\w+\b)", line)
        rec["sn"] = sn_m.group(1) if sn_m else None
        rec["test_case"] = tc_m.group(1).strip() if tc_m else None
        cells = line.split()
        for i, c in enumerate(cells):
            if c.lower() == "scan" and i > 0:
                rec["detector"] = cells[i - 1].capitalize() if cells[i-1].lower() != "scan" else None
                rec["unit_config"] = cells[i + 1] if i + 1 < len(cells) else None
                tail = " ".join(cells[i + 1:])
                if rec["sn"]:
                    tail = tail.replace(rec["sn"], " ", 1)
                if rec["test_case"]:
                    tail = tail.replace(rec["test_case"], " ", 1)
                rec["config_desc"] = re.sub(r"\s{2,}", " ", tail).strip()
                break
        if rec:
            meta[gname] = rec
    return meta


def _pick_dataframe(tables):
    for df in tables:
        if df.shape[1] < 8:
            continue
        col0 = df.iloc[:, 0].astype(str).str.strip()
        nums = pd.to_numeric(col0, errors="coerce")
        if nums.notna().sum() >= 3:
            return df
    return None


def parse_graph_detail(html_path):
    """解析 graph_3~n 的 ScanData 明细"""
    try:
        tables = pd.read_html(html_path)
    except Exception:
        return []
    df = _pick_dataframe(tables)
    if df is None or df.shape[0] < 3:
        return []
    records = []
    for _, row in df.iloc[2:].iterrows():
        vals = [str(x) for x in row.tolist()]
        if len(vals) < 12 or _to_float(vals[0]) is None:
            continue
        records.append({
            "graph": "",
            "detector": None,
            "unit_config": None,
            "sn": None,
            "test_case": None,
            "Frequency(MHz)": _to_float(vals[0]),
            "Level(dBuV)": _to_float(vals[1]),
            "Limit(dBuV)": _to_float(vals[2]),
            "Delta(dB)": _to_float(vals[3]),
            "Raw(dBuV)": _to_float(vals[7]),
            "LISN(dB)": _to_float(vals[8]),
            "Cable(dB)": _to_float(vals[9]),
            "Xducer(dB)": _to_float(vals[10]),
            "AllFactors(dB)": _to_float(vals[11]),
        })
    return records


def parse_summary(html_path):
    """解析 graph_0 的 Results Summary"""
    try:
        tables = pd.read_html(html_path)
    except Exception:
        return []
    df = _pick_dataframe(tables)
    records = []
    if df is None or df.shape[0] < 3:
        return records
    for _, row in df.iloc[2:].iterrows():
        vals = [str(x) for x in row.tolist()]
        if len(vals) < 12 or _to_float(vals[0]) is None:
            continue
        m = re.search(r"graph_(\d+)", vals[-1])
        ref = "graph_" + m.group(1) if m else ""
        detector = None
        md = re.search(r"(Peak|QPeak|QP|Ave|Average)", vals[1], re.I)
        if md:
            detector = md.group(1).capitalize()
        records.append({
            "graph": ref,
            "detector": detector,
            "Frequency(MHz)": _to_float(vals[0]),
            "Level(dBuV)": _to_float(vals[1]),
            "Limit(dBuV)": _to_float(vals[2]),
            "Margin(dB)": _to_float(vals[3]),
        })
    return records


def find_file(folder, prefix):
    for f in os.listdir(folder):
        if re.match(re.escape(prefix) + r"\.(htm|html|txt)$", f, re.I):
            return os.path.join(folder, f)
    return None


# 所有可选列定义
ALL_COLUMNS = [
    "graph", "detector", "unit_config", "sn", "test_case",
    "Frequency(MHz)", "Level(dBuV)", "Limit(dBuV)", "Delta(dB)",
    "Raw(dBuV)", "LISN(dB)", "Cable(dB)", "Xducer(dB)", "AllFactors(dB)",
]

# 所有可选列（含 margin 汇总列）
ALL_COLUMNS_SUMMARY = [
    "graph", "detector", "Frequency(MHz)", "Level(dBuV)", "Limit(dBuV)", "Margin(dB)",
]


# ======================================================================
# 第二部分：GUI 界面
# ======================================================================

class EMCToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EMC HTM 数据整理工具 v1.0")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)

        # 数据
        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), "EMC汇总结果.xlsx"))
        self.summary_data = []        # graph_0 最差 margin
        self.detail_data = []         # graph_3~n 明细
        self.preview_df = None        # 当前预览的 DataFrame
        self.selected_columns = list(ALL_COLUMNS)  # 默认全选

        self._build_ui()

    def _build_ui(self):
        # ============ 顶部：路径选择区 ============
        top_frame = ttk.LabelFrame(self.root, text="① 选择文件夹 & 输出位置", padding=10)
        top_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(top_frame, text="HTM 文件夹：").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Entry(top_frame, textvariable=self.folder_path, width=60).grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Button(top_frame, text="浏览...", command=self._browse_folder).grid(row=0, column=2, padx=5)

        ttk.Label(top_frame, text="输出文件：").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(top_frame, textvariable=self.output_path, width=60).grid(row=1, column=1, padx=5, sticky="ew")
        ttk.Button(top_frame, text="浏览...", command=self._browse_output).grid(row=1, column=2, padx=5)

        top_frame.columnconfigure(1, weight=1)

        # ============ 中部：左右分栏 ============
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # --- 左侧：列选择 ---
        left_frame = ttk.LabelFrame(main_frame, text="② 选择输出列", padding=10)
        left_frame.pack(side="left", fill="y", padx=(0, 5))

        self.col_listbox = tk.Listbox(left_frame, selectmode="multiple", height=16, exportselection=False)
        for col in ALL_COLUMNS:
            self.col_listbox.insert("end", col)
        # 默认全选
        for i in range(len(ALL_COLUMNS)):
            self.col_listbox.selection_set(i)
        self.col_listbox.pack(fill="y", expand=True)

        col_btn_frame = ttk.Frame(left_frame)
        col_btn_frame.pack(fill="x", pady=5)
        ttk.Button(col_btn_frame, text="全选", command=self._col_select_all).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(col_btn_frame, text="全不选", command=self._col_select_none).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(col_btn_frame, text="反选", command=self._col_invert).pack(side="left", expand=True, fill="x", padx=2)

        # --- 右侧：Sheet 选择 + 操作按钮 + 预览 ---
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True)

        # Sheet 类型选择
        sheet_frame = ttk.LabelFrame(right_frame, text="③ 选择数据类型", padding=10)
        sheet_frame.pack(fill="x")

        self.sheet_type = tk.StringVar(value="明细")
        ttk.Radiobutton(sheet_frame, text="频点明细（graph_3~n）", variable=self.sheet_type,
                        value="明细", command=self._on_sheet_change).pack(side="left", padx=10)
        ttk.Radiobutton(sheet_frame, text="最差Margin汇总（graph_0）", variable=self.sheet_type,
                        value="汇总", command=self._on_sheet_change).pack(side="left", padx=10)

        # 操作按钮区
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill="x", pady=8)

        ttk.Button(btn_frame, text="▶ 开始解析", command=self._do_parse).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="🔄 刷新预览", command=self._refresh_preview).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="⬆ 上移行", command=self._move_row_up).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="⬇ 下移行", command=self._move_row_down).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="🗑 删除行", command=self._delete_row).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="💾 保存模板", command=self._save_template).pack(side="left", padx=3)
        ttk.Button(btn_frame, text="📂 加载模板", command=self._load_template).pack(side="left", padx=3)

        # 预览区
        preview_frame = ttk.LabelFrame(right_frame, text="④ 结果预览（可调整行顺序）", padding=5)
        preview_frame.pack(fill="both", expand=True)

        # Treeview + Scrollbar
        tree_container = ttk.Frame(preview_frame)
        tree_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_container, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

        # ============ 底部：导出按钮 ============
        export_frame = ttk.Frame(self.root)
        export_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(export_frame, text="📊 导出 Excel", command=self._do_export,
                   style="Accent.TButton" if hasattr(ttk, "Style") else None).pack(side="right", padx=5)
        ttk.Label(export_frame, text="提示：解析后可在预览区调整行顺序，导出时按当前顺序保存").pack(side="left")

    # ------------------------------------------------------------------
    # 文件夹/输出路径
    # ------------------------------------------------------------------
    def _browse_folder(self):
        d = filedialog.askdirectory(title="选择 HTM 文件夹")
        if d:
            self.folder_path.set(d)

    def _browse_output(self):
        f = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if f:
            self.output_path.set(f)

    # ------------------------------------------------------------------
    # 列选择操作
    # ------------------------------------------------------------------
    def _get_selected_cols(self):
        sel = self.col_listbox.curselection()
        return [self.col_listbox.get(i) for i in sel]

    def _col_select_all(self):
        for i in range(self.col_listbox.size()):
            self.col_listbox.selection_set(i)

    def _col_select_none(self):
        self.col_listbox.selection_clear(0, "end")

    def _col_invert(self):
        all_indices = list(range(self.col_listbox.size()))
        sel = set(self.col_listbox.curselection())
        self.col_listbox.selection_clear(0, "end")
        for i in all_indices:
            if i not in sel:
                self.col_listbox.selection_set(i)

    # ------------------------------------------------------------------
    # Sheet 类型切换
    # ------------------------------------------------------------------
    def _on_sheet_change(self):
        """切换数据类型时，更新列选择列表"""
        self.col_listbox.delete(0, "end")
        cols = ALL_COLUMNS_SUMMARY if self.sheet_type.get() == "汇总" else ALL_COLUMNS
        for col in cols:
            self.col_listbox.insert("end", col)
        for i in range(len(cols)):
            self.col_listbox.selection_set(i)

    # ------------------------------------------------------------------
    # 解析
    # ------------------------------------------------------------------
    def _do_parse(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("错误", "请先选择有效的 HTM 文件夹！")
            return

        try:
            self.status_var.set("正在解析...")
            self.root.update()

            files = os.listdir(folder)

            # 解析 index
            index_file = find_file(folder, "index")
            index_meta = parse_index(index_file) if index_file else {}
            self.status_var.set(f"index 解析完成，共 {len(index_meta)} 条元数据")
            self.root.update()

            # 解析 graph_0 最差 margin
            g0 = find_file(folder, "graph_0")
            self.summary_data = parse_summary(g0) if g0 else []
            for r in self.summary_data:
                meta = index_meta.get(r["graph"])
                if meta and meta.get("detector"):
                    r["detector"] = meta["detector"]

            # 解析 graph_3~n 明细
            self.detail_data = []
            detail_files = []
            for f in files:
                m = re.match(r"graph_(\d+)\.(htm|html|txt)$", f, re.I)
                if m and int(m.group(1)) >= 3:
                    detail_files.append((int(m.group(1)), f))
            detail_files.sort()

            for num, f in detail_files:
                gname = f"graph_{num}"
                recs = parse_graph_detail(os.path.join(folder, f))
                if not recs:
                    continue
                meta = index_meta.get(gname, {})
                for r in recs:
                    r.update({
                        "graph": gname,
                        "detector": meta.get("detector"),
                        "unit_config": meta.get("unit_config"),
                        "sn": meta.get("sn"),
                        "test_case": meta.get("test_case"),
                    })
                self.detail_data.extend(recs)

            msg = f"解析完成：明细 {len(self.detail_data)} 行，最差汇总 {len(self.summary_data)} 行"
            self.status_var.set(msg)
            messagebox.showinfo("解析完成", msg)
            self._refresh_preview()

        except Exception as e:
            self.status_var.set(f"解析失败：{e}")
            messagebox.showerror("解析失败", str(e))

    # ------------------------------------------------------------------
    # 预览
    # ------------------------------------------------------------------
    def _refresh_preview(self):
        selected_cols = self._get_selected_cols()
        if not selected_cols:
            messagebox.showwarning("提示", "请至少选择一列！")
            return

        # 选择数据源
        if self.sheet_type.get() == "汇总":
            data = self.summary_data
        else:
            data = self.detail_data

        if not data:
            messagebox.showwarning("提示", "暂无数据，请先点击「开始解析」")
            return

        # 过滤列
        df = pd.DataFrame(data)
        available_cols = [c for c in selected_cols if c in df.columns]
        df = df[available_cols]

        self.preview_df = df
        self._render_tree(df)

        self.status_var.set(f"预览已刷新：{len(df)} 行 × {len(df.columns)} 列")

    def _render_tree(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=60, anchor="center")

        for idx, row in df.iterrows():
            values = [str(v) if pd.notna(v) else "" for v in row.tolist()]
            self.tree.insert("", "end", iid=str(idx), values=values)

    # ------------------------------------------------------------------
    # 行操作
    # ------------------------------------------------------------------
    def _get_selected_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选中一行")
            return None
        return sel[0]

    def _move_row_up(self):
        if self.preview_df is None:
            return
        iid = self._get_selected_item()
        if iid is None:
            return
        idx = list(self.tree.get_children()).index(iid)
        if idx == 0:
            return
        # 在 DataFrame 中交换行
        df = self.preview_df
        i1, i2 = df.index[idx], df.index[idx - 1]
        df.iloc[[idx, idx - 1]] = df.iloc[[idx - 1, idx]].values
        self.preview_df = df
        self._render_tree(df)
        self.tree.selection_set(iid)

    def _move_row_down(self):
        if self.preview_df is None:
            return
        iid = self._get_selected_item()
        if iid is None:
            return
        children = list(self.tree.get_children())
        idx = children.index(iid)
        if idx >= len(children) - 1:
            return
        df = self.preview_df
        df.iloc[[idx, idx + 1]] = df.iloc[[idx + 1, idx]].values
        self.preview_df = df
        self._render_tree(df)
        self.tree.selection_set(iid)

    def _delete_row(self):
        if self.preview_df is None:
            return
        iid = self._get_selected_item()
        if iid is None:
            return
        idx = list(self.tree.get_children()).index(iid)
        self.preview_df = self.preview_df.drop(self.preview_df.index[idx]).reset_index(drop=True)
        self._render_tree(self.preview_df)
        self.status_var.set(f"已删除第 {idx+1} 行")

    # ------------------------------------------------------------------
    # 模板保存/加载
    # ------------------------------------------------------------------
    def _save_template(self):
        cols = self._get_selected_cols()
        if not cols:
            messagebox.showwarning("提示", "请至少选择一列！")
            return
        name = simpledialog.askstring("保存模板", "请输入模板名称：")
        if not name:
            return
        tmpl_dir = os.path.join(os.getcwd(), "templates")
        os.makedirs(tmpl_dir, exist_ok=True)
        path = os.path.join(tmpl_dir, f"{name}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({
                "columns": cols,
                "sheet_type": self.sheet_type.get(),
            }, f, ensure_ascii=False, indent=2)
        self.status_var.set(f"模板已保存：{path}")
        messagebox.showinfo("保存成功", f"模板已保存到：\n{path}")

    def _load_template(self):
        tmpl_dir = os.path.join(os.getcwd(), "templates")
        if not os.path.isdir(tmpl_dir):
            messagebox.showwarning("提示", "暂无保存的模板")
            return
        path = filedialog.askopenfilename(
            title="选择模板文件",
            initialdir=tmpl_dir,
            filetypes=[("JSON 模板", "*.json")]
        )
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            tmpl = json.load(f)
        cols = tmpl.get("columns", [])
        sheet_type = tmpl.get("sheet_type", "明细")

        # 切换 sheet 类型
        self.sheet_type.set(sheet_type)
        self._on_sheet_change()

        # 选中对应列
        self.col_listbox.selection_clear(0, "end")
        for col in cols:
            for i in range(self.col_listbox.size()):
                if self.col_listbox.get(i) == col:
                    self.col_listbox.selection_set(i)
                    break
        self.status_var.set(f"模板已加载：{os.path.basename(path)}")

    # ------------------------------------------------------------------
    # 导出
    # ------------------------------------------------------------------
    def _do_export(self):
        if self.preview_df is None or self.preview_df.empty:
            messagebox.showwarning("提示", "请先解析数据并刷新预览！")
            return

        out = self.output_path.get().strip()
        if not out:
            messagebox.showerror("错误", "请设置输出文件路径！")
            return

        try:
            self.status_var.set("正在导出...")
            self.root.update()

            df = self.preview_df

            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                sheet_name = "最差Margin汇总" if self.sheet_type.get() == "汇总" else "频点明细"
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # 如果有另一份数据也顺便导出到第二个 sheet
                if self.sheet_type.get() == "汇总" and self.detail_data:
                    ddf = pd.DataFrame(self.detail_data)
                    cols = [c for c in ALL_COLUMNS if c in ddf.columns]
                    ddf[cols].to_excel(writer, sheet_name="频点明细", index=False)
                elif self.sheet_type.get() == "明细" and self.summary_data:
                    sdf = pd.DataFrame(self.summary_data)
                    cols = [c for c in ALL_COLUMNS_SUMMARY if c in sdf.columns]
                    sdf[cols].to_excel(writer, sheet_name="最差Margin汇总", index=False)

            self.status_var.set(f"导出成功：{out}")
            messagebox.showinfo("导出成功", f"文件已保存到：\n{out}")

        except Exception as e:
            self.status_var.set(f"导出失败：{e}")
            messagebox.showerror("导出失败", str(e))


# ======================================================================
# 主入口
# ======================================================================

def main():
    root = tk.Tk()
    app = EMCToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
