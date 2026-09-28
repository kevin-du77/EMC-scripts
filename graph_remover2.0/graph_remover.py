#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Graph Data Remover — HTM 文件 Graph 数据段删除工具 (GUI)
依赖: graph_core.py（核心处理逻辑）

功能：
1. UI 界面，可选起始/终止 graph 编号删除数据段
2. 支持对整个文件进行操作（删除所有 graph 行）
3. 文件筛选：全部 / index.htm / graph_0.htm
4. 可选同时删除对应的 graph_x.htm 物理文件
5. 自动备份原文件（.bak）

用法:
    python3 graph_remover.py
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import glob

# 确保可以导入 graph_core
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from graph_core import GraphProcessor


class GraphRemoverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Graph Data Remover — HTM 文件处理工具")
        self.root.geometry("800x700")

        # 变量定义
        self.folder_path = tk.StringVar()
        self.file_filter_var = tk.StringVar(value="all")
        self.start_var = tk.StringVar(value="1")
        self.end_var = tk.StringVar(value="256")
        self.whole_file_var = tk.BooleanVar(value=False)
        self.backup_var = tk.BooleanVar(value=True)
        self.del_graph_files_var = tk.BooleanVar(value=True)

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ========== ① 选择文件/文件夹 ==========
        file_frame = ttk.LabelFrame(main_frame, text="① 选择文件或文件夹", padding=5)
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Entry(file_frame, textvariable=self.folder_path, width=50).pack(
            side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="浏览文件夹...", command=self.browse_folder).pack(
            side=tk.LEFT, padx=2)
        ttk.Button(file_frame, text="选单个文件...", command=self.browse_file).pack(
            side=tk.LEFT, padx=2)

        # ========== ② 文件筛选 ==========
        filter_frame = ttk.LabelFrame(main_frame, text="② 文件筛选（单选）", padding=5)
        filter_frame.pack(fill=tk.X, pady=5)

        ttk.Radiobutton(filter_frame, text="全部 .htm 文件", variable=self.file_filter_var,
                        value="all", command=self.on_file_filter_change).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(filter_frame, text="index.htm", variable=self.file_filter_var,
                        value="index", command=self.on_file_filter_change).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(filter_frame, text="graph_0.htm", variable=self.file_filter_var,
                        value="graph_0", command=self.on_file_filter_change).pack(side=tk.LEFT, padx=10)

        # ========== ③ Graph 范围 ==========
        range_frame = ttk.LabelFrame(main_frame, text="③ Graph 范围", padding=5)
        range_frame.pack(fill=tk.X, pady=5)

        ttk.Label(range_frame, text="起始 Graph 编号:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(range_frame, textvariable=self.start_var, width=10).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(range_frame, text="终止 Graph 编号:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(range_frame, textvariable=self.end_var, width=10).grid(row=0, column=3, padx=5, pady=5)

        self.whole_file_cb = ttk.Checkbutton(
            range_frame, text="对整个文件进行操作（忽略编号范围，删除所有 graph 行）",
            variable=self.whole_file_var, command=self.on_whole_file_toggle)
        self.whole_file_cb.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)

        # ========== ④ 选项 ==========
        options_frame = ttk.LabelFrame(main_frame, text="④ 选项", padding=5)
        options_frame.pack(fill=tk.X, pady=5)

        ttk.Checkbutton(options_frame, text="备份原文件（.bak）",
                        variable=self.backup_var).pack(anchor=tk.W, padx=5, pady=2)

        ttk.Checkbutton(options_frame,
                        text="同时删除对应的 graph_x.htm 文件（如 graph_253.htm ~ graph_256.htm）",
                        variable=self.del_graph_files_var).pack(anchor=tk.W, padx=5, pady=2)

        # ========== ⑤ 按钮 ==========
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="▶ 预览删除内容", command=self.preview_delete).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="▶▶ 执行删除并保存", command=self.execute_delete).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="🔄 刷新", command=self.refresh).pack(side=tk.LEFT, padx=10)

        # ========== ⑥ 运行日志 ==========
        log_frame = ttk.LabelFrame(main_frame, text="⑤ 运行日志", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ========== 事件处理 ==========
    def on_whole_file_toggle(self):
        """勾选"整个文件"时自动设置范围"""
        if self.whole_file_var.get():
            self.start_var.set("1")
            self.end_var.set("999999")

    def on_file_filter_change(self):
        pass

    def refresh(self):
        self.log_text.delete(1.0, tk.END)
        self.log("已刷新日志区域")

    # ========== 文件浏览 ==========
    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log(f"已选择文件夹: {path}")

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("HTM files", "*.htm *.html"), ("All files", "*.*")])
        if path:
            self.folder_path.set(path)
            self.log(f"已选择文件: {path}")

    # ========== 日志 ==========
    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # ========== 获取待处理文件列表 ==========
    def get_target_files(self):
        path = self.folder_path.get().strip()
        if not path:
            messagebox.showwarning("警告", "请先选择文件或文件夹")
            return []

        processor = GraphProcessor()
        files = processor.get_target_files(path, self.file_filter_var.get())

        if not files:
            messagebox.showwarning("警告", "未找到符合条件的文件")
            return []

        return files

    # ========== 预览 ==========
    def preview_delete(self):
        self.log_text.delete(1.0, tk.END)
        files = self.get_target_files()
        if not files:
            return

        try:
            start_num = int(self.start_var.get()) if not self.whole_file_var.get() else 0
            end_num = int(self.end_var.get()) if not self.whole_file_var.get() else 999999
        except ValueError:
            messagebox.showerror("错误", "起始/终止编号必须是数字")
            return

        if not self.whole_file_var.get() and start_num > end_num:
            messagebox.showerror("错误", "起始编号不能大于终止编号")
            return

        processor = GraphProcessor()

        for file_path in files:
            self.log(f"\n{'='*60}")
            self.log(f"预览文件: {os.path.basename(file_path)}")
            self.log(f"{'='*60}")

            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
            except Exception as e:
                self.log(f"❌ 读取失败: {e}")
                continue

            _, deleted_count, to_delete = processor.process_content(
                content, start_num, end_num, self.whole_file_var.get())

            if deleted_count > 0:
                self.log(f"  📋 将删除 {deleted_count} 行数据:")
                for g_num, gs, ge, text in to_delete:
                    freq_match = re.search(r'<TD>\s*([\d.]+)\s*</TD>', text)
                    freq = freq_match.group(1) if freq_match else "?"
                    self.log(f"    graph_{g_num} (频率: {freq} MHz)")
            else:
                self.log("  ✅ 没有找到符合条件的 graph 数据行")

            # 预览物理文件
            if self.del_graph_files_var.get():
                dir_path = os.path.dirname(os.path.abspath(file_path))
                self.log(f"\n  🗑 将删除的物理文件:")
                found_any = False
                actual_start = start_num if not self.whole_file_var.get() else 1
                actual_end = end_num if not self.whole_file_var.get() else 999999
                for g_num in range(actual_start, actual_end + 1):
                    g_file = os.path.join(dir_path, f"graph_{g_num}.htm")
                    if os.path.exists(g_file):
                        self.log(f"    graph_{g_num}.htm")
                        found_any = True
                if not found_any:
                    self.log(f"    (目录中未找到对应的 graph 文件)")

    # ========== 执行删除 ==========
    def execute_delete(self):
        files = self.get_target_files()
        if not files:
            return

        try:
            start_num = int(self.start_var.get()) if not self.whole_file_var.get() else 0
            end_num = int(self.end_var.get()) if not self.whole_file_var.get() else 999999
        except ValueError:
            messagebox.showerror("错误", "起始/终止编号必须是数字")
            return

        if not self.whole_file_var.get() and start_num > end_num:
            messagebox.showerror("错误", "起始编号不能大于终止编号")
            return

        scope_desc = "整个文件的所有 graph 行" if self.whole_file_var.get() \
            else f"graph_{start_num} ~ graph_{end_num}"

        msg = f"确定要处理 {len(files)} 个文件吗？\n\n"
        msg += f"删除范围: {scope_desc}\n"
        if self.del_graph_files_var.get():
            msg += f"同时删除对应的 graph_x.htm 物理文件: 是\n"
        else:
            msg += f"同时删除对应的 graph_x.htm 物理文件: 否\n"
        msg += "\n此操作不可逆（但有备份）！"

        if not messagebox.askyesno("确认", msg):
            return

        self.log_text.delete(1.0, tk.END)
        self.log("🚀 开始执行删除操作...\n")

        processor = GraphProcessor(
            backup=self.backup_var.get(),
            delete_graph_files=self.del_graph_files_var.get())

        total_rows = 0
        total_files = 0

        for file_path in files:
            self.log(f"\n{'='*60}")
            self.log(f"处理文件: {os.path.basename(file_path)}")
            self.log(f"{'='*60}")

            result = processor.process_file(
                file_path, start_num, end_num, self.whole_file_var.get())

            if result['error']:
                self.log(f"  ❌ {result['error']}")
                continue

            if result['deleted_rows'] > 0:
                self.log(f"  ✅ 删除了 {result['deleted_rows']} 行数据")
                total_rows += result['deleted_rows']
            else:
                self.log("  ✅ 没有找到符合条件的 graph 数据行")

            if result['deleted_files'] > 0:
                self.log(f"  🗑 删除了 {result['deleted_files']} 个 graph 文件")
                total_files += result['deleted_files']

        # 汇总
        self.log(f"\n{'='*60}")
        self.log(f"🎉 处理完成！")
        self.log(f"  删除数据行: {total_rows} 行")
        self.log(f"  删除物理文件: {total_files} 个")
        self.log(f"{'='*60}")

        messagebox.showinfo(
            "完成",
            f"处理完成！\n\n删除数据行: {total_rows}\n删除物理文件: {total_files}")


def main():
    root = tk.Tk()
    app = GraphRemoverApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
