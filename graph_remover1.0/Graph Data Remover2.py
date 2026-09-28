#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Graph Data Remover — HTM 文件 Graph 数据段删除工具
功能：
1. UI 界面，可选起始/终止 graph 编号删除数据段
2. 支持对整个文件进行操作
3. 文件筛选：全部 / index.htm / graph_0.htm
4. 可选同时删除对应的 graph_x.htm 物理文件
5. 自动备份原文件（.bak）
"""

import os
import re
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import glob


class GraphRemoverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Graph Data Remover — HTM 文件处理工具")
        self.root.geometry("750x650")

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

        ttk.Entry(file_frame, textvariable=self.folder_path, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="浏览文件夹...", command=self.browse_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(file_frame, text="选单个文件...", command=self.browse_file).pack(side=tk.LEFT, padx=2)

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

        ttk.Checkbutton(range_frame, text="对整个文件进行操作（忽略编号范围，删除所有 graph 行）",
                        variable=self.whole_file_var).grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)

        # ========== ④ 选项 ==========
        options_frame = ttk.LabelFrame(main_frame, text="④ 选项", padding=5)
        options_frame.pack(fill=tk.X, pady=5)

        ttk.Checkbutton(options_frame, text="备份原文件（.bak）",
                        variable=self.backup_var).pack(anchor=tk.W, padx=5, pady=2)

        # ✅ 新增：是否同时删除对应的 graph_x.htm 文件
        ttk.Checkbutton(options_frame,
                        text="同时删除对应的 graph_x.htm 文件（如 graph_253.htm）",
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

    # ========== 文件浏览 ==========
    def browse_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log(f"已选择文件夹: {path}")

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("HTM files", "*.htm *.html"), ("All files", "*.*")])
        if path:
            self.folder_path.set(path)
            self.log(f"已选择文件: {path}")

    def on_file_filter_change(self):
        pass  # 预留接口

    def refresh(self):
        self.log_text.delete(1.0, tk.END)
        self.log("已刷新日志区域")

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

        files = []
        if os.path.isfile(path):
            files.append(path)
        elif os.path.isdir(path):
            filter_type = self.file_filter_var.get()
            if filter_type == "all":
                files = glob.glob(os.path.join(path, "*.htm")) + glob.glob(os.path.join(path, "*.html"))
            elif filter_type == "index":
                f = os.path.join(path, "index.htm")
                if os.path.exists(f):
                    files.append(f)
            elif filter_type == "graph_0":
                f = os.path.join(path, "graph_0.htm")
                if os.path.exists(f):
                    files.append(f)

        if not files:
            messagebox.showwarning("警告", "未找到符合条件的文件")
            return []

        return sorted(set(files))

    # ========== 预览 ==========
    def preview_delete(self):
        self.log_text.delete(1.0, tk.END)
        files = self.get_target_files()
        if not files:
            return

        start_num = int(self.start_var.get()) if not self.whole_file_var.get() else 0
        end_num = int(self.end_var.get()) if not self.whole_file_var.get() else 999999

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

            # 定位 Results Summary 表格
            table_match = re.search(r'<TABLE[^>]*>.*?</TABLE>', content, re.DOTALL | re.IGNORECASE)
            if not table_match:
                self.log("  ⚠ 未找到表格")
                continue

            table_content = table_match.group(0)
            rows = re.findall(r'<TR>.*?</TR>', table_content, re.DOTALL | re.IGNORECASE)

            self.log(f"  表格总行数: {len(rows)}")

            # 筛选要删除的行
            to_delete = []
            for i, row in enumerate(rows):
                graph_match = re.search(r'graph_(\d+)\.htm', row, re.IGNORECASE)
                if graph_match:
                    graph_num = int(graph_match.group(1))
                    if self.whole_file_var.get() or (start_num <= graph_num <= end_num):
                        to_delete.append((i, graph_num, row))

            if to_delete:
                self.log(f"  📋 将删除 {len(to_delete)} 行:")
                for idx, (row_idx, g_num, row) in enumerate(to_delete):
                    # 提取频率信息
                    freq_match = re.search(r'<TD>\s*([\d.]+)\s*</TD>', row)
                    freq = freq_match.group(1) if freq_match else "?"
                    self.log(f"    graph_{g_num}.htm (行{row_idx}, 频率: {freq} MHz)")
            else:
                self.log("  ✅ 没有找到符合条件的 graph 数据行")

            # 预览要删除的物理文件
            if self.del_graph_files_var.get():
                dir_path = os.path.dirname(os.path.abspath(file_path))
                deleted_count = 0
                self.log(f"  🗑 将删除的物理文件:")
                for g_num in range(start_num, end_num + 1):
                    g_filepath = os.path.join(dir_path, f"graph_{g_num}.htm")
                    if os.path.exists(g_filepath):
                        self.log(f"    graph_{g_num}.htm")
                        deleted_count += 1
                if deleted_count == 0:
                    self.log(f"    (未找到对应的 graph_{start_num}.htm ~ graph_{end_num}.htm 文件)")

    # ========== 执行删除 ==========
    def execute_delete(self):
        files = self.get_target_files()
        if not files:
            return

        if not messagebox.askyesno("确认", f"确定要处理 {len(files)} 个文件吗？\n\n此操作不可逆（但有备份）！"):
            return

        self.log_text.delete(1.0, tk.END)
        self.log("🚀 开始执行删除操作...\n")

        start_num = int(self.start_var.get()) if not self.whole_file_var.get() else 0
        end_num = int(self.end_var.get()) if not self.whole_file_var.get() else 999999

        total_deleted_rows = 0
        total_deleted_files = 0

        for file_path in files:
            self.log(f"\n{'='*60}")
            self.log(f"处理文件: {os.path.basename(file_path)}")
            self.log(f"{'='*60}")

            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
            except Exception as e:
                self.log(f"❌ 读取失败: {e}")
                continue

            # 备份
            if self.backup_var.get():
                bak_path = file_path + ".bak"
                try:
                    shutil.copy2(file_path, bak_path)
                    self.log(f"  💾 已备份: {os.path.basename(bak_path)}")
                except Exception as e:
                    self.log(f"  ❌ 备份失败: {e}")

            # 定位 Results Summary 表格
            table_match = re.search(r'<TABLE[^>]*>.*?</TABLE>', content, re.DOTALL | re.IGNORECASE)
            if not table_match:
                self.log("  ⚠ 未找到表格，跳过")
                continue

            table_start = table_match.start()
            table_end = table_match.end()
            table_content = table_match.group(0)

            # 解析所有行
            rows = list(re.finditer(r'<TR>.*?</TR>', table_content, re.DOTALL | re.IGNORECASE))

            self.log(f"  表格总行数: {len(rows)}")

            # 确定要删除的行（从后往前删）
            to_delete_indices = []
            for i, row_match in enumerate(rows):
                row_text = row_match.group(0)
                graph_match = re.search(r'graph_(\d+)\.htm', row_text, re.IGNORECASE)
                if graph_match:
                    graph_num = int(graph_match.group(1))
                    if self.whole_file_var.get() or (start_num <= graph_num <= end_num):
                        to_delete_indices.append(i)

            if not to_delete_indices:
                self.log("  ✅ 没有找到符合条件的 graph 数据行，跳过")
                continue

            self.log(f"  📋 将删除 {len(to_delete_indices)} 行数据")

            # 从后往前删除行
            new_table_content = table_content
            deleted_count = 0
            for idx in reversed(to_delete_indices):
                row_match = rows[idx]
                # 获取在 table_content 中的绝对位置
                row_text = row_match.group(0)
                # 在 new_table_content 中找到并删除
                new_table_content = new_table_content.replace(row_text, '', 1)
                deleted_count += 1

            # 替换原表格
            new_content = content[:table_start] + new_table_content + content[table_end:]

            # 写入文件
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                self.log(f"  ✅ 已保存修改，删除了 {deleted_count} 行数据")
                total_deleted_rows += deleted_count
            except Exception as e:
                self.log(f"  ❌ 写入失败: {e}")
                continue

            # ✅ 新增：删除对应的 graph_x.htm 物理文件
            if self.del_graph_files_var.get():
                dir_path = os.path.dirname(os.path.abspath(file_path))
                file_deleted_count = 0

                for g_num in range(start_num, end_num + 1):
                    g_filename = f"graph_{g_num}.htm"
                    g_filepath = os.path.join(dir_path, g_filename)

                    if os.path.exists(g_filepath):
                        try:
                            os.remove(g_filepath)
                            file_deleted_count += 1
                            self.log(f"  🗑 已删除文件: {g_filename}")
                        except Exception as e:
                            self.log(f"  ❌ 删除 {g_filename} 失败: {e}")

                if file_deleted_count > 0:
                    self.log(f"  ✅ 共删除 {file_deleted_count} 个 graph 文件")
                    total_deleted_files += file_deleted_count
                else:
                    self.log(f"  ⚠ 目录中未找到对应的 graph_{start_num}.htm ~ graph_{end_num}.htm 文件")

        # 汇总
        self.log(f"\n{'='*60}")
        self.log(f"🎉 处理完成！")
        self.log(f"  删除数据行: {total_deleted_rows} 行")
        self.log(f"  删除物理文件: {total_deleted_files} 个")
        self.log(f"{'='*60}")

        messagebox.showinfo("完成", f"处理完成！\n删除数据行: {total_deleted_rows}\n删除物理文件: {total_deleted_files}")

    # ========== 核心解析函数（备用） ==========
    def parse_graph_rows(self, content):
        """解析内容中所有 graph 数据行，返回 [(graph_num, row_text, start_pos, end_pos), ...]"""
        results = []
        rows = list(re.finditer(r'<TR>.*?</TR>', content, re.DOTALL | re.IGNORECASE))
        for row_match in rows:
            row_text = row_match.group(0)
            graph_match = re.search(r'graph_(\d+)\.htm', row_text, re.IGNORECASE)
            if graph_match:
                graph_num = int(graph_match.group(1))
                results.append((graph_num, row_text, row_match.start(), row_match.end()))
        return results


def main():
    root = tk.Tk()
    app = GraphRemoverApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()