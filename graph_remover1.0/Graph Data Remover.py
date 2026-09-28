#!/usr/bin/env python3
"""
Graph Data Remover - HTM 文件 graph 数据段删除工具
功能：
  1. UI 界面（tkinter）
  2. 可选起始和终止 graph 编号，删除区间内的所有 graph 数据段
  3. 可选择对整个文件进行操作
  4. 复选框指定特定文件名：index.htm、graph_0.htm、graph_x.htm、全部
"""

import re
import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path


# ============================================================
# 核心逻辑：解析与删除
# ============================================================

def find_data_table(lines):
    """
    找到 Results Summary 表格的范围。
    返回 (table_start_idx, table_end_idx)，即 <TABLE ... Results Summary ...> 到 </TABLE> 的行索引。
    如果没有找到，返回 None。
    """
    table_start = None
    for i, line in enumerate(lines):
        # 找到包含 Results Summary 的 <TABLE> 标签
        if '<TABLE' in line and ('Results Summary' in line or (
            i + 1 < len(lines) and 'Results Summary' in lines[i + 1]
        )):
            table_start = i
            break

    if table_start is None:
        # 尝试通过 SatyaTable 标记查找
        for i, line in enumerate(lines):
            if '<TABLE' in line and i > 300:
                # 检查后面几行是否有 Results Summary
                for j in range(i, min(i + 5, len(lines))):
                    if 'Results Summary' in lines[j]:
                        table_start = i
                        break
                if table_start is not None:
                    break

    if table_start is None:
        return None

    # 从 table_start 开始找 </TABLE>
    table_end = None
    for i in range(table_start, len(lines)):
        if '</TABLE>' in lines[i]:
            table_end = i
            break

    if table_end is None:
        return None

    return table_start, table_end


def find_graph_rows(lines, table_start, table_end):
    """
    在表格范围内，找到所有属于 graph 数据行的 <TR>...</TR> 行索引范围。
    每个数据行包含 graph_NN.htm 的链接。
    返回 [(row_start_idx, row_end_idx, graph_num), ...]
    """
    graph_rows = []
    i = table_start
    while i <= table_end:
        line = lines[i]
        if '<TR>' in line or '<TR ' in line:
            # 检查这个 <TR> 是否包含 graph 链接
            # 收集从 <TR> 到 </TR> 的所有行
            row_lines = []
            j = i
            while j <= table_end and '</TR>' not in lines[j]:
                row_lines.append(lines[j])
                j += 1
            if j <= table_end:
                row_lines.append(lines[j])  # 包含 </TR>
            
            row_text = '\n'.join(row_lines)
            # 匹配 graph_N.htm 或 graph_NN.htm
            match = re.search(r'graph_(\d+)\.htm', row_text)
            if match:
                graph_num = int(match.group(1))
                graph_rows.append((i, j, graph_num))
            i = j + 1
        else:
            i += 1
    return graph_rows


def delete_graphs_in_content(content, start_graph, end_graph, whole_file=False):
    """
    删除 content 中 graph 编号在 [start_graph, end_graph] 范围内的数据。

    参数:
        whole_file: 如果 True，对整个文件中的所有 graph 引用进行处理
                    （包括导航区域的 graph 链接）；
                    如果 False，仅处理 Results Summary 表格中的数据行。

    返回 (修改后的 content, 删除的 graph 编号列表)
    """
    lines = content.split('\n')

    if whole_file:
        # === 对整个文件操作 ===
        # 1) 先删除 Results Summary 表格中的数据行
        result = find_data_table(lines)
        deleted_graphs = []
        if result is not None:
            table_start, table_end = result
            graph_rows = find_graph_rows(lines, table_start, table_end)
            to_delete = [
                (rs, re_, gn) for rs, re_, gn in graph_rows
                if start_graph <= gn <= end_graph
            ]
            for rs, re_, gn in sorted(to_delete, key=lambda x: x[0], reverse=True):
                del lines[rs:re_ + 1]
                deleted_graphs.append(gn)

        # 2) 对整个文件，移除导航区域中孤立的 graph 链接（可选）
        #    为避免破坏结构，仅清理 <TD> 内指向被删 graph 的整行
        new_content = '\n'.join(lines)
        return new_content, sorted(set(deleted_graphs))

    else:
        # === 仅操作 Results Summary 表格 ===
        result = find_data_table(lines)
        if result is None:
            return content, []

        table_start, table_end = result
        graph_rows = find_graph_rows(lines, table_start, table_end)

        to_delete = [
            (rs, re_, gn) for rs, re_, gn in graph_rows
            if start_graph <= gn <= end_graph
        ]

        if not to_delete:
            return content, []

        deleted_graphs = []
        for rs, re_, gn in sorted(to_delete, key=lambda x: x[0], reverse=True):
            del lines[rs:re_ + 1]
            deleted_graphs.append(gn)

        new_content = '\n'.join(lines)
        return new_content, deleted_graphs


def collect_target_files(folder_path, file_filter_mode, specific_files):
    """
    根据复选框选择，收集需要处理的文件列表。
    file_filter_mode: 'all' / 'index' / 'graph_0' / 'graph_x'
    specific_files: 当 mode='graph_x' 时，为具体的文件名列表
    """
    folder = Path(folder_path)
    all_files = list(folder.glob('*.htm')) + list(folder.glob('*.html'))

    if file_filter_mode == 'all':
        return sorted(all_files)

    selected = []
    for f in sorted(all_files):
        fname = f.name
        if file_filter_mode == 'index' and fname == 'index.htm':
            selected.append(f)
        elif file_filter_mode == 'graph_0' and fname == 'graph_0.htm':
            selected.append(f)
        elif file_filter_mode == 'graph_x' and fname in specific_files:
            selected.append(f)
    return selected


# ============================================================
# UI 界面
# ============================================================

class GraphRemoverApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Graph Data Remover - HTM 文件编辑工具')
        self.root.geometry('820x720')

        # 状态变量
        self.folder_path = tk.StringVar()
        self.start_graph = tk.StringVar(value='1')
        self.end_graph = tk.StringVar(value='256')
        self.filter_mode = tk.StringVar(value='all')  # all / index / graph_0 / graph_x
        self.process_whole_file = tk.BooleanVar(value=True)
        self.backup_enabled = tk.BooleanVar(value=True)

        # 特定文件复选框状态（最多动态创建）
        self.specific_file_vars = {}

        self._build_ui()

    def _build_ui(self):
        # === 顶部：文件夹选择 ===
        top_frame = ttk.LabelFrame(self.root, text='① 选择文件 / 文件夹', padding=10)
        top_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(top_frame, text='文件夹路径：').grid(row=0, column=0, sticky='w')
        self.folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path, width=55)
        self.folder_entry.grid(row=0, column=1, padx=5)
        ttk.Button(top_frame, text='浏览...', command=self._browse_folder).grid(row=0, column=2)

        ttk.Label(top_frame, text='或直接选择单个文件：').grid(row=1, column=0, sticky='w', pady=(5, 0))
        self.single_file_path = tk.StringVar()
        self.single_file_entry = ttk.Entry(top_frame, textvariable=self.single_file_path, width=55)
        self.single_file_entry.grid(row=1, column=1, padx=5, pady=(5, 0))
        ttk.Button(top_frame, text='选择文件...', command=self._browse_file).grid(row=1, column=2, pady=(5, 0))

        # === 文件筛选 ===
        filter_frame = ttk.LabelFrame(self.root, text='② 文件筛选（复选框）', padding=10)
        filter_frame.pack(fill='x', padx=10, pady=5)

        self.cb_all = ttk.Radiobutton(
            filter_frame, text='全部 (*.htm / *.html)', variable=self.filter_mode, value='all',
            command=self._on_filter_change
        )
        self.cb_all.grid(row=0, column=0, sticky='w', padx=10)

        self.cb_index = ttk.Radiobutton(
            filter_frame, text='仅 index.htm', variable=self.filter_mode, value='index',
            command=self._on_filter_change
        )
        self.cb_index.grid(row=0, column=1, sticky='w', padx=10)

        self.cb_graph0 = ttk.Radiobutton(
            filter_frame, text='仅 graph_0.htm', variable=self.filter_mode, value='graph_0',
            command=self._on_filter_change
        )
        self.cb_graph0.grid(row=0, column=2, sticky='w', padx=10)

        self.cb_graphx = ttk.Radiobutton(
            filter_frame, text='指定 graph_x.htm：', variable=self.filter_mode, value='graph_x',
            command=self._on_filter_change
        )
        self.cb_graphx.grid(row=0, column=3, sticky='w', padx=10)

        # 指定文件复选框区域（动态填充）
        self.specific_frame = ttk.Frame(filter_frame)
        self.specific_frame.grid(row=1, column=0, columnspan=4, sticky='w', pady=(5, 0))
        ttk.Label(self.specific_frame, text='（请先选择文件夹，然后勾选要处理的文件）').pack(anchor='w')

        # === Graph 范围 ===
        range_frame = ttk.LabelFrame(self.root, text='③ Graph 范围', padding=10)
        range_frame.pack(fill='x', padx=10, pady=5)

        ttk.Label(range_frame, text='起始 Graph 编号：').grid(row=0, column=0, sticky='w')
        ttk.Entry(range_frame, textvariable=self.start_graph, width=10).grid(row=0, column=1, padx=5)

        ttk.Label(range_frame, text='终止 Graph 编号：').grid(row=0, column=2, sticky='w', padx=(20, 0))
        ttk.Entry(range_frame, textvariable=self.end_graph, width=10).grid(row=0, column=3, padx=5)

        ttk.Checkbutton(
            range_frame, text='对整个文件进行操作（删除指定范围内的所有 graph 行）',
            variable=self.process_whole_file
        ).grid(row=1, column=0, columnspan=4, sticky='w', pady=(10, 0))

        # === 选项 ===
        opt_frame = ttk.LabelFrame(self.root, text='④ 选项', padding=10)
        opt_frame.pack(fill='x', padx=10, pady=5)

        ttk.Checkbutton(
            opt_frame, text='备份原文件（.bak）', variable=self.backup_enabled
        ).grid(row=0, column=0, sticky='w')

        # === 操作按钮 ===
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(btn_frame, text='▶ 预览删除内容', command=self._preview).pack(side='left', padx=5)
        ttk.Button(btn_frame, text='▶▶ 执行删除并保存', command=self._execute, style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(btn_frame, text='🔄 刷新文件列表', command=self._refresh_specific_files).pack(side='left', padx=5)

        # === 日志输出 ===
        log_frame = ttk.LabelFrame(self.root, text='⑤ 运行日志', padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, wrap='word')
        self.log_text.pack(fill='both', expand=True)

        # 尝试设置主题色
        try:
            style = ttk.Style()
            style.configure('Accent.TButton', foreground='white', background='#2196F3')
        except Exception:
            pass

    # ---------- UI 回调 ----------

    def _log(self, msg):
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.root.update_idletasks()

    def _browse_folder(self):
        path = filedialog.askdirectory(title='选择包含 HTM 文件的文件夹')
        if path:
            self.folder_path.set(path)
            self._refresh_specific_files()

    def _browse_file(self):
        path = filedialog.askopenfilename(
            title='选择单个 HTM 文件',
            filetypes=[('HTM files', '*.htm *.html'), ('All files', '*.*')]
        )
        if path:
            self.single_file_path.set(path)
            self.folder_path.set(os.path.dirname(path))
            self._refresh_specific_files()

    def _on_filter_change(self):
        pass  # 可扩展：联动更新

    def _refresh_specific_files(self):
        """扫描文件夹，动态创建 graph_x.htm 的复选框"""
        for widget in self.specific_frame.winfo_children():
            widget.destroy()

        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            ttk.Label(self.specific_frame, text='（请先选择文件夹）').pack(anchor='w')
            return

        # 找到所有 graph_N.htm 文件
        graph_files = sorted(
            [f for f in os.listdir(folder)
             if re.match(r'graph_\d+\.htm', f) or re.match(r'graph_\d+\.html', f)]
        )

        if not graph_files:
            ttk.Label(self.specific_frame, text='（未找到 graph_N.htm 文件）').pack(anchor='w')
            return

        self.specific_file_vars.clear()

        # 使用内部 Frame + 滚动
        canvas = tk.Canvas(self.specific_frame, height=120)
        scrollbar = ttk.Scrollbar(self.specific_frame, orient='vertical', command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            '<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all'))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # 每行放 4 个复选框
        cols = 4
        for idx, fname in enumerate(graph_files):
            var = tk.BooleanVar(value=False)
            self.specific_file_vars[fname] = var
            cb = ttk.Checkbutton(scrollable_frame, text=fname, variable=var)
            cb.grid(row=idx // cols, column=idx % cols, sticky='w', padx=5, pady=1)

        self._log(f'已扫描文件夹，找到 {len(graph_files)} 个 graph 文件。')

    def _get_selected_specific_files(self):
        return [fname for fname, var in self.specific_file_vars.items() if var.get()]

    def _get_target_files(self):
        """返回要处理的文件路径列表"""
        single = self.single_file_path.get().strip()
        if single and os.path.isfile(single):
            return [Path(single)]

        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            return []

        mode = self.filter_mode.get()
        if mode == 'graph_x':
            specific = self._get_selected_specific_files()
            return collect_target_files(folder, 'graph_x', specific)
        else:
            return collect_target_files(folder, mode, [])

    def _preview(self):
        self.log_text.delete('1.0', 'end')
        self._do_work(dry_run=True)

    def _execute(self):
        self.log_text.delete('1.0', 'end')
        if not messagebox.askyesno('确认', '确定要删除指定 graph 数据并保存到文件吗？\n（建议先预览）'):
            return
        self._do_work(dry_run=False)

    def _do_work(self, dry_run=False):
        try:
            start_g = int(self.start_graph.get().strip())
            end_g = int(self.end_graph.get().strip())
        except ValueError:
            messagebox.showerror('错误', 'Graph 编号必须是整数')
            return

        if start_g > end_g:
            messagebox.showerror('错误', '起始编号不能大于终止编号')
            return

        target_files = self._get_target_files()
        if not target_files:
            self._log('⚠️ 没有找到匹配的文件。请检查文件夹路径和文件筛选选项。')
            return

        mode_text = {
            'all': '全部文件',
            'index': '仅 index.htm',
            'graph_0': '仅 graph_0.htm',
            'graph_x': '指定 graph_x.htm',
        }.get(self.filter_mode.get(), '?')

        self._log(f'=== {"预览" if dry_run else "执行"} ===')
        self._log(f'Graph 范围：{start_g} ~ {end_g}')
        self._log(f'文件筛选模式：{mode_text}')
        self._log(f'目标文件数：{len(target_files)}')
        self._log('-' * 50)

        total_deleted = 0
        for filepath in target_files:
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(filepath, 'r', encoding='latin-1') as f:
                    content = f.read()

            whole_file = self.process_whole_file.get()
            new_content, deleted = delete_graphs_in_content(content, start_g, end_g, whole_file)

            if not deleted:
                self._log(f'  {filepath.name}: 未找到匹配的 graph 行')
                continue

            self._log(f'  {filepath.name}: 将删除 graph {deleted}')

            if not dry_run:
                # 备份
                if self.backup_enabled.get():
                    bak_path = str(filepath) + '.bak'
                    shutil.copy2(str(filepath), bak_path)

                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                self._log(f'    ✅ 已保存（备份：{os.path.basename(bak_path)}）')

            total_deleted += len(deleted)

        self._log('-' * 50)
        if dry_run:
            self._log(f'📋 预览完成。共将删除 {total_deleted} 个 graph 数据行。')
        else:
            self._log(f'✅ 执行完成。共删除 {total_deleted} 个 graph 数据行。')


def main():
    root = tk.Tk()
    try:
        # 尝试使用主题
        style = ttk.Style()
        available = style.theme_names()
        if 'clam' in available:
            style.theme_use('clam')
    except Exception:
        pass
    app = GraphRemoverApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
