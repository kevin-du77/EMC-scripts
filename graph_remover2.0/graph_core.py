#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Graph Data Remover — 核心处理逻辑
独立于 GUI，方便测试和复用。
"""


import os
import re
import shutil


class GraphProcessor:
    """核心处理器：解析 HTM 文件中的 Results Summary 表格，删除指定 graph 数据行。"""

    def __init__(self, backup=True, delete_graph_files=True):
        self.backup = backup
        self.delete_graph_files = delete_graph_files

    # ========== 公共 API ==========

    def process_file(self, file_path, start_num=0, end_num=999999, whole_file=False):
        """
        处理单个文件：删除指定范围的 graph 数据行，可选删除物理文件。

        参数:
            file_path:       HTM 文件路径
            start_num:       起始 graph 编号（whole_file=True 时忽略）
            end_num:         终止 graph 编号（whole_file=True 时忽略）
            whole_file:      是否删除所有 graph 行

        返回:
            dict: {
                'deleted_rows': int,      # 删除的数据行数
                'deleted_files': int,     # 删除的物理文件数
                'error': str or None      # 错误信息
            }
        """
        result = {'deleted_rows': 0, 'deleted_files': 0, 'error': None}

        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
        except Exception as e:
            result['error'] = f"读取失败: {e}"
            return result

        # 备份
        if self.backup:
            try:
                shutil.copy2(file_path, file_path + ".bak")
            except Exception as e:
                result['error'] = f"备份失败: {e}"
                return result

        # 处理内容
        new_content, deleted_count, to_delete = self.process_content(
            content, start_num, end_num, whole_file)

        if deleted_count > 0:
            # 写入
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(new_content)
            except Exception as e:
                result['error'] = f"写入失败: {e}"
                return result

        result['deleted_rows'] = deleted_count

        # 删除物理文件
        if self.delete_graph_files and deleted_count > 0:
            dir_path = os.path.dirname(os.path.abspath(file_path))
            deleted_file_count = self._delete_physical_files(
                dir_path, start_num, end_num, whole_file, to_delete)
            result['deleted_files'] = deleted_file_count

        return result

    def process_content(self, content, start_num, end_num, whole_file):
        """
        在 content 中找到 Results Summary 表格，删除指定范围的 graph 数据行。

        返回: (new_content, deleted_count, to_delete_list)
            to_delete_list: [(graph_num, global_start, global_end, row_text), ...]
        """
        # 1. 找到包含 "Results Summary" 的 TABLE
        all_tables = list(re.finditer(
            r'<TABLE[^>]*>.*?</TABLE>', content, re.DOTALL | re.IGNORECASE))

        results_table = None
        for tm in all_tables:
            if 'Results Summary' in tm.group(0):
                results_table = tm
                break

        if not results_table:
            return content, 0, []

        table_abs_start = results_table.start()
        table_content = results_table.group(0)

        # 2. 提取表格内所有 <TR> 行（用全局偏移）
        rows = list(re.finditer(r'<TR>.*?</TR>', table_content, re.DOTALL | re.IGNORECASE))

        # 3. 找出包含 graph_x.htm 的数据行
        graph_rows = []
        for row_match in rows:
            row_text = row_match.group(0)
            gm = re.search(r'graph_(\d+)\.htm', row_text, re.IGNORECASE)
            if gm:
                g_num = int(gm.group(1))
                global_start = table_abs_start + row_match.start()
                global_end = table_abs_start + row_match.end()
                graph_rows.append((g_num, global_start, global_end, row_text))

        # 4. 筛选要删除的行
        if whole_file:
            to_delete = graph_rows
        else:
            to_delete = [(gn, gs, ge, t) for gn, gs, ge, t in graph_rows
                         if start_num <= gn <= end_num]

        if not to_delete:
            return content, 0, []

        # 5. 从后往前删除（避免偏移错位）
        new_content = content
        deleted_count = 0
        for g_num, gs, ge, text in sorted(to_delete, key=lambda x: x[1], reverse=True):
            snippet = new_content[gs:ge]
            if f'graph_{g_num}' in snippet or f'graph_{g_num}.' in snippet:
                new_content = new_content[:gs] + new_content[ge:]
                deleted_count += 1

        return new_content, deleted_count, to_delete

    def parse_graph_rows(self, content):
        """
        解析 content 中所有 graph 数据行。
        返回: [(graph_num, global_start, global_end, row_text), ...]
        """
        all_tables = list(re.finditer(
            r'<TABLE[^>]*>.*?</TABLE>', content, re.DOTALL | re.IGNORECASE))
        results_table = None
        for tm in all_tables:
            if 'Results Summary' in tm.group(0):
                results_table = tm
                break
        if not results_table:
            return []

        table_abs_start = results_table.start()
        table_content = results_table.group(0)
        rows = list(re.finditer(r'<TR>.*?</TR>', table_content, re.DOTALL | re.IGNORECASE))

        results = []
        for row_match in rows:
            row_text = row_match.group(0)
            gm = re.search(r'graph_(\d+)\.htm', row_text, re.IGNORECASE)
            if gm:
                g_num = int(gm.group(1))
                gs = table_abs_start + row_match.start()
                ge = table_abs_start + row_match.end()
                results.append((g_num, gs, ge, row_text))
        return results

    def get_target_files(self, path, filter_type="all"):
        """
        根据筛选条件获取待处理文件列表。

        参数:
            path:        文件或文件夹路径
            filter_type: "all" / "index" / "graph_0"

        返回:
            list: 排序后的文件路径列表
        """
        import glob

        if os.path.isfile(path):
            return [path]

        files = []
        if filter_type == "all":
            files = glob.glob(os.path.join(path, "*.htm")) + \
                    glob.glob(os.path.join(path, "*.html"))
        elif filter_type == "index":
            f = os.path.join(path, "index.htm")
            if os.path.exists(f):
                files.append(f)
        elif filter_type == "graph_0":
            f = os.path.join(path, "graph_0.htm")
            if os.path.exists(f):
                files.append(f)

        return sorted(set(files))

    # ========== 私有方法 ==========

    def _delete_physical_files(self, dir_path, start_num, end_num, whole_file, to_delete):
        """删除对应的 graph_x.htm 物理文件，返回删除数量。"""
        deleted = 0

        if whole_file:
            # 从已删除的行中提取实际 graph 编号
            for g_num in sorted(set(gn for gn, gs, ge, t in to_delete)):
                g_file = os.path.join(dir_path, f"graph_{g_num}.htm")
                if os.path.exists(g_file):
                    try:
                        os.remove(g_file)
                        deleted += 1
                    except Exception:
                        pass
        else:
            for g_num in range(start_num, end_num + 1):
                g_file = os.path.join(dir_path, f"graph_{g_num}.htm")
                if os.path.exists(g_file):
                    try:
                        os.remove(g_file)
                        deleted += 1
                    except Exception:
                        pass

        return deleted
