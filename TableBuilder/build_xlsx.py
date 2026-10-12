# -*- coding: utf-8 -*-
"""
把 graph_0（最差 margin 汇总）+ graph_3~graph_n（各频点扫描明细）整理成
以“工作表”模板为结构的 xlsx。

使用方法：
    python build_xlsx.py /path/to/htm_folder   # 指定文件夹
    python build_xlsx.py                        # 默认处理当前目录
输出：./EMC汇总结果.xlsx
"""
import os
import re
import sys
import glob

import pandas as pd
from bs4 import BeautifulSoup, NavigableString

# ----------------------------------------------------------------------
# 1. 工具函数
# ----------------------------------------------------------------------
def _text(tag):
    """安全取标签文本并规整空白"""
    return re.sub(r"\s+", " ", tag.get_text(" ", strip=True)).strip() if tag else ""


def _to_float(s):
    """把 '34.6 Av' / '66.2' 这类值转成浮点数，失败返回 None"""
    m = re.search(r"-?\d+\.?\d*", str(s))
    return float(m.group()) if m else None


def parse_scan_metadata(html):
    """从单个 graph 页面提取头部元数据（配置/序列号/检测模式等）"""
    soup = BeautifulSoup(html, "lxml")
    info = {}

    # 用 HTML 注释里的字段：<!-- key -->value<!-- /key -->
    for cmt in soup.find_all(string=lambda t: isinstance(t, NavigableString)):
        txt = str(cmt).strip()
        m = re.match(r"<!--\s*([A-Za-z]+)\s*-->\s*(.*)", txt, re.S)
        if m:
            key, val = m.group(1), re.sub(r"\s+", " ", m.group(2)).strip()
            if val:
                info[key] = val

    # 检测模式单独从页面文字里抠（如 "Receiver Ave Scan" / "Receiver Peak Scan"）
    body_text = soup.get_text(" ", strip=True)
    m = re.search(r"Receiver\s+(Peak|QPeak|QP|Ave|Average)\s+Scan", body_text, re.I)
    if m:
        info["Detector"] = m.group(1).capitalize()

    return info


DET_MAP = {"peak": "Peak", "qpeak": "QPeak", "qp": "QP", "ave": "Ave", "average": "Ave"}


def _norm_detector(s):
    """把 'Peak Scan'/'QPeak Scan'/'Ave Scan' 归一化"""
    if not s:
        return None
    key = s.lower().replace(" scan", "").strip()
    return DET_MAP.get(key, s.strip())


def parse_index(index_path):
    """
    解析 index 页，返回 { 'graph_3': {unit, sn, detector, config, ...}, ... }
    每行是一条 <A>，内容形如：
      graph_3  Line 1  Ave Scan  P1C-M_D3A1-R1  C-#5365(...)  H17C,...  6x USBC: ...
    按空白切分后：索引0=图名, 2=Detector, 3=UnitConfig, 4=SN, 5=TestCase, 6+=配置描述
    """
    soup = BeautifulSoup(open(index_path, encoding="utf-8").read(), "lxml")
    meta = {}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        m = re.match(r"graph_(\d+)\.htm", href, re.I)
        if not m:
            continue
        gname = "graph_" + m.group(1)
        cells = _text(a).split()
        rec = {}
        try:
            # 固定位置字段
            if len(cells) > 2:
                rec["detector"] = _norm_detector(cells[2] + " " + cells[3]
                                                 if cells[3].lower().startswith("scan") else cells[2])
            # 稳健策略：用正则特征定位
            # 整行原文（保留空格），用正则稳健提取 SN / TestCase
            line = _text(a)
            sn_m = re.search(r"(C-#\d+\([^)]*\))", line)
            tc_m = re.search(r"((?:H\d+C|Test)[^,]*,\s*[^,]*,\s*[^,]*,\s*\d+TB\s+\w+\b)", line)
            rec["sn"] = sn_m.group(1) if sn_m else None
            rec["test_case"] = tc_m.group(1).strip() if tc_m else None

            # detector 形如 'xxx Scan'
            for i, c in enumerate(cells):
                if c.lower() == "scan" and i > 0:
                    rec["detector"] = _norm_detector(cells[i - 1])
                    rec["unit_config"] = cells[i + 1] if i + 1 < len(cells) else None
                    # config_desc = detector 之后到行尾，扣掉已提取的 sn/testcase
                    tail = " ".join(cells[i + 1:])
                    if rec["sn"]:
                        tail = tail.replace(rec["sn"], " ", 1)
                    if rec["test_case"]:
                        tail = tail.replace(rec["test_case"], " ", 1)
                    rec["config_desc"] = re.sub(r"\s{2,}", " ", tail).strip()
                    break
        except Exception:
            pass
        meta[gname] = rec
    return meta


def _pick_dataframe(tables):
    """从 read_html 结果里挑出 ScanData 数据表：首列可转 float、且含频率数据"""
    for df in tables:
        if df.shape[1] < 8:
            continue
        col0 = df.iloc[:, 0].astype(str).str.strip()
        nums = pd.to_numeric(col0, errors="coerce")
        if nums.notna().sum() >= 3:          # 至少 3 行是数字才算数据表
            return df
    return None


def parse_graph_detail(html_path):
    """解析 graph_3~n 的 ScanData 明细表"""
    tables = pd.read_html(html_path)
    df = _pick_dataframe(tables)
    if df is None or df.shape[0] < 3:
        return []

    data = df.iloc[2:].copy()                # 跳过两行表头
    data = data[[c for c in data.columns if "Unnamed" not in str(c) or True]]
    records = []
    for _, row in data.iterrows():
        vals = [str(x) for x in row.tolist()]
        if len(vals) < 12 or _to_float(vals[0]) is None:
            continue
        records.append({
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
    """解析 graph_0（Results Summary）：每个 graph 的最差 margin 一行"""
    tables = pd.read_html(html_path)
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


# ----------------------------------------------------------------------
# 2. 主流程
# ----------------------------------------------------------------------
def build(folder):
    files = os.listdir(folder)
    htm = lambda n: os.path.join(folder, n)

    def find_file(prefix):
        """按前缀找文件，后缀支持 .htm / .html / .txt"""
        for f in files:
            if re.match(re.escape(prefix) + r"\.(htm|html|txt)$", f, re.I):
                return f
        return None

    # 2.1 解析 index 拿到所有 graph 的元数据
    index_file = find_file("index")
    index_meta = parse_index(htm(index_file)) if index_file else {}

    # 2.2 解析 graph_0 最差 margin 汇总
    g0 = find_file("graph_0")
    summary = parse_summary(htm(g0)) if g0 else []
    # 用 index 里的 detector 信息补全/覆盖
    for r in summary:
        meta = index_meta.get(r["graph"])
        if meta and meta.get("detector"):
            r["detector"] = meta["detector"]

    # 2.3 解析 graph_3 ~ graph_n 明细（跳过 graph_1 / graph_2）
    detail_files = []
    for f in files:
        m = re.match(r"graph_(\d+)\.(htm|html|txt)$", f, re.I)
        if m and int(m.group(1)) >= 3:
            detail_files.append((int(m.group(1)), f))
    detail_files.sort()

    all_detail = []
    for num, f in detail_files:
        gname = f"graph_{num}"
        recs = parse_graph_detail(htm(f))
        if not recs:
            print(f"  [跳过] {f}: 未找到 ScanData 表")
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
        all_detail.extend(recs)
        print(f"  [解析] {f} -> {len(recs)} 行")

    # 2.4 组装 Excel：以“工作表”模板为结构
    writer = pd.ExcelWriter("EMC汇总结果.xlsx", engine="openpyxl")

    # Sheet1: 模板（照搬“工作表”样式）
    template = pd.DataFrame({
        "test date": ["", "", "SN", "CFG", "item"],
        "2026/10/12": ["", "", "SN1", "CFG1", "test mode1"],
        "": ["", "", "", "", "test mode2"],
    })
    pd.DataFrame({
        "test date": ["2026/10/12", "SN", "CFG", "item", "test mode1", "test mode2"],
        "2026/10/13": ["", "SN1", "CFG1", "freq", 1, 2],
        "": ["", "SN2", "CFG2", "margin", 1, 2],
    }).to_excel(writer, sheet_name="汇总模板", index=False)

    # Sheet2: 最差 margin 汇总（来自 graph_0）
    if summary:
        sdf = pd.DataFrame(summary)
        sdf.to_excel(writer, sheet_name="最差Margin汇总", index=False)

    # Sheet3: 全部频点明细（graph_3~n 拼接）
    if all_detail:
        ddf = pd.DataFrame(all_detail)
        cols = ["graph", "detector", "unit_config", "sn", "test_case",
                "Frequency(MHz)", "Level(dBuV)", "Limit(dBuV)", "Margin(dB)",
                "Raw(dBuV)", "LISN(dB)", "Cable(dB)", "Xducer(dB)", "AllFactors(dB)"]
        ddf = ddf[[c for c in cols if c in ddf.columns]]
        ddf.to_excel(writer, sheet_name="频点明细", index=False)

    writer.close()
    print(f"\n完成：明细 {len(all_detail)} 行，最差汇总 {len(summary)} 行 -> EMC汇总结果.xlsx")


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    build(folder)
