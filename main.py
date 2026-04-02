#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import shutil
import traceback
from datetime import datetime

import pandas as pd
from flask import Flask, request, jsonify, send_from_directory, render_template
from openpyxl import load_workbook


# ===================== 基础配置 =====================

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

LOG_FILE = os.path.join(OUTPUT_DIR, "运行日志.txt")
OUTPUT_RESULT_XLSX = os.path.join(OUTPUT_DIR, "结果.xlsx")
OUTPUT_REPORT_TXT = os.path.join(OUTPUT_DIR, "通报.txt")
OUTPUT_TEMPLATE_XLSX = os.path.join(OUTPUT_DIR, "通报模板.xlsx")
OUTPUT_STANDARD_XLSX = os.path.join(OUTPUT_DIR, "标准销量表.xlsx")
FIXED_TEMPLATE_FILE = os.path.join(ASSETS_DIR, "通报模板.xlsx")

# 到客户 sheet 固定列
CUSTOMER_COLS = {
    "DAY_P1": 9,
    "D5_P1": 12,
    "DAY_P2": 18,
    "D5_P2": 21,
}

# 到门店 sheet 固定列
STORE_COLS = {
    "DAY_P1": 5,
    "D5_P1": 8,
    "DAY_P2": 12,
    "D5_P2": 15,
}

app = Flask(__name__)


# ===================== 通用工具 =====================

def ensure_dirs():
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(ASSETS_DIR, exist_ok=True)


def reset_log():
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("start\n")


def log(msg: str):
    now = datetime.now().strftime("%H:%M:%S")
    line = f"[{now}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def check_fixed_template():
    if not os.path.exists(FIXED_TEMPLATE_FILE):
        log("警告：assets/通报模板.xlsx 不存在，网页端模板下载将不可用")


def safe_str(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def norm(x) -> str:
    s = safe_str(x)
    s = s.lower()
    s = s.replace("　", " ").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]", "", s)
    return s


def detect_column(df: pd.DataFrame, candidates):
    norm_cols = {norm(c): c for c in df.columns}
    for cand in candidates:
        nc = norm(cand)
        if nc in norm_cols:
            return norm_cols[nc]
    return None


def extract_brandless(text: str) -> str:
    s = norm(text)
    for word in ["huawei", "华为", "荣耀", "honor", "oppo", "vivo", "xiaomi", "redmi", "apple", "苹果"]:
        s = s.replace(word, "")
    return s


def fuzzy_match_product(raw_name: str, p1: str, p2: str):
    raw = norm(raw_name)
    raw2 = extract_brandless(raw_name)

    p1n = norm(p1)
    p2n = norm(p2)
    p1b = extract_brandless(p1)
    p2b = extract_brandless(p2)

    def matched(src_a, src_b, target_a, target_b):
        if not target_a:
            return False
        return (
            target_a in src_a or src_a in target_a or
            target_a in src_b or src_b in target_a or
            (target_b and (target_b in src_a or src_a in target_b or target_b in src_b or src_b in target_b))
        )

    if matched(raw, raw2, p1n, p1b):
        return "P1"
    if matched(raw, raw2, p2n, p2b):
        return "P2"
    return None


def contains_match(text: str, candidates):
    nt = norm(text)
    if not nt:
        return ""

    for item in candidates:
        if norm(item) == nt:
            return item

    for item in candidates:
        ni = norm(item)
        if ni and (ni in nt or nt in ni):
            return item

    return ""


def customer_group_name(name: str) -> str:
    s = safe_str(name)
    if "联启" in s:
        return "联启"
    if "朝龙" in s:
        return "朝龙"
    if "同文" in s:
        return "同文"
    if "九机" in s:
        return "九机"
    return "其他客户"


def store_short_name(name: str) -> str:
    s = safe_str(name)
    s = re.sub(r"^(云南省)?曲靖市宣威市", "", s)
    s = s.replace("联启通讯", "")
    s = s.replace("华为授权体验店", "体验店")
    s = s.replace("华为体验店", "体验店")
    return s.strip() or safe_str(name)


def format_dt_cn(dt: pd.Timestamp) -> str:
    dt = pd.to_datetime(dt)
    return f"{dt.month}月{dt.day}日"


def copy_template_to_output(template_path: str):
    shutil.copyfile(template_path, OUTPUT_TEMPLATE_XLSX)


# ===================== 读取模板 =====================

def load_template(template_file: str):
    wb = load_workbook(template_file)
    if "到客户" not in wb.sheetnames or "到门店" not in wb.sheetnames:
        raise ValueError("模板错误：必须包含【到客户】和【到门店】两个工作表")

    ws_customer = wb["到客户"]
    ws_store = wb["到门店"]

    customers = []
    for r in range(2, ws_customer.max_row + 1):
        name = safe_str(ws_customer.cell(r, 1).value)
        if name and "合计" not in name and "总计" not in name:
            customers.append(name)

    stores = []
    store_to_guide = {}
    for r in range(3, ws_store.max_row + 1):
        store_name = safe_str(ws_store.cell(r, 1).value)
        guide_name = safe_str(ws_store.cell(r, 2).value)
        if store_name and "合计" not in store_name and "总计" not in store_name:
            stores.append(store_name)
            store_to_guide[store_name] = guide_name

    return wb, ws_customer, ws_store, customers, stores, store_to_guide


# ===================== 读取原表 =====================

def read_data(data_file: str) -> pd.DataFrame:
    df = pd.read_excel(data_file)

    customer_col = detect_column(df, ["所属客户渠道名称", "客户渠道名称", "客户名称", "渠道名称"])
    store_col = detect_column(df, ["门店名称", "门店", "门店名"])
    sn_col = detect_column(df, ["SN", "sn", "sn码", "sn号"])
    date_col = detect_column(df, ["销售日期", "日期", "激活日期", "开票日期"])
    promoter_col = detect_column(df, ["促销员姓名", "促销员", "导购", "导购员"])
    seller_col = detect_column(df, ["销售人姓名", "销售人", "店员姓名"])

    if not customer_col or not store_col or not sn_col:
        raise ValueError("原表缺少必要字段：客户 / 门店 / SN")

    product_cols = []
    for c in ["传播名", "产品型号", "Offering别称", "产品名称", "机型名称", "商品名称"]:
        real = detect_column(df, [c])
        if real:
            product_cols.append(real)

    if not product_cols:
        raise ValueError("原表缺少产品字段：传播名 / 产品型号 / Offering别称 / 产品名称")

    out = pd.DataFrame()
    out["原始客户"] = df[customer_col].map(lambda x: safe_str(x).split("_")[0])
    out["原始门店"] = df[store_col].map(safe_str)
    out["SN"] = df[sn_col].map(safe_str)

    out["产品名称"] = ""
    for c in product_cols:
        out["产品名称"] = (out["产品名称"] + " " + df[c].fillna("").astype(str)).str.strip()

    if promoter_col and seller_col:
        out["导购"] = df.apply(
            lambda r: safe_str(r[promoter_col]) if safe_str(r[promoter_col]) else safe_str(r[seller_col]),
            axis=1
        )
    elif promoter_col:
        out["导购"] = df[promoter_col].map(safe_str)
    elif seller_col:
        out["导购"] = df[seller_col].map(safe_str)
    else:
        out["导购"] = ""

    if date_col:
        out["销售日期"] = pd.to_datetime(df[date_col], errors="coerce").dt.normalize()
    else:
        out["销售日期"] = pd.NaT

    return out


# ===================== 清洗 / 匹配 / 分期 =====================

def clean_and_match(df: pd.DataFrame, customers, stores, store_to_guide, p1: str, p2: str, launch_date: str):
    launch_dt = pd.to_datetime(launch_date).normalize()
    end_5d = launch_dt + pd.Timedelta(days=4)

    df = df.copy()

    df = df[df["产品名称"].map(lambda x: safe_str(x) != "")]
    df = df[df["SN"].map(lambda x: safe_str(x) != "")]

    df["产品分类"] = df["产品名称"].map(lambda x: fuzzy_match_product(x, p1, p2))
    df["模板客户"] = df["原始客户"].map(lambda x: contains_match(x, customers))
    df["模板门店"] = df["原始门店"].map(lambda x: contains_match(x, stores))
    df["模板导购"] = df["模板门店"].map(lambda x: store_to_guide.get(x, ""))

    if df["销售日期"].notna().any():
        df["首销日标记"] = (df["销售日期"] == launch_dt).astype(int)
        df["首销5日标记"] = ((df["销售日期"] >= launch_dt) & (df["销售日期"] <= end_5d)).astype(int)
    else:
        df["首销日标记"] = 1
        df["首销5日标记"] = 1

    valid = df[
        (df["产品分类"].notna()) &
        (df["模板客户"] != "") &
        (df["模板门店"] != "") &
        (df["SN"] != "")
    ].copy()

    day_df = valid[valid["首销日标记"] == 1].copy()
    day_df = day_df.drop_duplicates(subset=["模板门店", "SN", "产品分类"])
    day_df["数量"] = 1

    d5_df = valid[valid["首销5日标记"] == 1].copy()
    d5_df = d5_df.drop_duplicates(subset=["模板门店", "SN", "产品分类"])
    d5_df["数量"] = 1

    std = valid.drop_duplicates(subset=["模板门店", "SN", "产品分类"]).copy()
    std.to_excel(OUTPUT_STANDARD_XLSX, index=False)

    log(f"原始数据：{len(df)}")
    log(f"有效数据：{len(std)}")
    log(f"首销日有效：{len(day_df)}")
    log(f"首销5日有效：{len(d5_df)}")

    return day_df, d5_df, std


# ===================== 聚合 =====================

def build_agg(df: pd.DataFrame):
    result = {
        "region": {},
        "customer": {},
        "store": {},
        "guide": {},
    }

    if df.empty:
        return result

    result["region"] = df.groupby("产品分类")["数量"].sum().to_dict()
    result["customer"] = df.groupby(["模板客户", "产品分类"])["数量"].sum().to_dict()
    result["store"] = df.groupby(["模板门店", "产品分类"])["数量"].sum().to_dict()
    result["guide"] = df.groupby(["模板导购", "产品分类"])["数量"].sum().to_dict()
    return result


# ===================== 回填 Excel =====================

def fill_customer_sheet(ws, day_data, d5_data):
    for r in range(2, ws.max_row + 1):
        name = safe_str(ws.cell(r, 1).value)
        if not name or "合计" in name or "总计" in name:
            continue

        ws.cell(r, CUSTOMER_COLS["DAY_P1"]).value = day_data["customer"].get((name, "P1"), 0)
        ws.cell(r, CUSTOMER_COLS["D5_P1"]).value = d5_data["customer"].get((name, "P1"), 0)

        ws.cell(r, CUSTOMER_COLS["DAY_P2"]).value = day_data["customer"].get((name, "P2"), 0)
        ws.cell(r, CUSTOMER_COLS["D5_P2"]).value = d5_data["customer"].get((name, "P2"), 0)


def fill_store_sheet(ws, day_data, d5_data):
    for r in range(3, ws.max_row + 1):
        name = safe_str(ws.cell(r, 1).value)
        if not name or "合计" in name or "总计" in name:
            continue

        ws.cell(r, STORE_COLS["DAY_P1"]).value = day_data["store"].get((name, "P1"), 0)
        ws.cell(r, STORE_COLS["D5_P1"]).value = d5_data["store"].get((name, "P1"), 0)

        ws.cell(r, STORE_COLS["DAY_P2"]).value = day_data["store"].get((name, "P2"), 0)
        ws.cell(r, STORE_COLS["D5_P2"]).value = d5_data["store"].get((name, "P2"), 0)


# ===================== 文本通报 =====================

def build_group_customer(customers, result, prod_key):
    out = {
        "联启": 0,
        "朝龙": 0,
        "同文": 0,
        "九机": 0,
        "其他客户": 0,
    }
    for c in customers:
        grp = customer_group_name(c)
        out[grp] += result["customer"].get((c, prod_key), 0)
    return out


def build_store_lines(stores, result):
    rows = []
    for s in stores:
        p1 = result["store"].get((s, "P1"), 0)
        p2 = result["store"].get((s, "P2"), 0)
        total = p1 + p2
        rows.append({
            "name": s,
            "short": store_short_name(s),
            "p1": p1,
            "p2": p2,
            "total": total,
        })
    return rows


def build_guide_lines(store_to_guide, result):
    guide_names = []
    for _, guide in store_to_guide.items():
        g = safe_str(guide)
        if g and g != "/" and "合计" not in g and "总计" not in g and g not in guide_names:
            guide_names.append(g)

    rows = []
    for g in guide_names:
        p1 = result["guide"].get((g, "P1"), 0)
        p2 = result["guide"].get((g, "P2"), 0)
        rows.append({
            "name": g,
            "p1": p1,
            "p2": p2,
            "total": p1 + p2,
        })
    return rows


def format_report_text(title: str, p1_name: str, p2_name: str, customers, stores, store_to_guide, result, low_threshold=2):
    region_p1 = result["region"].get("P1", 0)
    region_p2 = result["region"].get("P2", 0)
    total = region_p1 + region_p2

    cust_p1 = build_group_customer(customers, result, "P1")
    cust_p2 = build_group_customer(customers, result, "P2")

    store_rows = build_store_lines(stores, result)
    guide_rows = build_guide_lines(store_to_guide, result)

    top5 = sorted(store_rows, key=lambda x: x["total"], reverse=True)[:5]
    low_rows = [x for x in store_rows if x["total"] < low_threshold]

    lines = []
    lines.append(title)
    lines.append("")
    lines.append("【区域整体】")
    lines.append(f"{p1_name}：{region_p1}台")
    lines.append(f"{p2_name}：{region_p2}台")
    lines.append(f"合计：{total}台")
    lines.append("")

    lines.append("【客户】")
    for grp in ["联启", "朝龙", "同文", "九机", "其他客户"]:
        lines.append(f"{grp}：{p1_name}：{cust_p1[grp]}台，{p2_name}：{cust_p2[grp]}台")
    lines.append("")

    lines.append("【门店】")
    for row in store_rows:
        lines.append(f"{row['short']}：{p1_name}：{row['p1']}台，{p2_name}：{row['p2']}台")
    lines.append("")

    lines.append("【导购】")
    for row in guide_rows:
        lines.append(f"{row['name']}：{p1_name}：{row['p1']}台，{p2_name}：{row['p2']}台")
    lines.append("")

    lines.append("【TOP门店】")
    if top5:
        for row in top5:
            lines.append(f"{row['short']}：{row['total']}台")
    else:
        lines.append("无")
    lines.append("")

    lines.append("【落后门店】")
    if low_rows:
        for row in low_rows:
            lines.append(f"{row['short']}：{row['total']}台")
    else:
        lines.append("无")

    return "\n".join(lines)


def build_summary_text(p1_name, p2_name, day_result, d5_result, stores):
    day_p1 = day_result["region"].get("P1", 0)
    day_p2 = day_result["region"].get("P2", 0)
    d5_p1 = d5_result["region"].get("P1", 0)
    d5_p2 = d5_result["region"].get("P2", 0)

    day_store_rows = build_store_lines(stores, day_result)
    top5_day = sorted(day_store_rows, key=lambda x: x["total"], reverse=True)[:5]

    top_txt = "、".join([f"{x['short']} {x['total']}台" for x in top5_day if x["total"] > 0]) or "无"

    return (
        f"首销日：{p1_name} {day_p1}台，{p2_name} {day_p2}台，合计 {day_p1 + day_p2}台\n"
        f"首销5日：{p1_name} {d5_p1}台，{p2_name} {d5_p2}台，合计 {d5_p1 + d5_p2}台\n"
        f"TOP门店：{top_txt}"
    )


# ===================== 主处理流程 =====================

def main_process(data_file, template_file, launch_date, p1_name, p2_name):
    reset_log()
    log(f"模板：{template_file}")
    log(f"数据：{data_file}")
    log(f"首销日期：{launch_date}")
    log(f"产品1：{p1_name}")
    log(f"产品2：{p2_name}")

    copy_template_to_output(template_file)

    wb, ws_customer, ws_store, customers, stores, store_to_guide = load_template(template_file)
    df = read_data(data_file)
    day_df, d5_df, _ = clean_and_match(df, customers, stores, store_to_guide, p1_name, p2_name, launch_date)

    day_result = build_agg(day_df)
    d5_result = build_agg(d5_df)

    fill_customer_sheet(ws_customer, day_result, d5_result)
    fill_store_sheet(ws_store, day_result, d5_result)
    wb.save(OUTPUT_RESULT_XLSX)

    launch_dt = pd.to_datetime(launch_date).normalize()
    end_5d = launch_dt + pd.Timedelta(days=4)

    report_day = format_report_text(
        title=f"【首销通报】{format_dt_cn(launch_dt)}",
        p1_name=p1_name,
        p2_name=p2_name,
        customers=customers,
        stores=stores,
        store_to_guide=store_to_guide,
        result=day_result,
    )

    report_5day = format_report_text(
        title=f"【首销通报】{format_dt_cn(launch_dt)}-{format_dt_cn(end_5d)}",
        p1_name=p1_name,
        p2_name=p2_name,
        customers=customers,
        stores=stores,
        store_to_guide=store_to_guide,
        result=d5_result,
    )

    with open(OUTPUT_REPORT_TXT, "w", encoding="utf-8") as f:
        f.write(report_day + "\n\n" + report_5day)

    summary = build_summary_text(p1_name, p2_name, day_result, d5_result, stores)

    log("结果Excel输出完成")
    log("文本通报输出完成")
    log("标准销量表输出完成")
    log("全部完成")

    return summary, report_day, report_5day


# ===================== Flask 路由 =====================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    try:
        ensure_dirs()

        data_file = request.files.get("data_file")
        template_file = request.files.get("template_file")
        launch_date = safe_str(request.form.get("launch_date"))
        p1_name = safe_str(request.form.get("p1")) or "产品1"
        p2_name = safe_str(request.form.get("p2")) or "产品2"

        if not data_file or not template_file:
            return jsonify({"ok": False, "msg": "请上传微服务原表和模板文件"})

        if not data_file.filename or not template_file.filename:
            return jsonify({"ok": False, "msg": "请选择文件"})

        if not launch_date:
            return jsonify({"ok": False, "msg": "请选择首销日期"})

        data_path = os.path.join(UPLOAD_DIR, data_file.filename)
        template_path = os.path.join(UPLOAD_DIR, template_file.filename)

        data_file.save(data_path)
        template_file.save(template_path)

        summary, report_day, report_5day = main_process(
            data_file=data_path,
            template_file=template_path,
            launch_date=launch_date,
            p1_name=p1_name,
            p2_name=p2_name,
        )

        return jsonify({
            "ok": True,
            "msg": "处理完成",
            "summary": summary,
            "report_day": report_day,
            "report_5day": report_5day,
        })

    except Exception as e:
        ensure_dirs()
        err = str(e) + "\n" + traceback.format_exc()
        log(err)
        return jsonify({"ok": False, "msg": err})


@app.route("/download/<path:filename>")
def download(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.route("/download-template")
def download_template():
    if not os.path.exists(FIXED_TEMPLATE_FILE):
        return jsonify({"ok": False, "msg": "标准模板不存在，请先在 assets 目录放入：通报模板.xlsx"}), 404

    return send_from_directory(
        ASSETS_DIR,
        "通报模板.xlsx",
        as_attachment=True
    )


if __name__ == "__main__":
    ensure_dirs()
    reset_log()
    check_fixed_template()
    app.run(debug=True)