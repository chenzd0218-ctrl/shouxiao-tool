import os
import glob
import traceback
from collections import OrderedDict
from uuid import uuid4

import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory
from openpyxl import load_workbook

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")

LOG_FILE = os.path.join(OUTPUT_DIR, "运行日志.txt")
OUTPUT_RESULT_XLSX = os.path.join(OUTPUT_DIR, "结果.xlsx")

CUSTOMER_SHEET = "到客户"
STORE_SHEET = "到门店"

app = Flask(__name__)


def ensure_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)


def reset_log():
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("start\n")


def log(msg: str):
    print(msg)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()


def norm(x):
    return (
        safe_str(x)
        .replace(" ", "")
        .replace("　", "")
        .replace("\n", "")
        .replace("\r", "")
        .lower()
    )


def parse_products(product_list):
    products = [safe_str(x) for x in product_list if safe_str(x)]
    products = list(dict.fromkeys(products))
    if not products:
        raise ValueError("请至少填写1个产品")
    if len(products) > 4:
        raise ValueError("最多支持4个产品")
    return products


def detect_column(df: pd.DataFrame, candidates):
    col_map = {norm(c): c for c in df.columns}
    for c in candidates:
        nc = norm(c)
        if nc in col_map:
            return col_map[nc]
    return None


def match_product(name: str, products):
    n = norm(name)
    for p in products:
        if norm(p) in n:
            return p
    return None


def load_data(file, products):
    df = pd.read_excel(file)

    customer_col = detect_column(df, ["所属客户渠道名称", "客户渠道名称", "客户名称", "渠道名称"])
    store_col = detect_column(df, ["门店名称", "门店", "门店名"])
    sn_col = detect_column(df, ["SN", "sn", "sn码", "sn号"])
    promoter_col = detect_column(df, ["促销员姓名", "促销员", "导购", "导购员"])
    seller_col = detect_column(df, ["销售人姓名", "销售人", "店员姓名"])
    product_col = (
        detect_column(df, ["传播名"])
        or detect_column(df, ["产品型号"])
        or detect_column(df, ["Offering别称"])
        or detect_column(df, ["产品名称"])
        or detect_column(df, ["机型名称"])
        or detect_column(df, ["商品名称"])
    )

    if not customer_col:
        raise ValueError("原表缺少客户字段")
    if not store_col:
        raise ValueError("原表缺少门店字段")
    if not sn_col:
        raise ValueError("原表缺少SN字段")
    if not product_col:
        raise ValueError("原表缺少产品字段")

    out = pd.DataFrame()
    out["客户名称"] = df[customer_col].map(lambda x: safe_str(x).split("_")[0])
    out["门店名称"] = df[store_col].map(safe_str)
    out["SN"] = df[sn_col].map(safe_str)
    out["产品名称"] = df[product_col].map(safe_str)

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

    out["产品归类"] = out["产品名称"].map(lambda x: match_product(x, products))

    out = out[
        (out["产品归类"].notna()) &
        (out["客户名称"] != "") &
        (out["门店名称"] != "") &
        (out["SN"] != "")
    ].copy()

    out = out.drop_duplicates(subset=["客户名称", "门店名称", "产品归类", "SN"]).copy()
    return out


def build_customer_rows(df, products):
    ordered_customers = list(OrderedDict.fromkeys(df["客户名称"].tolist()))

    rows = []
    for customer in ordered_customers:
        sub = df[df["客户名称"] == customer]
        row = {"客户名称": customer}
        total = 0
        for p in products:
            qty = int((sub["产品归类"] == p).sum())
            row[p] = qty
            total += qty
        row["合计"] = total
        rows.append(row)

    rows = sorted(rows, key=lambda x: x["合计"], reverse=True)
    return rows


def build_store_rows(df, products):
    ordered_stores = list(OrderedDict.fromkeys(df["门店名称"].tolist()))

    rows = []
    for store in ordered_stores:
        sub = df[df["门店名称"] == store]
        guide = safe_str(sub.iloc[0]["导购"]) if len(sub) > 0 else ""
        row = {"门店名称": store, "导购": guide}
        total = 0
        for p in products:
            qty = int((sub["产品归类"] == p).sum())
            row[p] = qty
            total += qty
        row["合计"] = total
        rows.append(row)

    rows = sorted(rows, key=lambda x: x["合计"], reverse=True)
    return rows


def write_customer_sheet(ws, rows, products):
    headers = ["客户名称"] + products + ["合计"]
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c).value = h

    for r_idx, row in enumerate(rows, start=2):
        ws.cell(r_idx, 1).value = row["客户名称"]
        col = 2
        for p in products:
            ws.cell(r_idx, col).value = row[p]
            col += 1
        ws.cell(r_idx, col).value = row["合计"]


def write_store_sheet(ws, rows, products):
    headers = ["门店名称", "导购"] + products + ["合计"]
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c).value = h

    for r_idx, row in enumerate(rows, start=2):
        ws.cell(r_idx, 1).value = row["门店名称"]
        ws.cell(r_idx, 2).value = row["导购"]
        col = 3
        for p in products:
            ws.cell(r_idx, col).value = row[p]
            col += 1
        ws.cell(r_idx, col).value = row["合计"]


def process_report(data_file, template_file, products):
    reset_log()
    log(f"本次产品：{products}")
    log(f"模板：{template_file}")
    log(f"数据：{data_file}")

    if not os.path.exists(template_file):
        raise FileNotFoundError(f"未找到模板文件：{template_file}")

    wb = load_workbook(template_file)

    if CUSTOMER_SHEET not in wb.sheetnames:
        raise ValueError(f"模板缺少工作表：{CUSTOMER_SHEET}")
    if STORE_SHEET not in wb.sheetnames:
        raise ValueError(f"模板缺少工作表：{STORE_SHEET}")

    ws_customer = wb[CUSTOMER_SHEET]
    ws_store = wb[STORE_SHEET]

    df = load_data(data_file, products)
    log(f"有效记录数：{len(df)}")

    customer_rows = build_customer_rows(df, products)
    store_rows = build_store_rows(df, products)

    log(f"客户数：{len(customer_rows)}")
    log(f"门店数：{len(store_rows)}")

    write_customer_sheet(ws_customer, customer_rows, products)
    write_store_sheet(ws_store, store_rows, products)

    out_name = f"结果_{uuid4().hex[:8]}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    wb.save(out_path)

    log("结果Excel输出完成")
    log("全部完成")

    return {
        "customer_rows": customer_rows,
        "store_rows": store_rows,
        "products": products,
        "result_file": out_name,
        "log_file": os.path.basename(LOG_FILE),
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process():
    try:
        ensure_dir()

        data_file = request.files.get("data_file")
        template_file = request.files.get("template_file")

        if not data_file or not template_file:
            return jsonify({"ok": False, "msg": "请上传微服务原表和模板文件"})

        products = parse_products([
            request.form.get("product1", ""),
            request.form.get("product2", ""),
            request.form.get("product3", ""),
            request.form.get("product4", ""),
        ])

        data_name = f"data_{uuid4().hex[:8]}_{data_file.filename}"
        template_name = f"template_{uuid4().hex[:8]}_{template_file.filename}"

        data_path = os.path.join(UPLOAD_DIR, data_name)
        template_path = os.path.join(UPLOAD_DIR, template_name)

        data_file.save(data_path)
        template_file.save(template_path)

        result = process_report(data_path, template_path, products)

        return jsonify({
            "ok": True,
            "msg": "处理完成",
            **result
        })

    except Exception as e:
        err = str(e)
        tb = traceback.format_exc()
        log(err)
        log(tb)
        return jsonify({"ok": False, "msg": err})


@app.route("/download/<path:filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)