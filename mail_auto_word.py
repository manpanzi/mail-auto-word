# -*- coding: utf-8 -*-
"""
mail_auto_word.py — 邮件监控 & 自动生成 Word 并回复
=====================================================
功能：
  1. 通过 IMAP 监控指定邮箱，定时检查新邮件
  2. 识别主题为空 + 带 Excel 附件（.xlsx/.xls）的邮件
  3. 解析邮件附件中的 Excel 表格（.xlsx），或正文中的 CSV 数据
  4. 自动识别数据类型（出库通知单 / 提货委托函），生成 Word 文档
  5. 将生成的 .docx 作为附件回复给发件人

支持邮箱：QQ 邮箱（@qq.com）、163 邮箱（@163.com）

用法：
  python mail_auto_word.py          # 前台运行，Ctrl+C 停止
  python mail_auto_word.py --once   # 只检查一次就退出
"""

import os
import re
import sys
import csv
import json
import time
import logging
import zipfile
import argparse
import tempfile
import threading
import email.policy
from io import StringIO
from datetime import datetime
from pathlib import Path

# ============================================================
#  依赖检查（首次运行提示安装）
# ============================================================
MISSING = []
try:
    import docx
    from docx import Document
    from docx.shared import Pt, Cm, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
except ImportError:
    MISSING.append("python-docx")

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    MISSING.append("openpyxl")

if MISSING:
    print("Missing deps: " + ", ".join(MISSING))
    print("Run: pip install " + " ".join(MISSING))
    sys.exit(1)

# ============================================================
#  配置区 — 按需修改
# ============================================================

# --- 邮箱账号 ---
# QQ 邮箱 / 163 邮箱均可，使用 授权码（非登录密码）登录
# 如何获取授权码：
#   QQ 邮箱：设置 → 账户 → POP3/IMAP/SMTP 服务 → 开启 IMAP/SMTP → 生成授权码
#   163 邮箱：设置 → POP3/SMTP/IMAP → 开启 IMAP/SMTP → 新增授权码
# --- 邮箱账号（优先从环境变量读取，适用于云端部署）---
ACCOUNT = {
    "email":    os.environ.get("MAIL_USER", "357818590@qq.com"),
    "password": os.environ.get("MAIL_PASS", "ukkdcvbewfesbhdc"),
}

# --- IMAP 收件服务器（根据邮箱自动选择，一般无需修改）---
# QQ:  imap.qq.com,  993, SSL
# 163: imap.163.com, 993, SSL
IMAP_CONFIG = {
    "qq.com":  {"host": "imap.qq.com",  "port": 993},
    "163.com": {"host": "imap.163.com", "port": 993},
}

# --- SMTP 发件服务器 ---
# QQ:  smtp.qq.com,  465, SSL
# 163: smtp.163.com, 465, SSL
SMTP_CONFIG = {
    "qq.com":  {"host": "smtp.qq.com",  "port": 465},
    "163.com": {"host": "smtp.163.com", "port": 465},
}

# --- 监控设置 ---
CHECK_INTERVAL = 60          # 检查间隔（秒），建议 ≥ 30
MARK_AS_READ   = True        # 处理完成后是否标记为已读
TRACK_FILE     = "processed_uids.json"  # 记录已处理邮件的 UID，避免重复处理

# --- 邮件识别 ---
SUBJECT_KEYWORDS = ["GENWORD"]  # 已废弃，现在通过 Excel 附件匹配

# --- 日志 ---
LOG_FILE = "mail_auto_word.log"

# ============================================================
#  日志配置
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

# ============================================================
#  工具函数
# ============================================================

def get_mail_domain(email_addr):
    """从邮箱地址提取域名，如 'user@qq.com' → 'qq.com'"""
    m = re.search(r"@(.+)$", email_addr)
    return m.group(1).lower() if m else None


def get_imap_config(email_addr):
    domain = get_mail_domain(email_addr)
    if domain in IMAP_CONFIG:
        return IMAP_CONFIG[domain]
    # 尝试通用推断
    if "qq" in domain:
        return IMAP_CONFIG["qq.com"]
    if "163" in domain:
        return IMAP_CONFIG["163.com"]
    raise ValueError(f"不支持的邮箱域名：{domain}，目前支持 QQ 邮箱和 163 邮箱。")


def get_smtp_config(email_addr):
    domain = get_mail_domain(email_addr)
    if domain in SMTP_CONFIG:
        return SMTP_CONFIG[domain]
    if "qq" in domain:
        return SMTP_CONFIG["qq.com"]
    if "163" in domain:
        return SMTP_CONFIG["163.com"]
    raise ValueError(f"不支持的邮箱域名：{domain}，目前支持 QQ 邮箱和 163 邮箱。")


# ---- UID 追踪 ----
def load_processed_uids(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()


def save_processed_uids(filepath, uid_set):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(list(uid_set), f)


# ============================================================
#  邮件解析
# ============================================================

def parse_email_body(msg):
    """
    从 email.Message 中提取纯文本正文。
    优先 text/plain，其次 text/html（简单去除标签）。
    """
    import email as em

    body_plain = ""
    body_html = ""

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdisp = str(part.get("Content-Disposition", ""))
            if "attachment" in cdisp:
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            charset = part.get_content_charset() or "utf-8"
            try:
                text = payload.decode(charset, errors="replace")
            except Exception:
                text = payload.decode("utf-8", errors="replace")
            if ctype == "text/plain":
                body_plain += text
            elif ctype == "text/html":
                body_html += text
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            try:
                body_plain = payload.decode(charset, errors="replace")
            except Exception:
                body_plain = payload.decode("utf-8", errors="replace")

    if body_plain:
        return body_plain.strip()
    if body_html:
        # 简易 HTML → 文本
        text = re.sub(r"<br\s*/?>", "\n", body_html, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(r"&nbsp;", " ", text)
        text = re.sub(r"&lt;", "<", text)
        text = re.sub(r"&gt;", ">", text)
        text = re.sub(r"&amp;", "&", text)
        return text.strip()
    return ""


def parse_csv_from_body(body_text):
    """
    从邮件正文中解析 CSV 表格数据。
    支持：
      1. 第一行是表头（中文列名）
      2. 后续每行是一条数据
      3. 列之间用逗号、制表符或中文全角逗号分隔
    返回 (headers, records) 或 (None, []) 如果解析失败。
    """
    if not body_text:
        return None, []

    lines = [l.strip() for l in body_text.splitlines() if l.strip()]
    if len(lines) < 2:
        return None, []

    # 跳过非表格行（如 #开头、类型标记等）
    data_lines = []
    for l in lines:
        if l.startswith("#") or l.startswith("类型") or l.startswith("模板"):
            continue
        data_lines.append(l)

    if len(data_lines) < 2:
        return None, []

    # 自动检测分隔符：优先逗号，其次 tab
    sample = data_lines[0]
    if "\t" in sample:
        sep = "\t"
    elif "，" in sample:
        sep = "，"
    elif "," in sample:
        sep = ","
    else:
        return None, []

    # 解析表头
    headers = [h.strip() for h in data_lines[0].split(sep)]
    if not headers:
        return None, []

    # 解析数据行
    records = []
    for line in data_lines[1:]:
        vals = [v.strip() for v in line.split(sep)]
        if len(vals) < len(headers):
            vals += [""] * (len(headers) - len(vals))
        elif len(vals) > len(headers):
            vals = vals[:len(headers)]
        rec = {}
        for i, h in enumerate(headers):
            rec[h] = vals[i] if i < len(vals) else ""
        records.append(rec)

    return headers, records


def extract_excel_attachments(msg):
    """
    从邮件中提取 Excel 附件（.xlsx / .xls），保存到临时文件。
    返回 [filepath, ...]
    """
    import email as em

    temp_files = []
    if not msg.is_multipart():
        return temp_files

    for part in msg.walk():
        cdisp = str(part.get("Content-Disposition", ""))
        if "attachment" not in cdisp:
            continue
        filename = part.get_filename()
        if not filename:
            continue
        # 解码中文文件名
        try:
            from email.header import decode_header
            dh = decode_header(filename)
            filename = "".join(
                t.decode(c or "utf-8", errors="replace") if isinstance(t, bytes) else str(t)
                for t, c in dh
            )
        except Exception:
            pass

        if not (filename.lower().endswith(".xlsx") or filename.lower().endswith(".xls")):
            continue

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp.write(payload)
        tmp.close()
        temp_files.append(tmp.name)
        log.info(f"提取 Excel 附件：{filename} → {tmp.name}")

    return temp_files


def parse_excel_data(filepath):
    """
    从 Excel 文件读取数据。
    第一行为表头，后续每行为一条数据。
    返回 (headers, records) 或 (None, [])。
    """
    try:
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        wb.close()

        if len(rows) < 2:
            return None, []

        headers = [str(h).strip() if h else "" for h in rows[0]]
        if not any(headers):
            return None, []

        records = []
        for row in rows[1:]:
            if row[0] is None and all(v is None for v in row):
                continue
            rec = {}
            for i, h in enumerate(headers):
                val = row[i] if i < len(row) else ""
                if val is None:
                    val = ""
                rec[h] = str(val).strip()
            # 过滤：跳过信息不完整的行（只有序号/备注等少于2个有效字段）
            skip_keywords = ["序号", "备注"]
            filled = sum(1 for k, v in rec.items() if v and not any(s in k for s in skip_keywords))
            if filled < 2:
                log.info(f"  跳过不完整行: {str(rec)[:100]}")
                continue
            records.append(rec)

        wb.close()
        log.info(f"Excel 表头: {headers}")
        if records:
            log.info(f"Excel 首行: {records[0]}")
        return headers, records

    except Exception as e:
        log.warning(f"解析 Excel 失败：{e}")
        return None, []


def detect_doc_type(headers):
    """
    根据表头自动判断文档类型。
    出库通知单特征列：目的港、船号、计划量
    提货委托函特征列：公司名、提货计划量、船舶联系人
    """
    出库通知单_keys = ["目的港", "船号", "计划量"]
    提货委托函_keys = ["公司名", "提货计划量", "船舶联系人"]

    score_a = sum(1 for k in 出库通知单_keys if any(k in h for h in headers))
    score_b = sum(1 for k in 提货委托函_keys if any(k in h for h in headers))

    if score_b > score_a:
        return "提货委托函"
    if score_a > 0:
        return "出库通知单"
    # fallback: 根据列数推断
    if len(headers) <= 5:
        return "出库通知单"
    return "提货委托函"


# ============================================================
#  Word 生成 — 出库通知单
# ============================================================

def _get_field(rec, *keys):
    """从 record 中取字段，支持多个候选 key 和部分匹配"""
    for k in keys:
        if k in rec:
            return rec[k]
    # 部分匹配：key 包含在表头中
    for k in keys:
        for h in rec:
            if k in h:
                return rec[h]
    return ""


def generate_chukutongzhidan(records, date_chinese, date_num):
    """
    生成「油品装船出库通知单」Word 文档。
    所有 records 放在同一个表格中，每行一个出库编号。
    返回临时文件路径。
    """
    doc = Document()

    # 默认页边距
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)

    style = doc.styles["Normal"]
    font = style.font
    font.name = "仿宋"
    font.size = Pt(12)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("上海久联集团有限公司")
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "仿宋"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("油品装船出库通知单")
    run.bold = True
    run.font.size = Pt(16)
    run.font.name = "仿宋"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    # 发件人信息
    doc.add_paragraph(
        f"发件人：宋南希        电话：18018586154                 日期：{date_chinese}"
    )

    doc.add_paragraph(" 近期将由以下船舶到油库装油，请提前安排船名申报、装船出库等相关事宜。")

    doc.add_paragraph()  # 空行

    # 表格
    headers = ["出库编号", "目的港", "船号", "油品品号", "作业罐号", "计划量（吨）", "备注", "允许装货时间（含）"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        for para in hdr_cells[i].paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.name = "仿宋"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    # 数据行
    for idx, rec in enumerate(records):
        counter = 99 + idx
        qty = _safe_float(_get_field(rec, "计划量", "提货计划量"))
        code = f"ZX-C0-0999-{date_num}-{qty:.2f}-{counter:03d}"

        row = table.add_row()
        values = [
            code,
            _get_field(rec, "目的港"),
            _get_field(rec, "船号", "船名"),
            "0号车用柴油（VI）",
            "T107/109",
            f"{qty:.2f}",
            _get_field(rec, "备注"),
            _get_field(rec, "允许装货时间"),
        ]
        for i, v in enumerate(values):
            row.cells[i].text = v
            for para in row.cells[i].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.name = "仿宋"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    doc.add_paragraph()
    doc.add_paragraph(f"制表：宋南希                              复核：")

    # 保存到临时文件
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    doc.save(tmp.name)
    return tmp.name


# ============================================================
#  Word 生成 — 提货委托函
# ============================================================

def generate_tihuoweituohan(rec, date_chinese):
    """
    生成「提货委托函」Word 文档（单条记录）。
    返回临时文件路径。
    """
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    style = doc.styles["Normal"]
    font = style.font
    font.name = "仿宋"
    font.size = Pt(14)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")

    def add_para(text, bold=False, size=14, align="left", font_name="仿宋"):
        p = doc.add_paragraph()
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        return p

    qty = _safe_float(_get_field(rec, "提货计划量", "计划量"))
    company = _get_field(rec, "公司名") or "上海久联集团石油化工有限公司"
    ship = _get_field(rec, "船名", "船号")
    dest = _get_field(rec, "目的港")
    contact = _get_field(rec, "船舶联系人")
    phone = _get_field(rec, "联系电话")
    eta = _get_field(rec, "预计到港时间")
    remark = _get_field(rec, "备注")

    add_para("提货委托函", bold=True, size=16, align="center")
    doc.add_paragraph()
    add_para(f"{company}：")
    add_para(
        f"    我公司委托船只到贵公司指定码头提取0号车用柴油（VI）{qty:.2f}吨，具体船舶信息如下："
    )

    # 信息表格
    table = doc.add_table(rows=9, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def set_cell(row, col, text, width=None, span=1, bold=False):
        cell = table.rows[row].cells[col]
        cell.text = text
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.size = Pt(10)
                run.font.name = "仿宋"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
                run.bold = bold
        # 合并单元格
        if span > 1:
            for j in range(1, span):
                cell.merge(table.rows[row].cells[col + j])
        return cell

    # 填充表格数据（模板固定结构）
    set_cell(0, 0, "船名", bold=True);          set_cell(0, 1, ship, span=3)
    set_cell(1, 0, "品名", bold=True);          set_cell(1, 1, "0号车用柴油（VI）", span=3)
    set_cell(2, 0, "提货计划量", bold=True);     set_cell(2, 1, f"{qty:.2f}吨", span=3)
    set_cell(3, 0, "目的港", bold=True);         set_cell(3, 1, dest, span=3)
    set_cell(4, 0, "下一港", bold=True);         set_cell(4, 1, "上海石洞口", span=3)
    set_cell(5, 0, "船舶联系人", bold=True);     set_cell(5, 1, contact); set_cell(5, 2, "联系电话", bold=True); set_cell(5, 3, phone)
    set_cell(6, 0, "业务联系人", bold=True);     set_cell(6, 1, "龚耀兵"); set_cell(6, 2, "联系电话", bold=True); set_cell(6, 3, "13601650745")
    set_cell(7, 0, "预计到港时间", bold=True);   set_cell(7, 1, eta, span=3)
    set_cell(8, 0, "备注", bold=True);           set_cell(8, 1, remark, span=3)

    doc.add_paragraph()
    add_para("    请贵公司予以安排装货计划。")
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    add_para("上海久联集团有限公司", align="right")
    add_para(date_chinese, align="right")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    doc.save(tmp.name)
    return tmp.name


def _cleanup_temps(file_list):
    for f in file_list:
        try:
            os.remove(f)
        except Exception:
            pass


def _safe_float(val):
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


# ============================================================
#  邮件操作（IMAP + SMTP）
# ============================================================

def connect_imap():
    """连接到 IMAP 收件服务器，返回 imaplib.IMAP4_SSL 对象"""
    import imaplib
    cfg = get_imap_config(ACCOUNT["email"])
    log.info(f"IMAP 连接：{cfg['host']}:{cfg['port']}")
    conn = imaplib.IMAP4_SSL(cfg["host"], cfg["port"])
    conn.login(ACCOUNT["email"], ACCOUNT["password"])
    log.info("IMAP 登录成功")
    return conn


def fetch_new_emails(imap_conn, processed_uids):
    """
    从 INBOX 拉取未处理的新邮件。
    QQ IMAP 的 SUBJECT SEARCH 不准确，直接拉取最近邮件在 Python 侧过滤主题。
    返回 [(msg_uid, email.Message), ...]
    """
    import email as em

    imap_conn.select("INBOX", readonly=False)

    # 拉取最近 100 封邮件（不依赖 IMAP 的 SUBJECT 搜索）
    status, data = imap_conn.uid("SEARCH", None, "ALL")
    if status != "OK" or not data[0]:
        log.info("未找到任何邮件")
        return []

    all_uids = sorted(int(u) for u in data[0].split())
    recent_uids = all_uids[-5:]

    log.info(f"收件箱共 {len(all_uids)} 封，扫描最近 {len(recent_uids)} 封")

    results = []
    for uid in recent_uids:
        if uid in processed_uids:
            continue
        try:
            # 直接获取完整邮件
            status, msg_data = imap_conn.uid("FETCH", str(uid).encode(), "(RFC822)")
            if status != "OK" or not msg_data[0]:
                continue
            raw_email = None
            for item in msg_data:
                if isinstance(item, tuple) and len(item) >= 2:
                    raw_email = item[1]
                    break
            if not raw_email:
                continue

            msg = em.message_from_bytes(raw_email, policy=email.policy.default)

            # 条件1：主题必须为空
            subject = str(msg.get("Subject", "")).strip()
            if subject:
                continue

            # 条件2：必须有 Excel 附件（.xlsx / .xls）
            has_excel = False
            if msg.is_multipart():
                for part in msg.walk():
                    cdisp = str(part.get("Content-Disposition", ""))
                    if "attachment" not in cdisp:
                        continue
                    fname = part.get_filename() or ""
                    if fname.lower().endswith((".xlsx", ".xls")):
                        has_excel = True
                        break

            if not has_excel:
                continue

            log.info(f"  [{uid}] 空主题 + Excel附件")
            results.append((uid, msg))

            if MARK_AS_READ:
                imap_conn.uid("STORE", str(uid).encode(), "+FLAGS", "\\Seen")
        except Exception as e:
            log.warning(f"获取邮件 UID={uid} 失败：{e}")

    log.info(f"扫描完成，匹配 {len(results)} 封")
    return results


def send_reply(original_msg, attachments, body_text=""):
    """
    回复邮件，带附件。
    original_msg: 原始 email.Message（获取发件人、主题）
    attachments: [(filename, filepath), ...]
    """
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    from email.utils import formataddr

    sender = ACCOUNT["email"]
    to_addr = original_msg["From"]
    orig_subject = original_msg.get("Subject", "")

    # 构建回复主题
    if orig_subject.startswith("Re:"):
        reply_subject = orig_subject
    else:
        reply_subject = f"Re: {orig_subject}"

    # 构建邮件
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_addr
    msg["Subject"] = reply_subject
    msg["In-Reply-To"] = original_msg.get("Message-ID", "")
    msg["References"] = original_msg.get("Message-ID", "")

    if not body_text:
        body_text = f"您好，\n\n已根据您的邮件生成 Word 文档，详见附件。\n\n生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n此邮件由系统自动发出。"

    msg.attach(MIMEText(body_text, "plain", "utf-8"))

    for fname, fpath in attachments:
        with open(fpath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{fname}"',
            )
            msg.attach(part)

    # 发送
    smtp_cfg = get_smtp_config(ACCOUNT["email"])
    log.info(f"SMTP 连接：{smtp_cfg['host']}:{smtp_cfg['port']}")
    with smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"]) as smtp:
        smtp.login(ACCOUNT["email"], ACCOUNT["password"])
        smtp.send_message(msg)

    log.info(f"回复已发送 → {to_addr}")


# ============================================================
#  核心处理逻辑
# ============================================================

def process_email(uid, msg):
    """
    处理单封邮件：
      1. 解析正文
      2. 提取表格数据
      3. 生成 Word
      4. 回复附件
    返回 True 表示处理成功。
    """
    sender = msg["From"]
    subject = msg.get("Subject", "")
    log.info(f"处理邮件 UID={uid} | 发件人：{sender} | 主题：{subject}")

    headers, records = None, []
    excel_temps = []

    # 1) 优先尝试 Excel 附件
    excel_files = extract_excel_attachments(msg)
    for xf in excel_files:
        h, r = parse_excel_data(xf)
        if h and r:
            headers = h
            records = r
            excel_temps.append(xf)
            break  # 只处理第一个有效 Excel 附件
        else:
            excel_temps.append(xf)

    # 2) 若没有 Excel 附件，则尝试解析邮件正文 CSV
    if not records:
        body = parse_email_body(msg)
        log.debug(f"邮件正文（前500字）：{body[:500] if body else '(空)'}")
        if body:
            headers, records = parse_csv_from_body(body)

    if not records:
        log.warning(f"邮件 UID={uid} 未识别到数据，跳过")
        send_reply(
            msg, [],
            "您好，\n\n未能从您的邮件中识别到数据。\n\n"
            "支持两种方式：\n"
            "  1. 在邮件中**添加 Excel 附件**（.xlsx），第一行为表头，后续为数据行\n"
            "  2. 在邮件正文中粘贴 CSV 数据，格式如下：\n\n"
            "【出库通知单】\n"
            "目的港,船号,计划量,备注,允许装货时间\n"
            "上海港,江淮油998,476.50,周亚年 13524901555,2026-06-01 08:00\n\n"
            "【提货委托函】\n"
            "公司名,船名,提货计划量,目的港,船舶联系人,联系电话,预计到港时间,备注\n"
            "上海久联集团石油化工有限公司,江淮油998,476.50,上海港,周亚年,13524901555,2026-06-01 08:00,"
        )
        _cleanup_temps(excel_temps)
        return False

    doc_type = detect_doc_type(headers)
    log.info(f"识别文档类型：{doc_type} | {len(records)} 条数据")

    today = datetime.now()
    date_chinese = today.strftime("%Y年%m月%d日")
    date_num = today.strftime("%Y%m%d")

    attachments = []
    temp_files = []

    try:
        if doc_type == "出库通知单":
            tmp = generate_chukutongzhidan(records, date_chinese, date_num)
            attachments.append((f"出库通知单_{date_num}.docx", tmp))
            temp_files.append(tmp)

        elif doc_type == "提货委托函":
            if len(records) == 1:
                tmp = generate_tihuoweituohan(records[0], date_chinese)
                ship = records[0].get("船名", records[0].get("船号", "unknown"))
                fname = f"提货委托函_{ship}.docx"
                attachments.append((fname, tmp))
                temp_files.append(tmp)
            else:
                # 多份文档打包成 ZIP
                zip_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
                zip_tmp.close()
                temp_files.append(zip_tmp.name)
                with zipfile.ZipFile(zip_tmp.name, "w", zipfile.ZIP_DEFLATED) as zf:
                    for i, rec in enumerate(records):
                        tmp = generate_tihuoweituohan(rec, date_chinese)
                        temp_files.append(tmp)
                        ship = _get_field(rec, "船名", "船号") or f"record_{i+1}"
                        zf.write(tmp, f"提货委托函_{ship}.docx")
                attachments.append(("提货委托函_批量.zip", zip_tmp.name))

        send_reply(msg, attachments)
        return True

    except Exception as e:
        log.error(f"处理邮件 UID={uid} 出错：{e}", exc_info=True)
        try:
            send_reply(msg, [], f"您好，\n\n生成 Word 时出错：{e}\n\n请检查数据格式后重试。")
        except Exception:
            pass
        return False

    finally:
        # 清理临时文件
        for tmp in temp_files:
            try:
                os.remove(tmp)
            except Exception:
                pass
        _cleanup_temps(excel_temps)


# ============================================================
#  主循环
# ============================================================

def main_loop():
    """主监控循环"""
    print("=" * 56)
    print("  邮件监控 & 自动生成 Word 工具")
    print("=" * 56)
    print(f"  监控邮箱：{ACCOUNT['email']}")
    print(f"  检查间隔：{CHECK_INTERVAL} 秒")
    print(f"  主题关键词：{', '.join(SUBJECT_KEYWORDS)}")
    print(f"  日志文件：{LOG_FILE}")
    print("=" * 56)
    print("  按 Ctrl+C 停止监控")
    print()

    processed = load_processed_uids(TRACK_FILE)
    log.info(f"已加载 {len(processed)} 条已处理邮件 UID")

    while True:
        try:
            imap = connect_imap()
            try:
                new_emails = fetch_new_emails(imap, processed)
                if new_emails:
                    log.info(f"发现 {len(new_emails)} 封新邮件，开始处理...")
                    for uid, msg in new_emails:
                        ok = process_email(uid, msg)
                        if ok:
                            processed.add(uid)
                        time.sleep(2)  # 处理间隔，避免被限流
                    save_processed_uids(TRACK_FILE, processed)
                else:
                    log.info("无新邮件")
            finally:
                try:
                    imap.logout()
                except Exception:
                    pass

        except Exception as e:
            log.error(f"检查邮件时出错：{e}", exc_info=True)

        time.sleep(CHECK_INTERVAL)


def run_once():
    """仅检查一次并处理，用于测试"""
    print("=" * 56)
    print("  邮件监控 — 单次检查模式")
    print("=" * 56)

    processed = load_processed_uids(TRACK_FILE)
    log.info(f"已加载 {len(processed)} 条已处理邮件 UID")

    imap = connect_imap()
    try:
        new_emails = fetch_new_emails(imap, processed)
        if not new_emails:
            print("无新邮件需要处理。")
            return

        print(f"发现 {len(new_emails)} 封新邮件，开始处理...")
        for uid, msg in new_emails:
            ok = process_email(uid, msg)
            if ok:
                processed.add(uid)
            time.sleep(1)

        save_processed_uids(TRACK_FILE, processed)
        print("处理完成。")
    finally:
        try:
            imap.logout()
        except Exception:
            pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="邮件监控 & 自动生成 Word")
    parser.add_argument(
        "--once", action="store_true",
        help="只检查一次就退出（用于测试）"
    )
    args = parser.parse_args()

    if args.once:
        run_once()
    else:
        main_loop()
