#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成《关于钛材公司控制系统升级改造项目（一期）应急采购说明的函件》正式版 PDF。"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
)

# 注册 Windows 自带中文字体
FONT_DIR = r"C:\Windows\Fonts"
pdfmetrics.registerFont(TTFont("Hei", FONT_DIR + r"\simhei.ttf"))
pdfmetrics.registerFont(TTFont("Fang", FONT_DIR + r"\simfang.ttf"))
pdfmetrics.registerFont(TTFont("Kai", FONT_DIR + r"\simkai.ttf"))

OUT = r"D:\workbuddy\Claw\关于钛材公司控制系统升级改造项目（一期）应急采购说明的函件.pdf"

styles = {
    "title": ParagraphStyle(
        "title", fontName="Hei", fontSize=22, leading=32,
        alignment=TA_CENTER, spaceAfter=6),
    "field": ParagraphStyle(
        "field", fontName="Fang", fontSize=13, leading=22, alignment=TA_LEFT),
    "doc_title": ParagraphStyle(
        "doc_title", fontName="Hei", fontSize=16, leading=28,
        alignment=TA_CENTER, spaceBefore=6, spaceAfter=18),
    "body": ParagraphStyle(
        "body", fontName="Fang", fontSize=16, leading=32,
        alignment=TA_LEFT, firstLineIndent=32, spaceAfter=0),
    "body_noindent": ParagraphStyle(
        "body_noindent", fontName="Fang", fontSize=16, leading=32,
        alignment=TA_LEFT, spaceAfter=0),
    "sign": ParagraphStyle(
        "sign", fontName="Fang", fontSize=16, leading=32, alignment=TA_RIGHT),
}

doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=26 * mm, rightMargin=24 * mm,
    topMargin=30 * mm, bottomMargin=28 * mm,
    title="关于钛材公司控制系统升级改造项目（一期）应急采购说明的函件",
    author="攀钢集团（攀枝花）钛金属材料有限公司设备室",
)

story = []

# 抬头单位名
story.append(Paragraph("攀钢集团（攀枝花）钛金属材料有限公司", styles["title"]))
story.append(Spacer(1, 6))

# 表头信息栏：单位 / 发件人 / 收件人 / 审核人 / 电话 / 签发人
def field(label, value):
    return Paragraph("<b>%s</b>%s" % (label, value), styles["field"])

rows = [
    [field("单　位：", "成都星云智联科技有限公司"), field("发件人：", "何剑")],
    [field("收件人：", "邓人夫"), field("审核人：", "")],
    [field("电　话：", ""), field("签发人：", "")],
]
tbl = Table(rows, colWidths=[100 * mm, 60 * mm])
tbl.setStyle(TableStyle([
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 4),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
]))
story.append(tbl)
story.append(Spacer(1, 4))
story.append(HRFlowable(width="100%", thickness=1.4, color="black"))
story.append(Spacer(1, 16))

# 函件标题
story.append(Paragraph(
    "关于钛材公司控制系统升级改造项目（一期）<br/>应急采购说明的函件",
    styles["doc_title"]))

# 正文
story.append(Paragraph("成都星云智联科技有限公司：", styles["body_noindent"]))
story.append(Spacer(1, 6))
story.append(Paragraph(
    "贵公司承担我公司钛材公司控制系统升级改造项目（一期）（工程编号：J20260018，"
    "合同编号：TCZBCL2026-378）。根据我公司生产计划安排，老线1#氯化炉大修需在"
    "2026年12月10日具备投产条件。为保障1#氯化炉按期投产，请贵公司在"
    "<font name=\"Hei\">2026年11月30日前</font>组织控制系统硬件到货，"
    "<font name=\"Hei\">2026年12月10日前</font>完成1#氯化炉控制系统升级改造。",
    styles["body"]))
story.append(Spacer(1, 40))

# 落款（右侧，预留盖章位置）
story.append(Paragraph(
    "攀钢集团（攀枝花）钛金属材料有限公司设备室", styles["sign"]))
story.append(Paragraph("2026年11月11日", styles["sign"]))
story.append(Spacer(1, 10))
story.append(Paragraph("（加盖公章）", ParagraphStyle(
    "stamp_note", fontName="Kai", fontSize=13, leading=22,
    alignment=TA_RIGHT, textColor="#808080")))

doc.build(story)
print("PDF 已生成：", OUT)
