# -*- coding: utf-8 -*-
"""
卡片批量生成器 v6.1（修复输出文件损坏 + 文字边距 2mm）
======================================
产品名称：卡片批量生成器（Card Batch Generator）
版    本：6.1.0.0
作    者：邱卓尔
版权所有：Copyright (c) 2026 邱卓尔. 保留所有权利.
软件说明：本软件用于批量生成图文卡片 Word 文档，支持多种卡片
          规格（3/4/5/6寸、横竖版、自定义）、图片自动匹配与
          三种适配模式（拉伸填充/铺满裁剪/完整显示）。
免责声明：本软件按"现状"提供，不对因使用本软件产生的任何直接
          或间接损失承担责任。

功能：
  1. 粘贴一段文字，自动按分隔符（，, 、空格 换行 ; ；）拆分词组
  2. 单元格由代码动态生成：需要几张卡片就生成几行
  3. 每张卡片分"图 + 词"两行，整体作为一个卡片，尺寸精确可控
  4. 卡片规格可选：3寸 / 4寸 / 5寸 / 6寸 / 自定义，横版 / 竖版
     （均按国家标准照片尺寸）
     - 3寸 竖版：宽 5.5 × 高 8.4 cm（默认；标准照片 55×84mm）
     - 4寸 竖版：宽 7.6 × 高 10.2 cm（标准照片 76×102mm）
     - 5寸 竖版：宽 8.9 × 高 12.7 cm（标准照片 89×127mm，即 3R）
     - 6寸 竖版：宽 10.2 × 高 15.2 cm（标准照片 102×152mm，即 4R）
  5. 每页行列数根据卡片尺寸自动计算（A4 纸），也可手动指定每页行数
  6. 填入的文字使用内置字体（微软雅黑加粗），并按单元格大小自动收缩字号：
     文字四周各预留 2mm，词组越长字号越小，保证不溢出、不被裁切；
     文字在单元格内水平、垂直居中，与右侧预览完全一致
  7. 图片可自动按词组文件名匹配嵌入，默认"拉伸填充"缩放到单元格大小
     （铺满不留白、不裁切），也可选"铺满裁剪"（不变形、居中裁切）或
     "完整显示"（不变形、可能留白）
  8. 页边距可选（窄 1.27cm / 适中 1.91cm / 常规 2.54cm / 宽 3.18cm /
     自定义），默认"窄"，即上下左右各 1.27cm
  9. 网格对齐可选（默认"满页基准"）：满页卡片在版心内上下左右居中；
     卡片数不满一页时，沿用满页那个左上角基准，网格不会下沉或重新居中
 10. 右侧实时预览：按真实排版比例绘制页面缩略图（纸张、页边距、网格锚
     点、每张卡片的图片与文字），多页可翻页查看，改动参数即时刷新
 11. 输出文件名：默认"卡片【当前日期与时间】"，也可自定义
 12. 紧凑参数栏：整体界面更窄、行距与内边距更小，小屏幕也放得下
 13. 百度搜图补图：预览中单击缺失图片的卡片，弹出搜索窗口按词组
     联网搜索百度图片，首屏一次加载 60 张，可点"加载更多"继续按页
     追加（接口单次上限 60，可一直翻到几百张，自动去重）；
     单击选中、双击确认，图片自动转存为"词组.jpg"到图片文件夹
     并立即刷新预览（生成时即嵌入文档）
 14. 预览中单击已有图片的卡片，直接打开图片所在文件夹并选中该文件；
     右键卡片图片还可选择：在文件夹中显示 / 用默认看图程序打开 /
     更换图片（百度搜图）/ 从本地文件更换；底部"刷新"按钮可在
     文件夹里换图后重新载入缩略图
 15. 版式完全内置：纸张（A4）、表格边框、单元格边距、字体全部由程序
     按规格自动计算生成，**不再需要任何模板文件** —— 打开即用，
     也不存在"模板丢失/路径失效"的问题

核心功能仅使用 Python 标准库；预览窗中的图片缩略图与百度搜图为
可选增强（若环境中有 Pillow 则显示真实缩略图，否则以占位框表示）。
"""

import os
import io
import re
import sys
import json
import queue
import ssl
import struct
import zipfile
import datetime
import subprocess
import threading
import webbrowser
import http.cookiejar
import urllib.parse
import urllib.request
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 可选依赖：Pillow 只用于右侧预览窗里的图片缩略图。
# 未安装时程序一切功能照常，只是预览的图片区显示为占位框。
try:
    from PIL import Image, ImageTk, ImageOps
    _HAS_PIL = True
except Exception:
    _HAS_PIL = False

# Pillow ≥10 把重采样常量移到了 Image.Resampling 命名空间
_RESAMPLE = (getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1)
             if _HAS_PIL else 1)

# --------------------------- 版本与版权信息 ---------------------------
APP_NAME = "卡片批量生成器"
APP_NAME_EN = "Card Batch Generator"
APP_VERSION = "6.1.0.0"
APP_AUTHOR = "邱卓尔"
APP_COMPANY = "邱卓尔"
APP_COPYRIGHT = "Copyright (c) 2026 邱卓尔. 保留所有权利."
APP_DESCRIPTION = ("批量生成图文卡片 Word 文档：版式随程序内置（界面无需选择模板），"
                   "支持 3/4/5/6 寸与自定义规格、"
                   "横竖版、图片自动匹配与拉伸填充/铺满裁剪/完整显示三种适配模式；"
                   "页边距可选（默认“窄”1.27cm），网格对齐可选（默认满页基准），"
                   "文字按单元格自动缩放（四周留 2mm、加粗居中），"
                   "内置右侧文档实时预览窗（与输出排版一致，可翻页）；"
                   "预览缺图可一键百度搜图补齐，点击已有图片可在文件夹中定位。")

# ----------------------------- 配置 -----------------------------

# 打包成 EXE（PyInstaller onedir）时，配置文件保存在 EXE 所在目录，
# 而不是脚本/临时解压目录，否则配置会丢失。
if getattr(sys, "frozen", False):
    _APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_APP_DIR, "卡片生成器配置.json")

# ---------------- 内置卡片版式（不再需要外部模板文件） ----------------
# 历史版本从"卡片【模板】.docx"中提取表格骨架与文字行原型。经逐项核对，模板
# 提供的内容全部是固定值 —— A4 纸张、0.5 磅单线表格边框、单元格左右内边距
# 108 缇、文字行（微软雅黑 36 磅加粗、水平居中）；而卡片的尺寸、行列数、
# 页边距、字号早由程序按规格自动计算。故把这些"版式底稿"内置为常量，
# 程序不再依赖任何模板文件，也能生成排版完全一致的文档。

# 纸张固定为 A4（尺寸常量见下方"页面设置"一节），页边距由界面设置，
# 卡片尺寸/行列/字号全部由程序自动计算。

BUILTIN_DOC_PROLOGUE = (
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document xmlns:wpc=\"http://s"
    "chemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.mic"
    "rosoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/d"
    "rawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/"
    "21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xml"
    "ns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http:/"
    "/schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.micro"
    "soft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/offic"
    "e/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016"
    "/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xm"
    "lns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.o"
    "rg/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeD"
    "ocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.mi"
    "crosoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlform"
    "ats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w=\"http://schemas.openxmlformats.org"
    "/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wo"
    "rdml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w15=\"http://schemas.micros"
    "oft.com/office/word/2012/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/201"
    "0/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordproc"
    "essingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"h"
    "ttp://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:wpsCustomData=\"h"
    "ttp://www.wps.cn/officeDocument/2013/wpsCustomData\" mc:Ignorable=\"w14 w15 wp14\"><w:body>"
)

BUILTIN_TBL_HEAD = (
    "<w:tbl><w:tblPr><w:tblStyle w:val=\"3\"/><w:tblW w:w=\"11040\" w:type=\"dxa\"/><w:tblInd w:w=\""
    "0\" w:type=\"dxa\"/><w:tblBorders><w:top w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\""
    "/><w:left w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:bottom w:val=\"single\" w"
    ":color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:right w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:sp"
    "ace=\"0\"/><w:insideH w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideV w:val"
    "=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/></w:tblBorders><w:tblLayout w:type=\"fixed"
    "\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom "
    "w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr><w:tblGr"
    "id><w:gridCol w:w=\"3680\"/><w:gridCol w:w=\"3680\"/><w:gridCol w:w=\"3680\"/></w:tblGrid>"
)

BUILTIN_CARD_ROW = (
    "<w:tr><w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\""
    "/><w:left w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:bottom w:val=\"single\" w"
    ":color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:right w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:sp"
    "ace=\"0\"/><w:insideH w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideV w:val"
    "=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/></w:tblBorders><w:tblCellMar><w:top w:w=\""
    "0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:righ"
    "t w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPrEx><w:trPr><w:trHeight w:val=\"1417\" w:"
    "hRule=\"atLeast\"/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"3680\" w:type=\"dxa\"/><w:vAlign w:val="
    "\"top\"/></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:hint=\"eastAsia\" w:e"
    "astAsiaTheme=\"minorEastAsia\"/><w:vertAlign w:val=\"baseline\"/><w:lang w:val=\"en-US\" w:eas"
    "tAsia=\"zh-CN\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:ascii=\"微软雅黑\" w:"
    "hAnsi=\"微软雅黑\" w:eastAsia=\"微软雅黑\" w:cs=\"微软雅黑\"/><w:b/><w:bCs/><w:sz w:val=\"72\"/><w:szCs w:va"
    "l=\"72\"/><w:vertAlign w:val=\"baseline\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\"/></w:rPr"
    "><w:t>文字</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3680\" w:type=\"dxa\"/><w:vAlign"
    " w:val=\"top\"/></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:vertAlign w:val=\"base"
    "line\"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:ascii=\"微软雅黑\" w:hAnsi=\"微"
    "软雅黑\" w:eastAsia=\"微软雅黑\" w:cs=\"微软雅黑\"/><w:b/><w:bCs/><w:sz w:val=\"72\"/><w:szCs w:val=\"72\"/>"
    "<w:vertAlign w:val=\"baseline\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\"/></w:rPr><w:t>文字"
    "</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3680\" w:type=\"dxa\"/><w:vAlign w:val=\""
    "top\"/></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:vertAlign w:val=\"baseline\"/><"
    "/w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint=\"eastAsia\" w:ascii=\"微软雅黑\" w:hAnsi=\"微软雅黑\" w:e"
    "astAsia=\"微软雅黑\" w:cs=\"微软雅黑\"/><w:b/><w:bCs/><w:sz w:val=\"72\"/><w:szCs w:val=\"72\"/><w:vertA"
    "lign w:val=\"baseline\"/><w:lang w:val=\"en-US\" w:eastAsia=\"zh-CN\"/></w:rPr><w:t>文字</w:t></"
    "w:r></w:p></w:tc></w:tr>"
)

BUILTIN_SPACER_ROW = (
    "<w:tr><w:tblPrEx><w:tblBorders><w:top w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\""
    "/><w:left w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:bottom w:val=\"single\" w"
    ":color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:right w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:sp"
    "ace=\"0\"/><w:insideH w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideV w:val"
    "=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/></w:tblBorders><w:tblCellMar><w:top w:w=\""
    "0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:righ"
    "t w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPrEx><w:trPr><w:trHeight w:val=\"3685\" w:"
    "hRule=\"atLeast\"/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"3680\" w:type=\"dxa\"/><w:vAlign w:val="
    "\"top\"/></w:tcPr><w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:rFonts w:hint=\"eastAsia\" w:e"
    "astAsiaTheme=\"minorEastAsia\"/><w:vertAlign w:val=\"baseline\"/><w:lang w:eastAsia=\"zh-CN\"/"
    "></w:rPr></w:pPr><w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"0\"/></"
    "w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3680\" w:type=\"dxa\"/><w:vAlign w:val=\"top\"/></w:tcPr"
    "><w:p><w:pPr><w:jc w:val=\"both\"/><w:rPr><w:vertAlign w:val=\"baseline\"/></w:rPr></w:pPr><"
    "/w:p></w:tc><w:tc><w:tcPr><w:tcW w:w=\"3680\" w:type=\"dxa\"/><w:vAlign w:val=\"top\"/></w:tcP"
    "r><w:p><w:pPr><w:jc w:val=\"both\"/><w:rPr><w:vertAlign w:val=\"baseline\"/></w:rPr></w:pPr>"
    "</w:p></w:tc></w:tr>"
)

BUILTIN_SECTPR = (
    "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\""
    "550\" w:left=\"380\" w:header=\"851\" w:footer=\"992\" w:gutter=\"0\"/><w:paperSrc/><w:cols w:spa"
    "ce=\"0\" w:num=\"1\"/><w:rtlGutter w:val=\"0\"/><w:docGrid w:type=\"lines\" w:linePitch=\"312\" w:"
    "charSpace=\"0\"/></w:sectPr>"
)

BUILTIN_STYLES_XML = (
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"http://schemas"
    ".openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.o"
    "rg/officeDocument/2006/relationships\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:as"
    "cii=\"微软雅黑\" w:hAnsi=\"微软雅黑\" w:eastAsia=\"微软雅黑\" w:cs=\"微软雅黑\"/><w:sz w:val=\"21\"/><w:szCs w:val"
    "=\"24\"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type=\"paragraph\""
    " w:default=\"1\" w:styleId=\"a\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style><w:style w:ty"
    "pe=\"character\" w:default=\"1\" w:styleId=\"a0\"><w:name w:val=\"Default Paragraph Font\"/><w:u"
    "iPriority w:val=\"1\"/><w:semiHidden/><w:unhideWhenUsed/></w:style><w:style w:type=\"table\""
    " w:default=\"1\" w:styleId=\"a1\"><w:name w:val=\"Normal Table\"/><w:uiPriority w:val=\"99\"/><w"
    ":semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><"
    "w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dx"
    "a\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="
    "\"table\" w:styleId=\"3\"><w:name w:val=\"Table Grid\"/><w:basedOn w:val=\"a1\"/><w:uiPriority w"
    ":val=\"0\"/><w:pPr><w:widowControl w:val=\"0\"/><w:jc w:val=\"both\"/></w:pPr><w:tblPr><w:tblB"
    "orders><w:top w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:left w:val=\"single\""
    " w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:bottom w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w"
    ":space=\"0\"/><w:right w:val=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideH w:va"
    "l=\"single\" w:color=\"auto\" w:sz=\"4\" w:space=\"0\"/><w:insideV w:val=\"single\" w:color=\"auto\""
    " w:sz=\"4\" w:space=\"0\"/></w:tblBorders></w:tblPr></w:style></w:styles>"
)

BUILTIN_SETTINGS_XML = (
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:settings xmlns:w=\"http://schem"
    "as.openxmlformats.org/wordprocessingml/2006/main\" xmlns:m=\"http://schemas.openxmlformats"
    ".org/officeDocument/2006/math\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r"
    "=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:zoom w:percent"
    "=\"100\"/><w:defaultTabStop w:val=\"420\"/><w:characterSpacingControl w:val=\"compressPunctua"
    "tion\"/><w:compat/><w:themeFontLang w:val=\"en-US\" w:eastAsia=\"zh-CN\"/></w:settings>"
)


# 卡片尺寸表（竖版：宽 cm, 高 cm）—— 按国家标准照片尺寸（1英寸=2.54cm）
#   3寸：55×84mm   4寸：76×102mm   5寸：89×127mm（3R）   6寸：102×152mm（4R）
# 横版自动宽高对调；三寸及以上均可放入对应塑封膜
CARD_SIZES = {
    "3寸": (5.5, 8.4),
    "4寸": (7.6, 10.2),
    "5寸": (8.9, 12.7),
    "6寸": (10.2, 15.2),
}
SIZE_ORDER = ["3寸", "4寸", "5寸", "6寸", "自定义"]
POUCH_INFO = {
    "3寸": "标准照片 55×84mm，适配 65×95mm 塑封膜",
    "4寸": "标准照片 76×102mm，适配 85×110mm 塑封膜",
    "5寸": "标准照片 89×127mm（3R），适配 95×135mm 塑封膜",
    "6寸": "标准照片 102×152mm（4R），适配 110×160mm 塑封膜",
    "自定义": "",
}

# 文字行高度策略：占卡片高度的 28%，限制在 1050～1417 缇之间
# （1417 为内置原型文字行高；1050 缇 ≈ 1.85cm）
WORD_ROW_MIN = 1050
WORD_ROW_MAX = 1417

# --------------------- 文字自适应（按单元格收缩字号） ---------------------
# 文字四周各预留 2mm，字号按单元格可用空间自动收缩，保证不溢出、不被裁切
TEXT_PAD_MM = 2.0
# 表格单元格左右默认内边距（缇）= 108，与模板 <w:tblCellMar> 一致
CELL_MAR_TW = 108
# 单倍行高系数：文字实际占位高度 ≈ 字号 × 1.38
LINE_FACTOR = 1.38
# 汉字/全角字符宽度 = 1 em；半角字符按经验宽度折算（微软雅黑）
CHAR_EM_HALF = 0.56
# 字号下限（磅），避免极长文字缩到看不清
TEXT_PT_MIN = 8.0
# 文字字号兜底上限（磅）= 内置文字行原型的 36 磅
DEFAULT_TEXT_PT = 36.0

# ----------------------- 页面设置（A4 + 页边距） -----------------------
# A4 纸张尺寸（缇）：210 × 297 mm
A4_W_TW = 11906
A4_H_TW = 16838

# 页边距预设（对应 Word「页面设置 → 页边距」内置预设）
#   窄 = 上下左右各 1.27cm（0.5 英寸 = 720 缇）—— 本软件默认
MARGIN_PRESETS = {
    "窄（1.27cm）": 1.27,
    "适中（1.91cm）": 1.91,
    "常规（2.54cm）": 2.54,
    "宽（3.18cm）": 3.18,
}
DEFAULT_MARGIN_KEY = "窄（1.27cm）"
MARGIN_ORDER = ["窄（1.27cm）", "适中（1.91cm）", "常规（2.54cm）", "宽（3.18cm）", "自定义"]
DEFAULT_MARGIN_CM = MARGIN_PRESETS[DEFAULT_MARGIN_KEY]

# 网格对齐方式：决定卡片表格在纸张上的锚点
#   anchor  —— 满页基准（默认）：以"满页卡片"的位置为基准做双向居中，
#              所有页共用同一个左上角基准 —— 满页时上下左右居中；
#              不满页时，网格左上角仍落在满页时那个左上角，不重新居中、不下沉
#   topleft —— 贴版心左上角（完全贴齐页边距）
#   center  —— 每页各自居中（内容不满页时整块居中）
ALIGN_MODES = {
    "满页基准（推荐）": "anchor",
    "贴版心左上角": "topleft",
    "每页各自居中": "center",
}
ALIGN_ORDER = list(ALIGN_MODES.keys())
DEFAULT_ALIGN_KEY = "满页基准（推荐）"
DEFAULT_ALIGN_MODE = ALIGN_MODES[DEFAULT_ALIGN_KEY]

# 极小占位段落（字号 1pt、行距固定 20 缇），用于承载分页符与节属性，
# 使其几乎不占用垂直空间，从而保证网格在页面内垂直居中精确
_TINY_RPR = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
_TINY_PPR = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
             'w:lineRule="exact"/>' + _TINY_RPR + '</w:pPr>')
TINY_P = '<w:p>' + _TINY_PPR + '</w:p>'
PAGE_BREAK_P = '<w:p>' + _TINY_PPR + '<w:r><w:br w:type="page"/></w:r></w:p>'


def _top_spacer(height_tw: int) -> str:
    """生成一个"高度精确"的空白段落，用于把每页网格统一下推到基准位置。
    必须带 snapToGrid=0，否则会被文档网格（linePitch=312）吸附而改变高度。"""
    return ('<w:p><w:pPr><w:snapToGrid w:val="0"/>'
            '<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/>'
            '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>'
            % max(0, int(height_tw)))

# 卡片之间的分隔空行高度（缇），约 0.7cm，用于裁剪
SPACER_HEIGHT = 400

# 分隔符：中英文逗号、顿号、分号、斜杠、空白符（空格/制表/换行）
DELIMITER_RE = re.compile(r"[，,、;；/\s]+")

# 支持的图片扩展名
IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.wmf', '.emf')

EXT_CONTENT_TYPE = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'gif': 'image/gif', 'bmp': 'image/bmp', 'tif': 'image/tiff',
    'tiff': 'image/tiff', 'wmf': 'image/x-wmf', 'emf': 'image/x-emf',
}


# --------------------- 百度图片搜索（v5.0 联网补图） ---------------------
# 通过百度图片公开搜索接口获取结果：
#   1) 首次请求先访问百度图片首页换取 Cookie（否则会被反爬拒绝）
#   2) acjson JSON 接口取缩略图/大图链接，失败时从搜索页 HTML 兜底抽取

_BD_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
_BD_OPENER = None


def _ssl_ctx():
    try:
        return ssl.create_default_context()
    except Exception:
        return ssl._create_unverified_context()


def _bd_opener():
    """带 Cookie 的 urllib opener（懒加载，仅初始化一次）"""
    global _BD_OPENER
    if _BD_OPENER is None:
        cj = http.cookiejar.CookieJar()
        _BD_OPENER = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(cj),
            urllib.request.HTTPSHandler(context=_ssl_ctx()))
        _BD_OPENER.addheaders = [
            ("User-Agent", _BD_UA),
            ("Accept-Language", "zh-CN,zh;q=0.9"),
        ]
        try:  # 先访问首页换取 Cookie，避免"Forbid spider access"
            _BD_OPENER.open("https://image.baidu.com/", timeout=10).read(65536)
        except Exception:
            pass
    return _BD_OPENER


def baidu_image_search(word, pn=0, rn=60):
    """百度图片搜索，返回 [{"cands":[大图候选URL...], "thumb":缩略图URL,
    "w":宽, "h":高}, ...]；网络异常时抛出异常。

    pn 为起始序号（翻页用），rn 为单次条数（接口单次最多返回 60 条）。"""
    if not word:
        return []
    q = urllib.parse.quote(word)
    url = ("https://image.baidu.com/search/acjson?tn=resultjson_com&ipn=rj"
           "&word=%s&queryWord=%s&pn=%d&rn=%d&ie=utf-8&oe=utf-8"
           "&ct=201326592&lm=-1&cl=2&fp=result" % (q, q, int(pn), int(rn)))
    raw = _bd_opener().open(url, timeout=15).read().decode("utf-8", "replace")
    if "antiFlag" in raw:
        raise RuntimeError("百度拒绝了本次访问，请稍后重试。")
    items = []
    s, e = raw.find("{"), raw.rfind("}")
    if s >= 0 and e > s:
        try:
            for it in (json.loads(raw[s:e + 1]).get("data") or []):
                thumb = it.get("thumbURL") or it.get("hoverURL") or ""
                if not thumb:
                    continue
                cands = []
                for k in ("objURL", "middleURL", "hoverURL", "thumbURL"):
                    u = it.get(k)
                    if u and u not in cands:
                        cands.append(u)
                items.append({"cands": cands, "thumb": thumb,
                              "w": it.get("width") or 0,
                              "h": it.get("height") or 0})
        except ValueError:
            pass
    if not items:  # 兜底：从搜索结果页 HTML 抽取
        url2 = ("https://image.baidu.com/search/index?tn=baiduimage"
                "&ie=utf-8&word=%s" % q)
        raw2 = _bd_opener().open(url2, timeout=15).read().decode("utf-8", "replace")
        thumbs = re.findall(r'"thumbURL":"(https?://[^"]+?)"', raw2)
        bigs = re.findall(r'"middleURL":"(https?://[^"]+?)"', raw2)
        for i, t in enumerate(thumbs):
            b = bigs[i] if i < len(bigs) else t
            items.append({"cands": [b, t], "thumb": t, "w": 0, "h": 0})
    return items


def fetch_image_bytes(cands, min_px=0):
    """依次尝试候选 URL 下载图片，返回 (bytes, PIL.Image)；失败抛异常。
    min_px > 0 时跳过分辨率过低的图（用于挑选高质量大图）。"""
    last_err = None
    for u in (cands or [])[:4]:
        try:
            req = urllib.request.Request(u, headers={
                "User-Agent": _BD_UA, "Referer": "https://image.baidu.com/"})
            data = _bd_opener().open(req, timeout=20).read()
            if not data:
                continue
            if _HAS_PIL:
                im = Image.open(io.BytesIO(data))
                im.load()
                if min_px and im.width < min_px and im.height < min_px:
                    last_err = RuntimeError("图片分辨率过低（%d×%d）"
                                            % (im.width, im.height))
                    continue
                return data, im
            return data, None
        except Exception as ex:
            last_err = ex
    raise last_err or RuntimeError("图片下载失败")


def save_word_image(word, image_dir, im):
    """把图片统一转存为 <词组>.jpg（文件名与词组一致，Word 兼容性最好）"""
    if im is None:
        raise RuntimeError("没有可保存的图片")
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    path = os.path.join(image_dir, word + ".jpg")
    im.save(path, "JPEG", quality=92)
    return path


def reveal_in_folder(path):
    """在系统文件管理器中打开所在文件夹并选中该文件"""
    path = os.path.abspath(path)
    try:
        if sys.platform.startswith("win"):
            # explorer 的 /select, 只接受命令行字符串形式，路径需带引号
            subprocess.Popen('explorer /select,"%s"'
                             % path.replace("/", "\\"), shell=False)
            return True
        if sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
            return True
        subprocess.Popen(["xdg-open", os.path.dirname(path)])
        return True
    except Exception:
        return False


def open_image_file(path):
    """用系统默认看图程序打开单张图片"""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)                       # noqa: S606 (Windows API)
            return True
        webbrowser.open("file://" + os.path.abspath(path))
        return True
    except Exception:
        return False


def xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def find_image_for(word: str, folder: str):
    """在图片文件夹中按"词组名+扩展名"查找图片，返回完整路径或 None"""
    if not folder or not word:
        return None
    for ext in IMAGE_EXTS:
        cand = os.path.join(folder, word + ext)
        if os.path.isfile(cand):
            return cand
    return None


def parse_words(text: str):
    """按分隔符拆分词组"""
    return [w.strip() for w in DELIMITER_RE.split(text or "") if w.strip()]


def cm_to_twips(cm: float) -> int:
    return int(round(cm / 2.54 * 1440))


def twips_to_cm(tw: int) -> float:
    return round(tw / 1440 * 2.54, 2)


# ------------------------- 图片尺寸读取 -------------------------

def _read_be32(data, off):
    return struct.unpack('>I', data[off:off + 4])[0]


def _read_le16(data, off):
    return struct.unpack('<H', data[off:off + 2])[0]


def _read_le32(data, off):
    return struct.unpack('<i', data[off:off + 4])[0]


def _get_image_size(path: str):
    try:
        with open(path, 'rb') as f:
            head = f.read(32)
        ext = os.path.splitext(path)[1].lower()
        if ext == '.png' and head[:8] == b'\x89PNG\r\n\x1a\n':
            return _read_be32(head, 16), _read_be32(head, 20)
        if ext == '.gif' and head[:6] in (b'GIF87a', b'GIF89a'):
            return _read_le16(head, 6), _read_le16(head, 8)
        if ext == '.bmp' and head[:2] == b'BM':
            return _read_le32(head, 18), abs(_read_le32(head, 22))
        if ext in ('.jpg', '.jpeg'):
            return _jpg_size(path)
    except Exception:
        pass
    return None, None


def _jpg_size(path):
    with open(path, 'rb') as f:
        if f.read(2) != b'\xff\xd8':
            return None, None
        while True:
            b = f.read(1)
            if not b:
                break
            if b != b'\xff':
                continue
            marker = f.read(1)
            if not marker:
                break
            if marker in (b'\xd9', b'\xda'):
                break
            if marker in (b'\xd8', b'\xd0', b'\xd1', b'\xd2', b'\xd3',
                          b'\xd4', b'\xd5', b'\xd6', b'\xd7'):
                continue
            size = struct.unpack('>H', f.read(2))[0]
            if size < 2:
                break
            if marker in (b'\xc0', b'\xc1', b'\xc2'):
                f.read(1)  # 精度(1字节)，此后紧跟 高(2字节)、宽(2字节)
                h = struct.unpack('>H', f.read(2))[0]
                w = struct.unpack('>H', f.read(2))[0]
                return w, h
            f.seek(size - 2, 1)
    return None, None


# ------------------------- 内置版式底稿 -------------------------

def _explicit_fonts(xml: str) -> str:
    """把 <w:rFonts> 里的主题字体引用换成显式"微软雅黑"。

    内置版式不再携带主题部件（theme1.xml），若保留 w:eastAsiaTheme 之类的
    主题引用，不同阅读器可能回退到别的字体；统一成显式字体后，任何环境下
    的字体、字号都完全确定（正文字符的字体原本就是显式指定的，不受影响）。
    """
    def repl(m):
        attrs = re.sub(r'\s+w:(?:ascii|hAnsi|eastAsia|cs)Theme="[^"]*"', '',
                       m.group(1))
        if 'w:ascii=' not in attrs:
            attrs += ' w:ascii="微软雅黑" w:hAnsi="微软雅黑"'
        if 'w:eastAsia=' not in attrs:
            attrs += ' w:eastAsia="微软雅黑"'
        if 'w:cs=' not in attrs:
            attrs += ' w:cs="微软雅黑"'
        return '<w:rFonts%s/>' % attrs
    return re.sub(r'<w:rFonts([^>]*?)/>', repl, xml)


def builtin_prototypes():
    """返回内置版式底稿（不再读取任何模板文件）。

    返回 (tbl_head, card_row, spacer_row, page_w, page_h)：
      - tbl_head  : 表格属性 + 列宽网格（列宽/列数由 _rebuild_tbl_head 重算）
      - card_row  : 文字行原型（微软雅黑加粗、水平居中、垂直居中由生成时统一设置）
      - spacer_row: 分隔空行原型（仅取其 <w:tr> 起始标签）
      - page_w/h  : 纸张尺寸（缇，A4）
    """
    return (_explicit_fonts(BUILTIN_TBL_HEAD), _explicit_fonts(BUILTIN_CARD_ROW),
            _explicit_fonts(BUILTIN_SPACER_ROW), A4_W_TW, A4_H_TW)


def content_size(page_w: int, page_h: int, margin_cm: float):
    """页面可用内容区（缇）= 纸张尺寸 - 上下左右页边距"""
    m = cm_to_twips(margin_cm)
    return max(1, page_w - 2 * m), max(1, page_h - 2 * m)


def template_page_size(template_path=None):
    """纸张尺寸（缇）：固定 A4，不再依赖模板文件（保留函数名兼容旧调用）"""
    return A4_W_TW, A4_H_TW


def _apply_page_setup(tail_xml: str, margin_tw: int, v_center: bool = False) -> str:
    """统一页边距为四边 margin_tw 缇；v_center 为真时页面内容垂直居中"""
    tail_xml = re.sub(
        r'<w:pgMar\s[^>]*/>',
        '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d"'
        ' w:header="851" w:footer="992" w:gutter="0"/>'
        % (margin_tw, margin_tw, margin_tw, margin_tw),
        tail_xml)
    # 内置 sectPr 中可能带有 vAlign，先按需清掉，保证"顶端对齐"模式真正顶端对齐
    tail_xml = re.sub(r'<w:vAlign\b[^>]*/>', '', tail_xml)
    if v_center:
        # <w:vAlign> 必须位于 <w:cols> 之后（CT_SectPr 元素顺序）；
        # <w:cols .../> 为自闭合标签，故按其结束位置插入
        m = re.search(r'<w:cols\b[^>]*/>', tail_xml)
        if m:
            pos = m.end()
        else:
            m2 = re.search(r'<w:pgMar\b[^>]*/>', tail_xml)
            pos = m2.end() if m2 else tail_xml.find('</w:sectPr>')
        if pos and pos > 0:
            tail_xml = tail_xml[:pos] + '<w:vAlign w:val="center"/>' + tail_xml[pos:]
    return tail_xml


# --------------------- 文字宽度估算与字号自适应 ---------------------

def _char_em(ch: str) -> float:
    """估算单个字符宽度（相对字号的 em 倍数）"""
    o = ord(ch)
    # 汉字、假名、韩文、全角标点等全角字符：宽度恰为 1 em
    if (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF or 0xAC00 <= o <= 0xD7A3
            or 0xF900 <= o <= 0xFAFF or 0xFE10 <= o <= 0xFE6F
            or 0xFF00 <= o <= 0xFF60 or 0xFFE0 <= o <= 0xFFE6):
        return 1.0
    if ch in " .,:;'!|iljI`()[]":
        return 0.32
    if ch.isdigit():
        return 0.58
    if 'A' <= ch <= 'Z':
        return 0.70
    return CHAR_EM_HALF


def text_em_width(s: str) -> float:
    """整串文字的宽度（单位：em，即字号倍数）"""
    return sum(_char_em(c) for c in s or "")


def compute_word_pt(word, col_tw, row_h_tw, max_pt=DEFAULT_TEXT_PT,
                    pad_mm=TEXT_PAD_MM):
    """按单元格可用空间计算文字字号（磅）。

    可用区域 = 单元格宽/高 − 单元格内边距 − 四周各 pad_mm；
    分别校验"折成 1..4 行"时宽度与高度允许的最大字号，取其中最优值，
    并封顶到 max_pt（只缩不放），保证文字既不溢出也不被裁切。
    """
    if not word:
        return max_pt
    avail_w = twips_to_cm(max(1, col_tw - 2 * CELL_MAR_TW)) - 2 * pad_mm / 10.0
    avail_w *= 0.99        # 1% 安全余量：避免临界贴合在取整/渲染时意外换行
    avail_h = twips_to_cm(max(1, row_h_tw)) - 2 * pad_mm / 10.0
    if avail_w <= 0.05 or avail_h <= 0.05:
        return min(TEXT_PT_MIN, max_pt)
    total_em = max(0.1, text_em_width(word))
    best_cm = 0.0
    for lines in range(1, 5):
        em_w = avail_w / (total_em / lines)                # 宽度约束
        em_h = avail_h / (lines * LINE_FACTOR)             # 高度约束
        best_cm = max(best_cm, min(em_w, em_h))
    pt = min(max_pt, best_cm / 2.54 * 72.0)
    pt = max(min(TEXT_PT_MIN, max_pt), pt)
    return int(pt * 2) / 2.0        # 取 0.5 磅步长（向下取整，防溢出）


def proto_word_pt(card_row: str) -> float:
    """文字行原型的原始字号（磅），作为自动收缩的上限"""
    szs = [int(x) for x in re.findall(r'<w:sz w:val="(\d+)"', card_row or "")]
    return (max(szs) / 2.0) if szs else DEFAULT_TEXT_PT


def _style_word_run(cell_xml: str, half_pt: int) -> str:
    """统一单元格内文字的字体尺寸（半磅）并确保加粗，与预览保持一致"""
    sz_tags = '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (half_pt, half_pt)
    if '<w:sz ' in cell_xml or '<w:szCs ' in cell_xml:
        cell_xml = re.sub(r'<w:sz w:val="\d+"/>',
                          '<w:sz w:val="%d"/>' % half_pt, cell_xml)
        cell_xml = re.sub(r'<w:szCs w:val="\d+"/>',
                          '<w:szCs w:val="%d"/>' % half_pt, cell_xml)
    else:
        cell_xml = cell_xml.replace('<w:rPr>', '<w:rPr>' + sz_tags)
        cell_xml = cell_xml.replace('<w:rPr/>', '<w:rPr>' + sz_tags + '</w:rPr>')
    if '<w:b/>' not in cell_xml:
        cell_xml = cell_xml.replace('<w:rPr>', '<w:rPr><w:b/><w:bCs/>')
        cell_xml = cell_xml.replace('<w:rPr/>', '<w:rPr><w:b/><w:bCs/></w:rPr>')
    return cell_xml


# ------------------------- 行/单元格生成 -------------------------

def _snap_off(xml: str) -> str:
    """给所有段落加 snapToGrid=0，防止文档网格把行高撑大，保证卡片尺寸精确"""
    xml = xml.replace('<w:pPr>', '<w:pPr><w:snapToGrid w:val="0"/>')
    return xml.replace('<w:pPr/>', '<w:pPr><w:snapToGrid w:val="0"/></w:pPr>')


def _tr_open(row_xml: str) -> str:
    m = re.match(r'(<w:tr(?:\s+[^>]*)>)', row_xml)
    if m:
        return re.sub(r'\s+w14:paraId="[0-9A-Fa-f]+"', "", m.group(1))
    return '<w:tr>'


def _tr_pr(height: int) -> str:
    return '<w:trPr><w:trHeight w:val="%d" w:hRule="exact"/></w:trPr>' % height


def _cell_width_fix(xml: str, col_twips: int) -> str:
    """把行内所有单元格宽度统一设置为列宽"""
    return re.sub(r'(<w:tcW w:w=")\d+(")', r'\g<1>%d\2' % col_twips, xml)


def _make_picture_cell(word, image_folder, img_index, rels, col_twips, max_emu,
                       fit_mode="stretch"):
    """生成单个图片单元格：匹配到图片则嵌入，否则显示占位符"""
    tc_pr = ('<w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="center"/>'
             '<w:shd w:val="clear" w:fill="FFFFFF"/></w:tcPr>' % col_twips)

    image_path = find_image_for(word, image_folder)

    if not image_path:
        return _snap_off(
            tc_pr +
            '<w:p><w:pPr><w:jc w:val="center"/>'
            '<w:rPr><w:color w:val="999999"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:pPr>'
            '<w:r><w:rPr><w:color w:val="999999"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr>'
            '<w:t>【图片】</w:t></w:r></w:p>')

    px_w, px_h = _get_image_size(image_path)
    if fit_mode == "stretch":
        # 拉伸填充：宽高分别缩放到单元格大小，完全铺满、不裁切、不留白
        emu_w, emu_h = max_emu
        src_rect = ""
    elif fit_mode == "cover":
        emu_w, emu_h, src_rect = _fit_cover(px_w, px_h, max_emu)
    else:
        emu_w, emu_h = _px_to_emu(px_w, px_h, max_emu)
        src_rect = ""

    idx = img_index[0]
    img_index[0] += 1
    rel_id = "rIdImg%d" % idx
    media_name = "image%d%s" % (idx, os.path.splitext(image_path)[1].lower())
    rels.append((rel_id, media_name, image_path))

    drawing = (
        '<w:drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="%s"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic>'
        '<pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/>%s<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
        % (emu_w, emu_h, idx, media_name, idx, media_name, rel_id,
           src_rect, emu_w, emu_h))

    return _snap_off(
        tc_pr +
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>%s</w:r></w:p>' % drawing)


def _px_to_emu(px_w, px_h, max_box):
    """像素 → EMU（96dpi），等比缩放不超过 max_box=(可用宽EMU, 可用高EMU)
    完整显示模式：长边贴合、不变形、不越界"""
    max_w, max_h = max_box
    if not px_w or not px_h:
        return max_w, max_h
    emu_w = px_w * 9525
    emu_h = px_h * 9525
    scale = min(max_w / emu_w, max_h / emu_h)
    return int(emu_w * scale), int(emu_h * scale)


def _fit_cover(px_w, px_h, box):
    """铺满裁剪模式：图片等比放大填满整个 box=(宽EMU, 高EMU)，超出部分居中裁剪。
    返回 (显示宽EMU, 显示高EMU, srcRect XML)；显示比例恒等于 box 比例，绝不拉伸。
    srcRect 数值为千分之百分比（100000 = 100%）。"""
    box_w, box_h = box
    if not px_w or not px_h:
        return box_w, box_h, ""
    ratio_img = px_w / px_h
    ratio_box = box_w / box_h
    if abs(ratio_img - ratio_box) / max(ratio_box, 1e-9) < 0.005:
        return box_w, box_h, ""  # 比例接近，无需裁剪
    if ratio_img > ratio_box:
        # 图片更宽：左右各裁一部分，可见宽度占比 = box比/图比
        f = ratio_box / ratio_img
        crop = int(round((1 - f) / 2 * 100000))
        return box_w, box_h, '<a:srcRect l="%d" r="%d"/>' % (crop, crop)
    # 图片更高：上下各裁一部分
    f = ratio_img / ratio_box
    crop = int(round((1 - f) / 2 * 100000))
    return box_w, box_h, '<a:srcRect t="%d" b="%d"/>' % (crop, crop)


def _make_picture_row(prototype_row, block_words, image_folder, img_index, rels,
                      col_twips, height, max_emu, cols, fit_mode="stretch"):
    cells = []
    for c in range(cols):
        w = block_words[c] if c < len(block_words) else ""
        cells.append('<w:tc>' + _make_picture_cell(
            w, image_folder, img_index, rels, col_twips, max_emu, fit_mode) + '</w:tc>')
    return (_tr_open(prototype_row) + _tr_pr(height) + "".join(cells) + '</w:tr>')


def _card_row_for_cols(card_row: str, cols: int) -> str:
    """把文字行原型调整成恰好 cols 个单元格。

    内置原型固定 3 个单元格，而自定义小卡片可能排出 4 列以上；不足时以原型的
    第一个单元格为模板复制补齐（保留其字体/加粗/居中格式），多余时截断 ——
    避免出现"列数多于原型时末尾词组丢失"的问题。
    """
    cells = re.findall(r'<w:tc>.*?</w:tc>', card_row, re.S)
    if not cells or len(cells) == cols:
        return card_row
    if cols < len(cells):
        cells = cells[:cols]
    else:
        # 用"文字"单元格作为模板复制（原型第一个单元格即含占位文字）
        proto_cell = cells[0] if cells else ""
        cells = cells + [proto_cell] * (cols - len(cells))
    i = card_row.find('<w:tc>')
    j = card_row.rfind('</w:tc>') + len('</w:tc>')
    return card_row[:i] + "".join(cells) + card_row[j:]


def _make_word_row(card_row_xml, block_words, keep_tail, col_twips, height, cols,
                   max_pt=DEFAULT_TEXT_PT):
    """生成文字行：每个单元格按内容长度自动收缩字号（四周留 TEXT_PAD_MM），
    文字加粗并与预览一致地水平、垂直居中。"""
    row_proto = _card_row_for_cols(card_row_xml, cols)
    texts = []
    for c in range(cols):
        if c < len(block_words):
            texts.append(block_words[c])
        elif keep_tail:
            texts.append("文字")
        else:
            texts.append("")
    row = _set_cell_texts(row_proto, texts)
    row = _cell_width_fix(row, col_twips)
    row = re.sub(r'<w:trPr>.*?</w:trPr>', _tr_pr(height), row, count=1, flags=re.S)

    # 逐个单元格套用"按单元格收缩"后的字号
    cells = []
    for idx, tc in enumerate(re.findall(r'<w:tc>.*?</w:tc>', row, re.S)):
        word = texts[idx] if idx < len(texts) else ""
        pt = compute_word_pt(word, col_twips, height, max_pt)
        tc = _style_word_run(tc, int(round(pt * 2)))
        tc = tc.replace('<w:vAlign w:val="top"/>', '<w:vAlign w:val="center"/>')
        cells.append(tc)
    i, j = row.find('<w:tc>'), row.rfind('</w:tc>')
    if i >= 0 and j > i:
        row = row[:i] + "".join(cells) + row[j + len('</w:tc>'):]
    return _snap_off(row)


def _make_spacer_row(prototype_row, col_twips, height, cols):
    cells = []
    for _ in range(cols):
        cells.append(
            '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>'
            '<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:tc>' % col_twips)
    return (_tr_open(prototype_row) + _tr_pr(height) + "".join(cells) + '</w:tr>')


def _set_cell_texts(row_xml: str, texts):
    matches = list(re.finditer(r"(<w:t[^>]*>)([^<]*)(</w:t>)", row_xml))
    if not matches:
        return row_xml
    out, pos, ti = [], 0, 0
    for m in matches:
        out.append(row_xml[pos:m.start()])
        t = texts[ti] if ti < len(texts) else ""
        out.append('<w:t xml:space="preserve">%s</w:t>' % xml_escape(t))
        ti += 1
        pos = m.end()
    out.append(row_xml[pos:])
    return "".join(out)


def _rebuild_tbl_head(tbl_head: str, col_twips: int, cols: int,
                      h_align: str = "left") -> str:
    """按新列宽/列数重建表头（tblW、表格对齐、tblGrid）"""
    total = col_twips * cols
    # 内置表头中 tblW 为自闭合标签 <w:tblW .../>，需按其完整结束位置替换
    head = re.sub(r'<w:tblW\b[^>]*/>',
                  '<w:tblW w:w="%d" w:type="dxa"/>' % total, tbl_head, count=1)
    # 表格对齐：<w:jc> 必须紧跟 <w:tblW> 之后（CT_TblPrBase 元素顺序）
    jc = ('<w:jc w:val="center"/>' if h_align == "center"
          else '<w:jc w:val="left"/>')
    m = re.search(r'<w:tblW\b[^>]*/>', head)
    if '<w:jc ' not in head and m:
        head = head[:m.end()] + jc + head[m.end():]
    grid = '<w:tblGrid>' + ('<w:gridCol w:w="%d"/>' % col_twips) * cols + '</w:tblGrid>'
    head = re.sub(r'<w:tblGrid>.*?</w:tblGrid>', grid, head, flags=re.S)
    return head


# ------------------------- 布局计算 -------------------------

def compute_layout(card_w_cm, card_h_cm, page_w, page_h, rows_per_page=None,
                   spacer_enabled=False):
    """根据卡片尺寸与页面可用区域计算布局，返回 dict"""
    col_tw = cm_to_twips(card_w_cm)
    card_tw = cm_to_twips(card_h_cm)
    cols = max(1, page_w // col_tw)
    word_h = min(WORD_ROW_MAX, max(WORD_ROW_MIN, int(card_tw * 0.28)))
    pic_h = card_tw - word_h
    block_h = card_tw + (SPACER_HEIGHT if spacer_enabled else 0)
    auto_rows = max(1, page_h // block_h)
    rows = max(1, rows_per_page) if rows_per_page else auto_rows
    return {
        "cols": cols, "col_tw": col_tw, "card_tw": card_tw,
        "word_h": word_h, "pic_h": pic_h, "block_h": block_h,
        "rows_per_page": rows, "auto_rows": auto_rows,
        "per_page": cols * rows,
        "page_w": page_w, "page_h": page_h,
    }


def build_pages(words, tbl_head, card_row, spacer_row, layout,
                keep_tail_placeholder=False, with_picture=True,
                image_folder="", spacer_enabled=False, fit_mode="stretch",
                h_align="left", anchor=False, text_max_pt=DEFAULT_TEXT_PT):
    """生成整篇正文，返回 (正文XML, 图片关系列表)"""
    cols = layout["cols"]
    total_blocks = (len(words) + cols - 1) // cols
    img_index = [1]
    rels = []
    # 图片适配区域：(可用宽EMU, 可用高EMU)
    # 宽度 = 列宽 - 单元格左右边距；高度 = 图片行高 - 少量余量
    max_emu = (int((layout["col_tw"] - 300) * 635),
               int((layout["pic_h"] - 60) * 635))

    pages = []
    for block_idx in range(total_blocks):
        page_i, row_in_page = divmod(block_idx, layout["rows_per_page"])
        if row_in_page == 0:
            pages.append([])

        block_words = words[block_idx * cols:(block_idx + 1) * cols]

        if with_picture:
            pages[page_i].append(_make_picture_row(
                card_row, block_words, image_folder, img_index, rels,
                layout["col_tw"], layout["pic_h"], max_emu, cols, fit_mode))
        pages[page_i].append(_make_word_row(
            card_row, block_words, keep_tail_placeholder,
            layout["col_tw"], layout["word_h"], cols, text_max_pt))

        is_last = (block_idx == total_blocks - 1)
        if spacer_enabled and not is_last:
            pages[page_i].append(_make_spacer_row(
                spacer_row or card_row, layout["col_tw"], SPACER_HEIGHT, cols))

    head = _rebuild_tbl_head(tbl_head, layout["col_tw"], cols, h_align)
    # 满页基准：以"满页卡片"（rows_per_page 行 × 每块高）为基准求居中偏移，
    # 所有页（含卡片数不足的末页）统一使用该偏移 —— 满页时即双向居中，
    # 不满页时网格左上角仍与满页一致，不会下沉。
    top_off = 0
    if anchor:
        full_h = layout["rows_per_page"] * layout["block_h"]
        top_off = max(0, (layout["page_h"] - full_h) // 2)
    pre = _top_spacer(top_off) if top_off > 0 else ""
    tables = [pre + head + "".join(p) + "</w:tbl>" for p in pages]
    return PAGE_BREAK_P.join(tables), rels


# ------------------------- 主流程（自建 docx 包，无需模板） -------------------------

_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
# 主文档关系的命名空间与包关系不同！曾误用 _PKG_REL + "/officeDocument"，
# 导致 Word 找不到文档主体，报"文件已损坏"（生成的文件全部打不开）。
_OD_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# 随程序分发的"版式模板"资源：生成 docx 时从它补齐主题/字体表等
# Office 必需部件（缺失时 Word 会判定文件损坏），界面上不出现模板选择。
LAYOUT_TEMPLATE_NAME = "卡片版式模板.docx"


def resource_path(name: str) -> str:
    """定位随程序分发的资源文件（兼容 PyInstaller 打包后的内部目录）"""
    candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, name))
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), name))
        candidates.append(os.path.join(os.path.dirname(sys.executable),
                                       "_internal", name))
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        here = os.getcwd()
    candidates.append(os.path.join(here, name))
    for p in candidates:
        if os.path.isfile(p):
            return p
    return candidates[-1]


def _layout_extra_parts():
    """从随包版式模板中取出 Word 必需的补充部件（缺失时返回空字典）。

    说明：自建最小 docx 包只带 document/styles/settings 时，Word 会因缺少
    主题（theme）等部件判定"文件已损坏"。这里改为从随包模板复制这几件，
    既保持界面无模板选择，生成的文档又与模板版完全等价。
    """
    parts = {}
    path = resource_path(LAYOUT_TEMPLATE_NAME)
    if not os.path.isfile(path):
        return parts
    wanted = ("word/theme/theme1.xml", "word/fontTable.xml", "docProps/custom.xml")
    try:
        with zipfile.ZipFile(path, "r") as z:
            names = set(z.namelist())
            for n in wanted:
                if n in names:
                    parts[n] = z.read(n)
    except Exception:
        return {}
    return parts


def _docx_content_types(extra_parts=()) -> str:
    """内置内容类型表（覆盖所有支持的图片扩展名）"""
    defaults = ''.join(
        '<Default Extension="%s" ContentType="%s"/>' % (ext, ct)
        for ext, ct in sorted(EXT_CONTENT_TYPE.items()))
    extras = ''
    if "word/theme/theme1.xml" in extra_parts:
        extras += ('<Override PartName="/word/theme/theme1.xml" '
                   'ContentType="application/vnd.openxmlformats-officedocument.'
                   'theme+xml"/>')
    if "word/fontTable.xml" in extra_parts:
        extras += ('<Override PartName="/word/fontTable.xml" '
                   'ContentType="application/vnd.openxmlformats-officedocument.'
                   'wordprocessingml.fontTable+xml"/>')
    if "docProps/custom.xml" in extra_parts:
        extras += ('<Override PartName="/docProps/custom.xml" '
                   'ContentType="application/vnd.openxmlformats-officedocument.'
                   'custom-properties+xml"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types">'
            '<Default Extension="rels" ContentType="application/vnd.'
            'openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + defaults +
            '<Override PartName="/word/document.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.'
            'openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.extended-properties+xml"/>'
            + extras +
            '</Types>')


def _docx_root_rels(extra_parts=()) -> str:
    custom = ''
    if "docProps/custom.xml" in extra_parts:
        custom = ('<Relationship Id="rId4" Type="http://schemas.openxmlformats.'
                  'org/officeDocument/2006/relationships/custom-properties" '
                  'Target="docProps/custom.xml"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="%s/officeDocument" '
            'Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="%s/metadata/core-properties" '
            'Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="%s/extended-properties" '
            'Target="docProps/app.xml"/>'
            '%s'
            '</Relationships>'
            % (_PKG_REL, _OD_REL, _PKG_REL, _PKG_REL, custom))


def _docx_document_rels(image_rels, extra_parts=()) -> str:
    """document.xml 的关系：样式、设置、随包主题/字体表 + 所有图片"""
    doc_rel = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/")
    parts = ['<Relationship Id="rIdStyles" Type="%sstyles" Target="styles.xml"/>' % doc_rel,
             '<Relationship Id="rIdSettings" Type="%ssettings" Target="settings.xml"/>' % doc_rel]
    if "word/theme/theme1.xml" in extra_parts:
        parts.append('<Relationship Id="rIdTheme" Type="%stheme" '
                     'Target="theme/theme1.xml"/>' % doc_rel)
    if "word/fontTable.xml" in extra_parts:
        parts.append('<Relationship Id="rIdFontTable" Type="%sfontTable" '
                     'Target="fontTable.xml"/>' % doc_rel)
    for rel_id, media_name, _ in image_rels:
        parts.append('<Relationship Id="%s" Type="%simage" Target="media/%s"/>'
                     % (rel_id, doc_rel, media_name))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">%s</Relationships>' % (_PKG_REL, ''.join(parts)))


def _docx_doc_props() -> tuple:
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/'
            'package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/'
            'elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>卡片</dc:title><dc:creator>%s</dc:creator>'
            '<cp:lastModifiedBy>%s</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>' % (APP_AUTHOR, APP_AUTHOR, now, now))
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/'
           '2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/'
           'officeDocument/2006/docPropsVTypes">'
           '<Application>%s</Application><Company>%s</Company>'
           '<AppVersion>1.0</AppVersion></Properties>'
           % (APP_NAME, APP_COMPANY))
    return core, app


def fill_template(words, output_path: str,
                  card_width_cm, card_height_cm,
                  rows_per_page=None,
                  keep_tail_placeholder=False,
                  with_picture=True,
                  image_folder="",
                  spacer_enabled=False,
                  fit_mode="stretch",
                  margin_cm=DEFAULT_MARGIN_CM,
                  align_mode=DEFAULT_ALIGN_MODE,
                  _legacy_template_path=None):
    """生成卡片 docx（完全内置版式，不需要任何模板文件），返回 (成功?, 消息)"""
    if not words:
        return False, "没有可用的词组，请先输入或粘贴文字。"
    if image_folder and not os.path.isdir(image_folder):
        return False, "图片文件夹不存在：%s" % image_folder
    if card_width_cm <= 0 or card_height_cm <= 0:
        return False, "卡片尺寸无效。"
    if margin_cm is None or margin_cm < 0:
        margin_cm = DEFAULT_MARGIN_CM

    tbl_head, card_row, spacer_row, page_w, page_h = builtin_prototypes()

    margin_tw = cm_to_twips(margin_cm)
    # 对齐方式：
    #   anchor  → 表格水平居中 + 所有页统一使用"满页居中"基准的左上角
    #   topleft → 表格贴版心左上角（无任何偏移）
    #   center  → 表格水平居中 + 每页内容垂直居中
    h_align = "center" if align_mode in ("anchor", "center", "hcenter") else "left"
    anchor = (align_mode == "anchor")
    v_center = (align_mode == "center")
    content_w, content_h = content_size(page_w, page_h, margin_cm)
    layout = compute_layout(card_width_cm, card_height_cm, content_w, content_h,
                            rows_per_page, spacer_enabled)
    # 文字字号上限 = 内置文字行原型的原始字号（只缩不放）
    text_max_pt = proto_word_pt(card_row)
    body, image_rels = build_pages(
        words, tbl_head, card_row, spacer_row, layout,
        keep_tail_placeholder=keep_tail_placeholder,
        with_picture=with_picture,
        image_folder=image_folder,
        spacer_enabled=spacer_enabled,
        fit_mode=fit_mode,
        h_align=h_align,
        anchor=anchor,
        text_max_pt=text_max_pt)

    # —— 组装文档 XML：内置序言 + 正文（表格） + 尾部（1pt 空段落 + sectPr） ——
    tail = TINY_P + BUILTIN_SECTPR + '</w:body></w:document>'
    tail = _apply_page_setup(tail, margin_tw, v_center)
    new_doc = BUILTIN_DOC_PROLOGUE + body + tail

    core_xml, app_xml = _docx_doc_props()
    # 随包版式模板提供的必需部件（theme / fontTable / custom properties）：
    # 缺了它们 Word 会报"文件已损坏"，这里从内置资源里补齐
    extra_parts = _layout_extra_parts()
    tmp_path = output_path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("[Content_Types].xml", _docx_content_types(extra_parts))
        zout.writestr("_rels/.rels", _docx_root_rels(extra_parts))
        zout.writestr("docProps/core.xml", core_xml)
        zout.writestr("docProps/app.xml", app_xml)
        zout.writestr("word/document.xml", new_doc)
        zout.writestr("word/styles.xml", BUILTIN_STYLES_XML)
        zout.writestr("word/settings.xml", BUILTIN_SETTINGS_XML)
        zout.writestr("word/_rels/document.xml.rels",
                      _docx_document_rels(image_rels, extra_parts))
        for n, data in extra_parts.items():
            zout.writestr(n, data)
        for _, media_name, src_path in image_rels:
            zout.write(src_path, "word/media/%s" % media_name)

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except OSError:
            pass
    os.replace(tmp_path, output_path)

    total_blocks = (len(words) + layout["cols"] - 1) // layout["cols"]
    pages = (total_blocks + layout["rows_per_page"] - 1) // layout["rows_per_page"]
    msg = ("已生成 %d 个词组：卡片 %.2f×%.2f cm，每页 %d 列 × %d 行 = %d 张，共 %d 页；"
           "页边距四边 %.2f cm，网格%s。%s文件：\n%s"
           % (len(words), card_width_cm, card_height_cm,
              layout["cols"], layout["rows_per_page"], layout["per_page"], pages,
              margin_cm,
              {"anchor": "满页基准对齐（满页双向居中，不满页贴同一基准左上角）",
               "topleft": "贴版心左上角",
               "center": "每页各自居中"}.get(align_mode, "满页基准对齐"),
              ("已嵌入 %d 张图片，其余为占位符。" % len(image_rels)) if with_picture else "",
              output_path))
    return True, msg


# ------------------------- 文件名 -------------------------

def build_output_filename(custom_name=None):
    """默认：卡片【当前日期与时间】.docx；自定义支持 {日期} {时间} 占位符"""
    now = datetime.datetime.now()
    if custom_name is None or not custom_name.strip():
        return "卡片【%s】.docx" % now.strftime("%Y-%m-%d %H-%M-%S")
    name = custom_name.strip()
    name = (name.replace("{日期}", now.strftime("%Y-%m-%d"))
                .replace("{时间}", now.strftime("%H-%M-%S")))
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    if not name:
        return "卡片【%s】.docx" % now.strftime("%Y-%m-%d %H-%M-%S")
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name


def unique_output_path(out_dir: str, filename: str) -> str:
    base, ext = os.path.splitext(filename)
    p = os.path.join(out_dir, filename)
    i = 2
    while os.path.exists(p):
        p = os.path.join(out_dir, "%s(%d)%s" % (base, i, ext))
        i += 1
    return p


def make_output_path(out_dir: str, custom_name=None) -> str:
    return unique_output_path(out_dir, build_output_filename(custom_name))


# ------------------------- 百度搜图对话框 -------------------------

class ImageSearchDialog(tk.Toplevel):
    """百度图片搜索窗口：为缺失图片的词组联网搜图，确认后存入图片文件夹。

    用法：ImageSearchDialog(master, word, image_dir, on_saved=回调)
    单击缩略图选中、双击（或点"使用选中的图片"）下载大图并保存为
    <词组>.jpg；保存成功后触发 on_saved(word, path)。
    """

    COLS = 4
    THUMB_W = 150
    THUMB_H = 120
    PAGE_RN = 60        # 接口单次最多返回 60 条
    MAX_ITEMS = 600     # 上限，避免无限加载
    THUMB_THREADS = 4   # 缩略图并行下载线程数

    def __init__(self, master, word, image_dir, on_saved=None):
        tk.Toplevel.__init__(self, master)
        self.title("百度搜图 · %s" % word)
        self.word = word
        self.image_dir = image_dir
        self.on_saved = on_saved
        self.results = []
        self._selected = None
        self._cells = {}
        self._photos = {}
        self._closed = False
        self._busy = False
        self._loading_more = False
        self._pn = 0            # 翻页游标（已取到的原始条数）
        self._seen = set()      # 已展示过的缩略图 URL，用于去重
        # 后台线程 → 主线程的安全通道（tkinter 不允许跨线程操作控件）
        self._queue = queue.Queue()

        self.geometry("740x600")
        self.minsize(640, 480)
        self.transient(master.winfo_toplevel())
        self.bind("<Escape>", lambda e: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        # 顶部说明
        top = ttk.Frame(self, padding=(10, 8, 10, 4))
        top.pack(fill="x")
        ttk.Label(top, text="为“%s”搜索图片（来源：百度图片）——单击选中，双击直接使用"
                  % word, font=("微软雅黑", 10, "bold")).pack(side="left")

        # 中部结果区（画布 + 滚动条）
        mid = ttk.Frame(self, padding=(10, 0))
        mid.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(mid, highlightthickness=0, bg="#ffffff")
        vs = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.grid_frm = ttk.Frame(self.canvas)
        self._grid_win = self.canvas.create_window((0, 0), window=self.grid_frm,
                                                   anchor="nw")
        self.grid_frm.bind("<Configure>", self._on_grid_conf)
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind_all(seq, self._on_wheel)

        # 底部：状态 + 按钮
        self.var_status = tk.StringVar(value="正在搜索…")
        ttk.Label(self, textvariable=self.var_status, foreground="#5a6472",
                  padding=(10, 2)).pack(side="bottom", anchor="w")
        btns = ttk.Frame(self, padding=(10, 6))
        btns.pack(side="bottom", fill="x")
        self.btn_use = ttk.Button(btns, text="使用选中的图片", state="disabled",
                                  command=self.use_selected)
        self.btn_use.pack(side="left")
        self.btn_more = ttk.Button(btns, text="加载更多…", state="disabled",
                                   command=self.load_more)
        self.btn_more.pack(side="left", padx=8)
        ttk.Button(btns, text="从本地文件选择…",
                   command=self.use_local_file).pack(side="left")
        ttk.Button(btns, text="在浏览器中打开搜索页",
                   command=self.open_browser).pack(side="left", padx=8)
        ttk.Button(btns, text="关闭", command=self.destroy).pack(side="right")

        self.after(60, self._pump_queue)   # 主线程轮询后台回调
        threading.Thread(target=self._search_worker, daemon=True).start()

    # ---------- 布局与滚动 ----------

    def _on_grid_conf(self, _e):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfigure(self._grid_win,
                                  width=max(self.canvas.winfo_width(),
                                            self.grid_frm.winfo_reqwidth()))

    def _on_wheel(self, e):
        if self._closed:
            return
        if getattr(e, "num", None) == 4 or e.delta > 0:
            self.canvas.yview_scroll(-2, "units")
        elif getattr(e, "num", None) == 5 or e.delta < 0:
            self.canvas.yview_scroll(2, "units")

    def _q(self, fn):
        """线程安全：后台线程把回调塞进队列，由主线程统一执行"""
        if not self._closed:
            self._queue.put(fn)

    def _pump_queue(self):
        """主线程轮询：执行后台线程投递的回调（tkinter 只能主线程碰控件）"""
        if self._closed:
            return
        while True:
            try:
                fn = self._queue.get_nowait()
            except queue.Empty:
                break
            try:
                fn()
            except Exception:
                pass
        try:
            self.after(60, self._pump_queue)
        except Exception:
            pass

    # ---------- 搜索与缩略图加载 ----------

    def _search_worker(self):
        try:
            items = baidu_image_search(self.word, pn=0, rn=self.PAGE_RN)
        except Exception as ex:
            self._q(lambda e=ex: self._search_failed(e))
            return
        self._q(lambda: self._populate(items))

    def _search_failed(self, ex):
        if self._closed:
            return
        self.btn_more.config(state="normal")
        self.var_status.set("搜索失败：%s" % ex)

    def _populate(self, items):
        """首屏结果：清空网格后填入"""
        if self._closed:
            return
        for w in self.grid_frm.winfo_children():
            w.destroy()
        self.results = []
        self._cells = {}
        self._photos = {}
        self._selected = None
        self._pn = 0
        self._seen = set()
        self.btn_use.config(state="disabled")
        for i in range(self.COLS):
            self.grid_frm.columnconfigure(i, weight=1)
        self._append(items, first=True)

    def _append(self, items, first=False):
        """把一批结果追加到网格（按缩略图 URL 去重），返回新增数量"""
        items = items or []
        self._pn += len(items)
        new_idx = []
        for item in items:
            thumb = item.get("thumb")
            if not thumb or thumb in self._seen:
                continue
            if len(self.results) >= self.MAX_ITEMS:
                break
            self._seen.add(thumb)
            self.results.append(item)
            idx = len(self.results) - 1
            self._make_cell(idx, item)
            new_idx.append(idx)
        if new_idx:                      # 缩略图分批并行下载
            self._start_thumbs(new_idx)
        if not self.results:
            self.btn_more.config(state="disabled", text="加载更多…")
            self.var_status.set(
                "没有搜到结果，可点击“在浏览器中打开搜索页”手动查找。")
            return len(new_idx)
        if len(self.results) >= self.MAX_ITEMS:
            self.btn_more.config(state="disabled", text="已达上限")
            self.var_status.set("已加载 %d 张（本次上限 %d 张）。"
                                % (len(self.results), self.MAX_ITEMS))
        else:
            self.btn_more.config(state="normal", text="加载更多…")
            self.var_status.set(
                "已加载 %d 张%s，单击选中 / 双击直接使用；"
                "点“加载更多”继续搜索。" % (
                    len(self.results),
                    "（缩略图加载中…）" if new_idx else ""))
        return len(new_idx)

    def load_more(self):
        """向后翻页，继续追加结果（接口单次上限 60 条）"""
        if self._closed or self._loading_more or self._busy:
            return
        if len(self.results) >= self.MAX_ITEMS:
            return
        self._loading_more = True
        self.btn_more.config(state="disabled", text="加载中…")
        self.var_status.set("正在加载更多结果…")
        pn = self._pn

        def worker():
            try:
                items = baidu_image_search(self.word, pn=pn, rn=self.PAGE_RN)
                err = None
            except Exception as ex:
                items, err = [], ex
            self._q(lambda i=items, e=err: self._load_more_done(i, e))

        threading.Thread(target=worker, daemon=True).start()

    def _load_more_done(self, items, err):
        self._loading_more = False
        if self._closed:
            return
        if err is not None:
            self.btn_more.config(state="normal", text="加载更多…")
            self.var_status.set("加载失败：%s" % err)
            return
        before = len(self.results)
        added = self._append(items, first=False)
        if not added:
            self.btn_more.config(state="disabled", text="没有更多了")
            self.var_status.set("没有更多结果了（共 %d 张）。" % before)

    def _make_cell(self, idx, item):
        frm = tk.Frame(self.grid_frm, bd=1, relief="solid", bg="#f7f8fa")
        frm.grid(row=idx // self.COLS, column=idx % self.COLS,
                 padx=6, pady=6, sticky="nsew")
        body = tk.Canvas(frm, width=self.THUMB_W, height=self.THUMB_H,
                         bg="#eef0f3", highlightthickness=0)
        body.create_text(self.THUMB_W / 2.0, self.THUMB_H / 2.0,
                         text="加载中…", fill="#9aa3ad", font=("微软雅黑", 9))
        body.pack(padx=2, pady=(2, 0))
        size = ""
        if item.get("w") and item.get("h"):
            size = "%d × %d" % (item["w"], item["h"])
        cap = tk.Label(frm, text=size or " ", bg="#f7f8fa", fg="#66707c",
                       font=("微软雅黑", 8))
        cap.pack(pady=(0, 2))
        self._cells[idx] = {"frm": frm, "body": body, "cap": cap}
        for w in (frm, body, cap):
            w.bind("<Button-1>", lambda e, i=idx: self._select(i))
            w.bind("<Double-Button-1>", lambda e, i=idx: (self._select(i),
                                                          self.use_selected()))

    def _start_thumbs(self, indices):
        """用若干线程并行下载缩略图，结果统一回到主线程绘制"""
        queue = list(indices)
        lock = threading.Lock()

        def run():
            while not self._closed:
                with lock:
                    if not queue:
                        return
                    idx = queue.pop(0)
                try:
                    item = self.results[idx]
                except IndexError:
                    return
                try:
                    _data, im = fetch_image_bytes([item["thumb"]])
                except Exception:
                    self._q(lambda i=idx: self._mark_fail(i))
                    continue
                self._q(lambda i=idx, im=im: self._show_thumb(i, im))

        for _ in range(max(1, min(self.THUMB_THREADS, len(queue)))):
            threading.Thread(target=run, daemon=True).start()

    def _show_thumb(self, idx, im):
        if self._closed:
            return
        cell = self._cells.get(idx)
        if not cell:
            return
        try:
            box_w, box_h = self.THUMB_W, self.THUMB_H
            scale = min(box_w / im.width, box_h / im.height, 1.0)
            w = max(1, int(im.width * scale))
            h = max(1, int(im.height * scale))
            resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1)
            photo = ImageTk.PhotoImage(im.resize((w, h), resample))
        except Exception:
            self._mark_fail(idx)
            return
        self._photos[idx] = photo
        cell["body"].destroy()
        cell["body"] = tk.Label(cell["frm"], image=photo, bd=0,
                                bg="#f7f8fa", cursor="hand2")
        # before=cap：让图片排在尺寸标注上方（替换占位框后保持原顺序）
        cell["body"].pack(padx=2, pady=(2, 0), before=cell["cap"])
        cell["body"].bind("<Button-1>", lambda e, i=idx: self._select(i))
        cell["body"].bind("<Double-Button-1>",
                          lambda e, i=idx: (self._select(i), self.use_selected()))

    def _mark_fail(self, idx):
        if self._closed:
            return
        cell = self._cells.get(idx)
        if not cell:
            return
        body = cell["body"]
        if isinstance(body, tk.Canvas):
            body.delete("all")
            body.create_text(self.THUMB_W / 2.0, self.THUMB_H / 2.0,
                             text="加载失败", fill="#c07070",
                             font=("微软雅黑", 9))

    # ---------- 选择与保存 ----------

    def _select(self, idx):
        if self._closed or idx not in self._cells:
            return
        self._selected = idx
        for i, cell in self._cells.items():
            on = (i == idx)
            cell["frm"].config(bg="#cfe0ff" if on else "#f7f8fa")
            cell["cap"].config(bg="#cfe0ff" if on else "#f7f8fa")
        self.btn_use.config(state="normal")
        item = self.results[idx]
        self.var_status.set("已选中第 %d 张%s。点击“使用选中的图片”下载并保存。"
                            % (idx + 1, "（%d×%d）" % (item["w"], item["h"])
                               if item.get("w") else ""))

    def _set_busy(self, busy):
        self._busy = busy
        st = "disabled" if busy else "normal"
        self.btn_use.config(state=st)
        if self._selected is None:
            self.btn_use.config(state="disabled")

    def use_selected(self):
        if self._closed or self._busy:
            return
        if self._selected is None or self._selected >= len(self.results):
            self.var_status.set("请先单击一张图片。")
            return
        item = self.results[self._selected]
        self._set_busy(True)
        self.var_status.set("正在下载图片…")

        def worker():
            try:
                _data, im = fetch_image_bytes(item["cands"], min_px=380)
            except Exception as ex:
                self._q(lambda e=ex: (self.var_status.set("下载失败：%s" % e),
                                       self._set_busy(False)))
                return
            self._q(lambda: self._confirm_save(im))

        threading.Thread(target=worker, daemon=True).start()

    def use_local_file(self):
        if self._closed or self._busy:
            return
        p = filedialog.askopenfilename(
            title="选择本地图片（将另存为 %s.jpg）" % self.word,
            filetypes=[("图片", "*.jpg *.jpeg *.png *.gif *.bmp *.webp"),
                       ("所有文件", "*.*")], parent=self)
        if not p:
            return
        try:
            im = Image.open(p)
            im.load()
        except Exception as ex:
            messagebox.showerror("失败", "无法读取图片：%s" % ex, parent=self)
            return
        self._confirm_save(im)

    def _confirm_save(self, im):
        path = os.path.join(self.image_dir, self.word + ".jpg")
        if os.path.exists(path) and not messagebox.askyesno(
                "覆盖确认", "文件已存在：\n%s\n\n是否覆盖？" % path, parent=self):
            self._set_busy(False)
            return
        try:
            saved = save_word_image(self.word, self.image_dir, im)
        except Exception as ex:
            messagebox.showerror("失败", "保存图片失败：%s" % ex, parent=self)
            self._set_busy(False)
            return
        self.var_status.set("已保存：%s" % saved)
        if self.on_saved:
            try:
                self.on_saved(self.word, saved)
            except Exception:
                pass
        self.after(500, self.destroy)

    def open_browser(self):
        webbrowser.open("https://image.baidu.com/search/index?tn=baiduimage"
                        "&ie=utf-8&word=%s" % urllib.parse.quote(self.word))

    def destroy(self):
        self._closed = True
        try:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.unbind_all(seq)
        except Exception:
            pass
        tk.Toplevel.destroy(self)


# ------------------------- 右侧文档预览 -------------------------

# 预览配色
PV_BG = "#eceff3"           # 画布底色
PV_PAPER = "#ffffff"        # 纸张
PV_PAPER_LINE = "#c9ced6"   # 纸张描边
PV_SHADOW = "#d8dde4"       # 纸张投影
PV_MARGIN_LINE = "#aebbc9"  # 版心（页边距）虚线
PV_GRID_LINE = "#b6c3d2"    # 卡片边框
PV_PIC_FILL = "#f1f3f6"     # 图片占位底色
PV_PIC_TEXT = "#a7aeb8"     # 图片占位文字
PV_WORD_TEXT = "#2b2f36"    # 词组文字
PV_PAD = 14                 # 画布内边距（px）


class PreviewPane(ttk.LabelFrame):
    """右侧文档预览面板。

    按生成时使用的真实布局（纸张、页边距、网格锚点偏移、行列数、图片
    适配方式）绘制一页缩略图；多页时用翻页按钮切换。参数一变就自动
    重绘，图片缩略图带缓存以免反复解码。
    """

    def __init__(self, master, app):
        ttk.LabelFrame.__init__(self, master, text=" 文档预览（实时） ",
                                padding=(6, 4))
        self.app = app
        self.page = 0
        self.total_pages = 1
        self._job = None            # 防抖用的 after id
        self._thumbs = {}           # 缩略图缓存
        self._imgs = []             # 持有 PhotoImage 引用，防止被回收
        self._missing_rects = []    # 本页缺失图片的区域 [(x0,y0,x1,y1,word)]
        self._image_rects = []      # 本页已有图片的区域 [(x0,y0,x1,y1,path,word)]

        # ---- 翻页工具条 ----
        bar = ttk.Frame(self)
        bar.pack(fill="x")
        self.btn_prev = ttk.Button(bar, text="◀ 上一页", width=9,
                                   command=self.prev_page)
        self.btn_prev.pack(side="left")
        self.var_page = tk.StringVar(value="第 1 / 1 页")
        ttk.Label(bar, textvariable=self.var_page, width=11,
                  anchor="center").pack(side="left", padx=6)
        self.btn_next = ttk.Button(bar, text="下一页 ▶", width=9,
                                   command=self.next_page)
        self.btn_next.pack(side="left")

        # ---- 画布 ----
        self.canvas = tk.Canvas(self, bg=PV_BG, highlightthickness=1,
                                highlightbackground="#d5dae0")
        self.canvas.pack(fill="both", expand=True, pady=(6, 4))
        self.canvas.bind("<Configure>", lambda e: self.schedule_refresh(60))
        # 单击缺图卡片 → 百度搜图；单击已有图片 → 文件夹定位；右键 → 更多操作
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<Button-3>", self._on_canvas_right_click)

        # ---- 底部开关与提示 ----
        foot = ttk.Frame(self)
        foot.pack(fill="x")
        self.var_thumbs = tk.BooleanVar(value=True)
        self.chk_thumbs = ttk.Checkbutton(foot, text="显示缩略图",
                                          variable=self.var_thumbs,
                                          command=self.draw)
        self.chk_thumbs.pack(side="left")
        ttk.Button(foot, text="刷新", width=5,
                   command=self.reload_thumbs).pack(side="left", padx=(6, 0))
        ttk.Label(foot, text="点缺图→联网搜图 · 点图片→文件夹定位 · 右键更多",
                  foreground="#5a6472").pack(side="left", padx=(8, 0))
        self.var_tip = tk.StringVar(value="")
        ttk.Label(foot, textvariable=self.var_tip,
                  foreground="#5a6472").pack(side="right")
        if not _HAS_PIL:
            self.var_thumbs.set(False)
            self.chk_thumbs.state(["disabled"])

    # ---------- 刷新调度 ----------

    def schedule_refresh(self, delay=240):
        """合并短时间内的多次刷新请求，避免连续输入时反复重绘"""
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        self._job = self.after(delay, self.draw)

    def prev_page(self):
        if self.page > 0:
            self.page -= 1
            self.draw()

    def next_page(self):
        if self.page < self.total_pages - 1:
            self.page += 1
            self.draw()

    # ---------- 图片缩略图 ----------

    def _thumb(self, path, box_w, box_h, mode):
        """按适配方式生成缩略图（结果缓存）；失败返回 None"""
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return None
        tw = max(1, int(round(box_w)))
        th = max(1, int(round(box_h)))
        key = (path, mtime, tw, th, mode)
        hit = self._thumbs.get(key)
        if hit is not None:
            return hit
        if len(self._thumbs) > 400:
            self._thumbs.clear()
        try:
            im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
            if mode == "stretch":
                out = im.resize((tw, th), _RESAMPLE)
            elif mode == "cover":
                try:
                    out = ImageOps.fit(im, (tw, th), method=_RESAMPLE,
                                       centering=(0.5, 0.5))
                except TypeError:
                    out = ImageOps.fit(im, (tw, th), centering=(0.5, 0.5))
            else:  # contain
                out = im.copy()
                out.thumbnail((tw, th), _RESAMPLE)
            photo = ImageTk.PhotoImage(out)
        except Exception:
            return None
        self._thumbs[key] = photo
        return photo

    # ---------- 绘制 ----------

    def draw(self):
        self._job = None
        c = self.canvas
        try:
            cw = c.winfo_width()
            ch = c.winfo_height()
        except Exception:
            return
        if cw < 40 or ch < 40:
            return
        c.delete("all")
        self._imgs = []
        self._missing_rects = []
        self._image_rects = []

        words = self.app.get_words()
        if not words:
            c.create_text(cw / 2.0, ch / 2.0, fill="#9aa3ad", justify="center",
                          font=("微软雅黑", 11),
                          text="输入词组后\n这里显示文档排版预览")
            self.var_page.set("第 0 / 0 页")
            self.var_tip.set("")
            self.btn_prev.state(["disabled"])
            self.btn_next.state(["disabled"])
            return

        layout = self.app.get_layout()[0]
        cols = layout["cols"]
        rows = layout["rows_per_page"]
        total_blocks = (len(words) + cols - 1) // cols
        self.total_pages = max(1, (total_blocks + rows - 1) // rows)
        self.page = min(max(0, self.page), self.total_pages - 1)

        page_w_tw, page_h_tw = template_page_size()
        margin_tw = cm_to_twips(self.app.get_margin_cm())
        content_w_tw = max(1, page_w_tw - 2 * margin_tw)
        content_h_tw = max(1, page_h_tw - 2 * margin_tw)

        # 纸张在画布内最大化显示，并保持纸张长宽比
        avail_w = max(20, cw - 2 * PV_PAD)
        avail_h = max(20, ch - 2 * PV_PAD)
        ratio = page_h_tw / float(page_w_tw)
        paper_w = avail_w
        paper_h = paper_w * ratio
        if paper_h > avail_h:
            paper_h = avail_h
            paper_w = paper_h / ratio
        scale = paper_w / float(page_w_tw)
        ox = (cw - paper_w) / 2.0
        oy = (ch - paper_h) / 2.0

        # 纸张（带投影）与版心虚线框
        c.create_rectangle(ox + 3, oy + 3, ox + paper_w + 3, oy + paper_h + 3,
                           fill=PV_SHADOW, outline="")
        c.create_rectangle(ox, oy, ox + paper_w, oy + paper_h,
                           fill=PV_PAPER, outline=PV_PAPER_LINE)
        mx = ox + margin_tw * scale
        my = oy + margin_tw * scale
        c.create_rectangle(mx, my, mx + content_w_tw * scale, my + content_h_tw * scale,
                           outline=PV_MARGIN_LINE, dash=(3, 3))

        col_px = layout["col_tw"] * scale
        pic_px = layout["pic_h"] * scale
        word_px = layout["word_h"] * scale

        # —— 文字样式与生成逻辑共用同一套规则，保证预览 = 输出 ——
        # 磅 → 像素：1 磅 = 1/72 英寸 = 20 缇 → px = pt × 20 × (px/缇)。
        # Tk 字号用负数表示"像素"，避免受系统 DPI 缩放影响导致偏大。
        px_per_cm = (1440.0 / 2.54) * scale
        pad_px = TEXT_PAD_MM / 10.0 * px_per_cm              # 四周 2mm 内边距
        text_area_px = max(10.0, (layout["col_tw"] - 2 * CELL_MAR_TW) * scale)
        wrap_px = max(8.0, text_area_px - 2 * pad_px)        # 与 docx 同宽换行
        max_pt = self.app.get_word_max_pt()
        pic_font = ("微软雅黑", max(5, int(round(10.5 * 20 * scale))))

        # 与生成逻辑保持一致的对齐偏移
        align_mode = self.app.get_align_mode()
        h_center = align_mode in ("anchor", "center", "hcenter")
        left_off = ((content_w_tw - cols * layout["col_tw"]) / 2.0) if h_center else 0.0
        block_start = self.page * rows
        n_rows = max(1, min(rows, total_blocks - block_start))
        if align_mode == "anchor":
            top_off = max(0.0, (content_h_tw - rows * layout["block_h"]) / 2.0)
        elif align_mode == "center":
            top_off = max(0.0, (content_h_tw - n_rows * layout["block_h"]) / 2.0)
        else:
            top_off = 0.0

        img_dir = self.app.var_imgdir.get().strip() if self.app.var_with_pic.get() else ""
        fit_mode = self.app.var_fit.get()
        show_pic = bool(img_dir) and self.var_thumbs.get()
        with_pic = self.app.var_with_pic.get()

        drawn = 0
        for r in range(n_rows):
            for cc in range(cols):
                wi = (block_start + r) * cols + cc
                if wi >= len(words):
                    break
                # 与输出同一算法：按单元格可用空间收缩字号，加粗
                wt = compute_word_pt(words[wi], layout["col_tw"], layout["word_h"],
                                     max_pt)
                word_font = ("微软雅黑",
                             # 向下取整（不用 round）：临界字号取整后只会更窄不会更宽，
                             # 与 compute_word_pt 的余量配合，保证预览不出现假换行
                             -max(5, int(wt * 20 * scale)), "bold")
                self._draw_card(
                    c,
                    mx + (left_off + cc * layout["col_tw"]) * scale,
                    my + (top_off + r * layout["block_h"]) * scale,
                    col_px, pic_px, word_px, words[wi],
                    img_dir, fit_mode, show_pic, with_pic,
                    word_font=word_font, wrap_px=wrap_px, pic_font=pic_font,
                    missing_list=self._missing_rects,
                    image_list=self._image_rects)
                drawn += 1

        self.var_page.set("第 %d / %d 页" % (self.page + 1, self.total_pages))
        self.btn_prev.state(["!disabled"] if self.page > 0 else ["disabled"])
        self.btn_next.state(["!disabled"]
                            if self.page < self.total_pages - 1 else ["disabled"])
        self.var_tip.set("本页 %d 张 · 每页 %d列×%d行=%d张"
                         % (drawn, cols, rows, layout["per_page"]))

    def _draw_card(self, c, x0, y0, col_px, pic_px, word_px, word,
                   img_dir, fit_mode, show_pic, with_pic=True,
                   word_font=None, wrap_px=None, pic_font=None,
                   missing_list=None, image_list=None):
        """画一张卡片：上行图片区 + 下行文字区（未勾选图片行时只有文字区）
        word_font / wrap_px 与 docx 生成共用同一算法，保证预览所见即所得。"""
        if word_font is None:
            word_font = ("微软雅黑", -12, "bold")
        if wrap_px is None:
            wrap_px = max(10, col_px - 6)
        if pic_font is None:
            pic_font = ("微软雅黑", -9)

        if not with_pic:
            c.create_rectangle(x0, y0, x0 + col_px, y0 + word_px,
                               fill=PV_PAPER, outline=PV_GRID_LINE)
            c.create_text(x0 + col_px / 2.0, y0 + word_px / 2.0,
                          text=word, fill=PV_WORD_TEXT,
                          font=word_font, width=wrap_px)
            return

        c.create_rectangle(x0, y0, x0 + col_px, y0 + pic_px + word_px,
                           fill=PV_PAPER, outline=PV_GRID_LINE)

        box_w = max(2.0, col_px - 2)
        box_h = max(2.0, pic_px - 2)
        # 只要磁盘上存在该词组的图片文件，就按"已有图片"处理（可点击定位到
        # 所在文件夹）；文件不存在才算缺图（点击联网搜图）。
        path = find_image_for(word, img_dir) if img_dir else None
        photo = self._thumb(path, box_w, box_h, fit_mode) if (show_pic and path) else None

        if photo is not None:
            pw, ph = photo.width(), photo.height()
            c.create_image(x0 + 1 + (box_w - pw) / 2.0,
                           y0 + 1 + (box_h - ph) / 2.0,
                           image=photo, anchor="nw")
            self._imgs.append(photo)
        else:
            c.create_rectangle(x0 + 1, y0 + 1, x0 + 1 + box_w, y0 + 1 + box_h,
                               fill=PV_PIC_FILL, outline="")
            c.create_text(x0 + col_px / 2.0, y0 + pic_px / 2.0,
                          text="【图片】", fill=PV_PIC_TEXT, font=pic_font)

        rect = (x0 + 1, y0 + 1, x0 + 1 + box_w, y0 + 1 + box_h)
        if path:
            if image_list is not None:
                image_list.append(rect + (path, word))
        elif missing_list is not None:
            missing_list.append(rect + (word,))

        c.create_text(x0 + col_px / 2.0, y0 + pic_px + word_px / 2.0,
                      text=word, fill=PV_WORD_TEXT,
                      font=word_font, width=wrap_px)

    # ---------- 点击卡片：缺图搜图 / 已有图片定位 ----------

    def reload_thumbs(self):
        """清空缩略图缓存后重绘（在文件夹里换了图之后用）"""
        self._thumbs.clear()
        self.draw()

    def _hit_image(self, ex, ey):
        for x0, y0, x1, y1, path, word in self._image_rects:
            if x0 <= ex <= x1 and y0 <= ey <= y1:
                return path, word
        return None, None

    def _hit_missing(self, ex, ey):
        for x0, y0, x1, y1, word in self._missing_rects:
            if x0 <= ex <= x1 and y0 <= ey <= y1:
                return word
        return None

    def _on_canvas_click(self, e):
        # 已有图片 → 打开所在文件夹并选中该图片
        path, word = self._hit_image(e.x, e.y)
        if path:
            self.reveal_image(path, word)
            return
        word = self._hit_missing(e.x, e.y)
        if word:
            self._open_search(word)

    def _on_canvas_right_click(self, e):
        """右键卡片图片：在文件夹中显示 / 打开图片 / 更换图片"""
        path, word = self._hit_image(e.x, e.y)
        missing_word = None if path else self._hit_missing(e.x, e.y)
        if not path and not missing_word:
            return
        m = tk.Menu(self, tearoff=0)
        if path:
            m.add_command(label="在文件夹中显示",
                          command=lambda: self.reveal_image(path, word))
            m.add_command(label="用默认看图程序打开",
                          command=lambda: open_image_file(path))
            m.add_separator()
            m.add_command(label="更换图片（百度搜图）…",
                          command=lambda: self._open_search(word))
            m.add_command(label="从本地文件更换…",
                          command=lambda: self._replace_from_local(word))
            m.add_separator()
            m.add_command(label="打开图片文件夹",
                          command=lambda: reveal_in_folder(path))
        else:
            m.add_command(label="联网搜图补图…",
                          command=lambda: self._open_search(missing_word))
            m.add_command(label="从本地文件选择…",
                          command=lambda: self._replace_from_local(missing_word))
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()

    def reveal_image(self, path, word=None):
        """打开图片所在文件夹并选中该文件"""
        if not path or not os.path.isfile(path):
            messagebox.showinfo("提示", "图片文件不存在：\n%s" % path, parent=self)
            return
        ok = reveal_in_folder(path)
        if ok:
            self.app.var_status.set(
                "已在文件夹中定位：%s%s"
                % (path, "（词组“%s”）" % word if word else ""))
        else:
            messagebox.showinfo("提示", "无法打开文件夹，图片位于：\n%s" % path,
                                parent=self)

    def _replace_from_local(self, word):
        """用本地图片文件替换该词组的图片（另存为 词组.jpg）"""
        img_dir = self.app.var_imgdir.get().strip()
        if not (img_dir and os.path.isdir(img_dir)):
            messagebox.showinfo("提示", "请先设置有效的图片文件夹。", parent=self)
            return
        if not _HAS_PIL:
            messagebox.showinfo("提示", "更换图片需要 Pillow 支持。", parent=self)
            return
        p = filedialog.askopenfilename(
            title="选择图片（将另存为 %s.jpg）" % word, parent=self,
            filetypes=[("图片", "*.jpg *.jpeg *.png *.gif *.bmp *.webp"),
                       ("所有文件", "*.*")])
        if not p:
            return
        try:
            im = Image.open(p)
            im.load()
            saved = save_word_image(word, img_dir, im)
        except Exception as ex:
            messagebox.showerror("失败", "替换图片失败：%s" % ex, parent=self)
            return
        self._thumbs.clear()
        self.draw()
        self.app.var_status.set("已更换词组“%s”的图片：%s" % (word, saved))

    def _open_search(self, word):
        """点击缺失图片的卡片：确保图片文件夹有效后打开百度搜图窗口"""
        if not _HAS_PIL:
            messagebox.showinfo(
                "提示", "百度搜图需要 Pillow 库支持缩略图显示。\n"
                       "当前环境未安装 Pillow，请先安装后重试。",
                parent=self)
            return
        img_dir = self.app.var_imgdir.get().strip()
        if not (img_dir and os.path.isdir(img_dir)):
            if not messagebox.askyesno(
                    "提示", "尚未设置图片文件夹。\n\n"
                           "是否现在选择用于保存搜到的图片的文件夹？",
                    parent=self):
                return
            img_dir = filedialog.askdirectory(title="选择图片文件夹", parent=self)
            if not img_dir or not os.path.isdir(img_dir):
                return
            self.app.var_imgdir.set(img_dir)
        ImageSearchDialog(self, word, img_dir,
                          on_saved=self._on_image_saved)

    def _on_image_saved(self, word, path):
        """搜图保存成功：清空缩略图缓存并重绘预览，词组立即有图"""
        try:
            self._thumbs.clear()
            self.draw()
            self.app.var_status.set("已保存图片：%s（词组“%s”已补图）"
                                    % (path, word))
        except Exception:
            pass


# ----------------------------- GUI -----------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("%s v%s（紧凑界面 · 实时预览 · 百度搜图）" % (APP_NAME, APP_VERSION))
        # 窗口尺寸按屏幕自适应并居中，避免在小屏/高缩放下超出屏幕
        # v4.9：参数栏压缩后整体更窄，默认窗口也相应收窄
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        win_w = max(980, min(1160, sw - 80))
        win_h = max(680, min(880, sh - 100))
        root.geometry("%dx%d+%d+%d" % (win_w, win_h,
                                       max(0, (sw - win_w) // 2),
                                       max(0, (sh - win_h) // 2 - 20)))
        root.minsize(950, 640)

        # 统一控件字体（紧凑）与主题
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure(".", font=("微软雅黑", 9))

        # ---- 菜单栏（帮助 / 关于）----
        menubar = tk.Menu(root)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于 卡片批量生成器", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        root.config(menu=menubar)

        cfg = self.load_config()

        # 左右两栏：左侧参数与输入，右侧文档预览。
        # 这里用 grid + columnconfigure(minsize) 而不用 PanedWindow：
        # PanedWindow 需要在天窗映射后调 sashpos 才能定位分隔条，若在窗口
        # 尚未映射时调用会被钳制，导致左栏被压成 0 宽（"只剩预览窗"）。
        # grid 的 minsize 是硬保证，任何情况下左栏都不会消失。
        outer = ttk.Frame(root)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, minsize=580, weight=3)
        outer.columnconfigure(1, minsize=320, weight=2)
        outer.rowconfigure(0, weight=1)

        frm = ttk.Frame(outer, padding=(8, 6))
        frm.grid(row=0, column=0, sticky="nsew")

        self.preview = PreviewPane(outer, self)
        self.preview.grid(row=0, column=1, sticky="nsew")

        # ---- 基本文件设置（不再需要模板文件，版式为内置） ----
        ttk.Label(frm, text="输出目录：").grid(row=0, column=0, sticky="w")
        self.var_outdir = tk.StringVar(value=cfg.get("outdir", self.default_outdir(cfg)))
        ttk.Entry(frm, textvariable=self.var_outdir).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(frm, text="浏览…", command=self.choose_outdir).grid(row=0, column=2)

        ttk.Label(frm, text="输出文件名：").grid(row=1, column=0, sticky="w")
        fnf = ttk.Frame(frm)
        fnf.grid(row=1, column=1, columnspan=2, sticky="we")
        self.var_auto_name = tk.BooleanVar(value=cfg.get("auto_name", True))
        ttk.Checkbutton(fnf, text="自动（卡片【当前日期与时间】）",
                        variable=self.var_auto_name,
                        command=self._toggle_fname).pack(side="left")
        self.var_fname = tk.StringVar(value=cfg.get("custom_name", "我的卡片"))
        self.ent_fname = ttk.Entry(fnf, textvariable=self.var_fname)
        self.ent_fname.pack(side="left", padx=(10, 2), fill="x", expand=True)
        ttk.Label(fnf, text=".docx", foreground="#666666").pack(side="left")

        # ---- 卡片规格 ----
        size_frm = ttk.LabelFrame(frm, text="卡片规格", padding=(6, 3))
        size_frm.grid(row=2, column=0, columnspan=3, sticky="we", pady=4)

        ttk.Label(size_frm, text="尺寸：").grid(row=0, column=0, sticky="w")
        self.var_size = tk.StringVar(value=cfg.get("card_size", "3寸"))
        if self.var_size.get() not in SIZE_ORDER:
            self.var_size.set("3寸")
        self.cmb_size = ttk.Combobox(size_frm, textvariable=self.var_size,
                                     values=SIZE_ORDER, state="readonly", width=5)
        self.cmb_size.grid(row=0, column=1, sticky="w")
        self.cmb_size.bind("<<ComboboxSelected>>", lambda e: self.on_size_change())

        ttk.Label(size_frm, text="方向：").grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.var_orient = tk.StringVar(value=cfg.get("orientation", "竖版"))
        if self.var_orient.get() not in ("竖版", "横版"):
            self.var_orient.set("竖版")
        ttk.Radiobutton(size_frm, text="竖版", variable=self.var_orient, value="竖版",
                        command=self.update_layout_info).grid(row=0, column=3, sticky="w")
        ttk.Radiobutton(size_frm, text="横版", variable=self.var_orient, value="横版",
                        command=self.update_layout_info).grid(row=0, column=4, sticky="w")

        # 自定义宽高
        ttk.Label(size_frm, text="宽(cm)：").grid(row=0, column=5, sticky="e", padx=(8, 0))
        self.var_cw = tk.StringVar(value=str(cfg.get("custom_w", 5.5)))
        self.ent_cw = ttk.Entry(size_frm, textvariable=self.var_cw, width=5)
        self.ent_cw.grid(row=0, column=6)
        ttk.Label(size_frm, text="高(cm)：").grid(row=0, column=7, sticky="e", padx=(6, 0))
        self.var_ch = tk.StringVar(value=str(cfg.get("custom_h", 8.4)))
        self.ent_ch = ttk.Entry(size_frm, textvariable=self.var_ch, width=5)
        self.ent_ch.grid(row=0, column=8)
        for w in (self.var_cw, self.var_ch):
            w.trace_add("write", lambda *a: self.update_layout_info())

        # 每页行列
        ttk.Label(size_frm, text="每页行数：").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.var_auto_layout = tk.BooleanVar(value=cfg.get("auto_layout", True))
        ttk.Checkbutton(size_frm, text="自动", variable=self.var_auto_layout,
                        command=self._toggle_layout).grid(row=1, column=1, sticky="w", pady=(4, 0))
        self.var_rpp = tk.IntVar(value=cfg.get("rows_per_page", 3))
        self.spin_rpp = ttk.Spinbox(size_frm, from_=1, to=20, width=4,
                                    textvariable=self.var_rpp)
        self.spin_rpp.grid(row=1, column=2, sticky="w", pady=(4, 0))
        self.var_spacer = tk.BooleanVar(value=cfg.get("spacer_enabled", False))
        ttk.Checkbutton(size_frm, text="卡片之间加分隔空行（裁剪用）",
                        variable=self.var_spacer,
                        command=self.update_layout_info).grid(
            row=1, column=3, columnspan=3, sticky="w", pady=(4, 0))

        # 页边距（默认"窄"，四边 1.27cm）
        ttk.Label(size_frm, text="页边距：").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.var_margin = tk.StringVar(value=cfg.get("margin_preset", DEFAULT_MARGIN_KEY))
        if self.var_margin.get() not in MARGIN_ORDER:
            self.var_margin.set(DEFAULT_MARGIN_KEY)
        self.cmb_margin = ttk.Combobox(size_frm, textvariable=self.var_margin,
                                       values=MARGIN_ORDER, state="readonly", width=12)
        self.cmb_margin.grid(row=2, column=1, sticky="w", pady=(4, 0))
        self.cmb_margin.bind("<<ComboboxSelected>>", lambda e: self.on_margin_change())
        ttk.Label(size_frm, text="自定义(cm)：").grid(row=2, column=2, sticky="e",
                                                     padx=(8, 0), pady=(4, 0))
        self.var_margin_cm = tk.StringVar(value=str(cfg.get("custom_margin", DEFAULT_MARGIN_CM)))
        self.ent_margin = ttk.Entry(size_frm, textvariable=self.var_margin_cm, width=5)
        self.ent_margin.grid(row=2, column=3, sticky="w", pady=(4, 0))
        ttk.Label(size_frm, text="（四边相同）", foreground="#666666").grid(
            row=2, column=4, sticky="w", padx=(4, 0), pady=(4, 0))
        self.var_margin_cm.trace_add("write", lambda *a: self.update_layout_info())

        # 网格对齐（默认左上角：所有页基准一致）
        ttk.Label(size_frm, text="网格对齐：").grid(row=3, column=0, sticky="w", pady=(4, 0))
        self.var_align = tk.StringVar(value=cfg.get("align_preset", DEFAULT_ALIGN_KEY))
        if self.var_align.get() not in ALIGN_ORDER:
            self.var_align.set(DEFAULT_ALIGN_KEY)
        self.cmb_align = ttk.Combobox(size_frm, textvariable=self.var_align,
                                      values=ALIGN_ORDER, state="readonly", width=20)
        self.cmb_align.grid(row=3, column=1, columnspan=2, sticky="w", pady=(4, 0))
        self.cmb_align.bind("<<ComboboxSelected>>", lambda e: self.update_layout_info())
        ttk.Label(size_frm, text="（不满一页时沿用同一左上角基准，不会下沉）",
                  foreground="#666666").grid(row=3, column=3, columnspan=5, sticky="w",
                                             padx=(8, 0), pady=(4, 0))

        # 规格信息
        self.var_size_info = tk.StringVar(value="")
        ttk.Label(size_frm, textvariable=self.var_size_info, foreground="#0066cc",
                  wraplength=540).grid(row=4, column=0, columnspan=9, sticky="w", pady=(3, 0))
        size_frm.columnconfigure(8, weight=1)

        # ---- 图片设置 ----
        pic_frm = ttk.LabelFrame(frm, text="图片设置", padding=(6, 3))
        pic_frm.grid(row=3, column=0, columnspan=3, sticky="we", pady=(0, 4))
        self.var_with_pic = tk.BooleanVar(value=cfg.get("with_picture", True))
        ttk.Checkbutton(pic_frm, text="每张卡片上方添加图片行",
                        variable=self.var_with_pic,
                        command=self._toggle_pic).grid(row=0, column=0, sticky="w")
        ttk.Label(pic_frm, text="图片文件夹：").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.var_imgdir = tk.StringVar(value=cfg.get("image_folder", ""))
        self.ent_imgdir = ttk.Entry(pic_frm, textvariable=self.var_imgdir)
        self.ent_imgdir.grid(row=1, column=1, sticky="we", padx=4, pady=(4, 0))
        self.btn_imgdir = ttk.Button(pic_frm, text="浏览…", command=self.choose_imgdir)
        self.btn_imgdir.grid(row=1, column=2, pady=(4, 0))
        self.var_imgdir.trace_add("write", lambda *a: self._refresh_preview())

        ttk.Label(pic_frm, text="适配方式：").grid(row=2, column=0, sticky="w")
        self._fit_labels = [
            "拉伸填充（铺满单元格，不裁切，可能变形）",
            "铺满裁剪（不变形，居中裁切多余部分）",
            "完整显示（不变形，可能留白边）",
        ]
        self._fit_values = ["stretch", "cover", "contain"]
        self.var_fit = tk.StringVar(
            value=cfg.get("image_fit", "stretch")
            if cfg.get("image_fit") in ("stretch", "cover", "contain") else "stretch")
        self.cmb_fit = ttk.Combobox(pic_frm, state="readonly", width=26,
                                     values=self._fit_labels)
        self.cmb_fit.current(self._fit_values.index(self.var_fit.get()))
        self.cmb_fit.grid(row=2, column=1, sticky="w")
        self.cmb_fit.bind("<<ComboboxSelected>>", lambda e: self._on_fit_change())

        ttk.Label(pic_frm, text="（留空则生成“【图片】”占位符；图片文件名需与词组一致；"
                                "预览中点击缺图可联网搜图、点击已有图片可打开所在文件夹）",
                  foreground="#666666").grid(row=3, column=0, columnspan=3, sticky="w")
        pic_frm.columnconfigure(1, weight=1)

        # ---- 版式信息 + 文件名预览 ----
        info_frm = ttk.Frame(frm)
        info_frm.grid(row=4, column=0, columnspan=3, sticky="we", pady=(4, 0))
        info_frm.columnconfigure(0, weight=1)
        self.var_info = tk.StringVar(value="")
        ttk.Label(info_frm, textvariable=self.var_info, foreground="#0066cc").grid(
            row=0, column=0, sticky="w")
        self.var_name_preview = tk.StringVar(value="")
        ttk.Label(info_frm, textvariable=self.var_name_preview, foreground="#b05a00").grid(
            row=1, column=0, sticky="w")

        # ---- 词组输入 ----
        ttk.Label(frm, text="粘贴文字（自动按 逗号 / 顿号 / 分号 / 空格 / 换行 分隔词组）：").grid(
            row=5, column=0, columnspan=3, sticky="w", pady=(5, 2))
        self.txt = scrolledtext.ScrolledText(frm, height=8, font=("微软雅黑", 10), wrap="word")
        self.txt.grid(row=6, column=0, columnspan=3, sticky="nsew", padx=(0, 2))

        # ---- 选项 ----
        opt = ttk.Frame(frm)
        opt.grid(row=7, column=0, columnspan=3, sticky="w", pady=4)
        ttk.Label(opt, text="最后一行多余格子：").pack(side="left")
        self.var_keep = tk.BooleanVar(value=cfg.get("keep_tail_placeholder", False))
        ttk.Radiobutton(opt, text="留空", variable=self.var_keep, value=False).pack(side="left")
        ttk.Radiobutton(opt, text="保留“文字”", variable=self.var_keep, value=True).pack(side="left")
        self.var_count = tk.StringVar(value="词组数：0")
        ttk.Label(opt, textvariable=self.var_count, foreground="#888888").pack(side="right")

        # ---- 按钮 ----
        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, columnspan=3, sticky="we", pady=(2, 6))
        ttk.Button(btns, text="生成卡片文档", command=self.generate).pack(side="left")
        ttk.Button(btns, text="打开输出文件夹", command=self.open_outdir).pack(side="left", padx=8)
        ttk.Button(btns, text="清空输入", command=lambda: (self.txt.delete("1.0", "end"),
                                                           self.update_count())).pack(side="left")

        # ---- 状态栏 ----
        self.var_status = tk.StringVar(
            value="就绪。版式完全内置（无需模板文件）：默认 3寸竖版（5.5×8.4cm），"
                  "页边距“窄”四边 1.27cm，网格按满页基准对齐；文字加粗并按单元格"
                  "自动缩放（四周留 2mm）。预览中点击缺图卡片可百度搜图补齐。")
        ttk.Label(frm, textvariable=self.var_status, foreground="#2a7a2a", wraplength=560).grid(
            row=9, column=0, columnspan=3, sticky="w")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(6, weight=1)

        self.txt.bind("<KeyRelease>", lambda e: self.update_count())
        # 右键粘贴 / 菜单粘贴不一定触发 KeyRelease，补两个绑定
        self.txt.bind("<<Paste>>", lambda e: self.root.after(30, self.update_count))
        self.txt.bind("<ButtonRelease-3>",
                      lambda e: self.root.after(30, self.update_count))
        self.refresh_layout_note()
        if cfg.get("sample_text"):
            self.txt.insert("1.0", cfg["sample_text"])
        self._toggle_pic()
        self._toggle_fname()
        self._toggle_layout()
        self._toggle_margin()
        self.on_size_change()
        # 首次预览（等待窗口完成布局后再绘制）
        self.root.after(180, self.preview.draw)


    def show_about(self):
        messagebox.showinfo(
            "关于 %s" % APP_NAME,
            "%s（%s）\n\n"
            "版本：%s\n"
            "作者：%s\n"
            "版权：%s\n\n"
            "%s\n\n"
            "本软件仅使用 Python 标准库编写。" % (
                APP_NAME, APP_NAME_EN, APP_VERSION, APP_AUTHOR,
                APP_COPYRIGHT, APP_DESCRIPTION,
            ),
            parent=self.root,
        )

    # ---------- 预览联动 ----------

    def get_words(self):
        """当前输入的词组列表"""
        try:
            return parse_words(self.txt.get("1.0", "end"))
        except Exception:
            return []

    def _refresh_preview(self):
        """请求右侧预览重绘（有防抖，短时间内多次调用只重绘一次）"""
        p = getattr(self, "preview", None)
        if p is not None:
            p.schedule_refresh()

    # ---------- 规格相关 ----------

    def get_card_size(self):
        """返回 (宽cm, 高cm, 尺寸名)"""
        name = self.var_size.get()
        orient = self.var_orient.get()
        if name == "自定义":
            try:
                w = float(self.var_cw.get())
                h = float(self.var_ch.get())
            except ValueError:
                w, h = CARD_SIZES["3寸"]
            if w <= 0 or h <= 0:
                w, h = CARD_SIZES["3寸"]
        else:
            w, h = CARD_SIZES[name]
        if orient == "横版":
            w, h = h, w
        return w, h, name

    def get_margin_cm(self):
        """返回当前页边距（cm）；预设无效时回退默认 1.27cm"""
        name = self.var_margin.get()
        if name in MARGIN_PRESETS:
            return MARGIN_PRESETS[name]
        try:
            v = float(self.var_margin_cm.get())
            return v if v >= 0 else DEFAULT_MARGIN_CM
        except (ValueError, TypeError):
            return DEFAULT_MARGIN_CM

    def get_align_mode(self):
        """返回当前网格对齐模式（topleft / hcenter / center）"""
        return ALIGN_MODES.get(self.var_align.get(), DEFAULT_ALIGN_MODE)

    def get_layout(self):
        w, h, name = self.get_card_size()
        try:
            rpp = None if self.var_auto_layout.get() else max(1, int(self.var_rpp.get()))
        except Exception:
            rpp = None
        pw, ph = template_page_size()
        cw, ch = content_size(pw, ph, self.get_margin_cm())
        return compute_layout(w, h, cw, ch, rpp, self.var_spacer.get()), w, h, name

    def on_size_change(self):
        custom = (self.var_size.get() == "自定义")
        st = "normal" if custom else "disabled"
        self.ent_cw.config(state=st)
        self.ent_ch.config(state=st)
        self.update_layout_info()

    def on_margin_change(self):
        self._toggle_margin()
        self.update_layout_info()

    def _toggle_margin(self):
        custom = (self.var_margin.get() == "自定义")
        self.ent_margin.config(state="normal" if custom else "disabled")

    def update_layout_info(self):
        try:
            layout, w, h, name = self.get_layout()
        except Exception:
            return
        pouch = POUCH_INFO.get(name, "")
        extra = ""
        if name == "自定义":
            extra = "（自定义尺寸）"
        elif pouch:
            extra = "（%s）" % pouch
        self.var_size_info.set(
            "卡片：宽 %.2f × 高 %.2f cm%s｜A4 每页 %d 列 × %d 行 = %d 张｜"
            "页边距 %.2f cm（四边）｜对齐：%s｜图片区高 %.1f cm、文字区高 %.1f cm"
            "（文字加粗、按单元格自动缩放，四周留 %.0fmm）"
            % (w, h, extra, layout["cols"], layout["rows_per_page"], layout["per_page"],
               self.get_margin_cm(), self.var_align.get(),
               twips_to_cm(layout["pic_h"]), twips_to_cm(layout["word_h"]),
               TEXT_PAD_MM))
        self.update_count()

    def _toggle_layout(self):
        st = "disabled" if self.var_auto_layout.get() else "normal"
        self.spin_rpp.config(state=st)
        self.update_layout_info()

    # ---------- 辅助 ----------

    @staticmethod
    def default_outdir(cfg):
        """默认输出到桌面（没有桌面就用用户主目录）"""
        d = os.path.join(os.path.expanduser("~"), "Desktop")
        return d if os.path.isdir(d) else os.path.expanduser("~")

    def load_config(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "outdir": self.var_outdir.get(),
                    "auto_name": self.var_auto_name.get(),
                    "custom_name": self.var_fname.get(),
                    "card_size": self.var_size.get(),
                    "orientation": self.var_orient.get(),
                    "custom_w": self.var_cw.get(),
                    "custom_h": self.var_ch.get(),
                    "margin_preset": self.var_margin.get(),
                    "custom_margin": self.var_margin_cm.get(),
                    "align_preset": self.var_align.get(),
                    "auto_layout": self.var_auto_layout.get(),
                    "rows_per_page": self.var_rpp.get(),
                    "spacer_enabled": self.var_spacer.get(),
                    "with_picture": self.var_with_pic.get(),
                    "image_folder": self.var_imgdir.get(),
                    "image_fit": self.var_fit.get(),
                    "keep_tail_placeholder": self.var_keep.get(),
                    "sample_text": self.txt.get("1.0", "end").strip()[:2000],
                }, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def choose_outdir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.var_outdir.set(d)

    def choose_imgdir(self):
        d = filedialog.askdirectory(title="选择图片文件夹")
        if d:
            self.var_imgdir.set(d)

    def _toggle_pic(self):
        st = "normal" if self.var_with_pic.get() else "disabled"
        self.ent_imgdir.config(state=st)
        self.btn_imgdir.config(state=st)
        self.cmb_fit.config(state="readonly" if st == "normal" else "disabled")
        self._refresh_preview()

    def _on_fit_change(self):
        idx = self.cmb_fit.current()
        if 0 <= idx < len(self._fit_values):
            self.var_fit.set(self._fit_values[idx])
        self._refresh_preview()

    def _toggle_fname(self):
        st = "disabled" if self.var_auto_name.get() else "normal"
        self.ent_fname.config(state=st)
        self.update_name_preview()

    def update_name_preview(self):
        try:
            custom = None if self.var_auto_name.get() else self.var_fname.get()
            self.var_name_preview.set("生成时将保存为：%s" % build_output_filename(custom))
        except Exception:
            pass

    def get_word_max_pt(self):
        """文字字号上限 = 内置文字行原型的原始字号（只缩不放）"""
        v = getattr(self, "_word_max_pt", None)
        if v:
            return v
        try:
            self._word_max_pt = proto_word_pt(BUILTIN_CARD_ROW)
        except Exception:
            self._word_max_pt = DEFAULT_TEXT_PT
        return self._word_max_pt

    def refresh_layout_note(self):
        """版式说明（版式内置，不再有模板文件）"""
        self._word_max_pt = proto_word_pt(BUILTIN_CARD_ROW)
        cells = len(re.findall(r"<w:tc>.*?</w:tc>", BUILTIN_CARD_ROW, re.S))
        self.var_info.set(
            "内置版式：A4 纸张，文字 %s 磅微软雅黑加粗、按单元格自动缩放（四周留 %.0fmm），"
            "表格 0.5 磅细线边框；尺寸/行列/页边距均自动计算，无需模板文件。"
            % ("%g" % self._word_max_pt, TEXT_PAD_MM))
        self.update_name_preview()

    def update_count(self):
        try:
            layout, w, h, name = self.get_layout()
        except Exception:
            self.var_count.set("词组数：0")
            self._refresh_preview()
            return
        words = self.get_words()
        if not words:
            self.var_count.set("词组数：0")
            self._refresh_preview()
            return
        total_cards = (len(words) + layout["cols"] - 1) // layout["cols"]
        pages = (total_cards + layout["rows_per_page"] - 1) // layout["rows_per_page"]
        self.var_count.set("词组数：%d  →  %d 张卡片，%d 页（每页 %d 列×%d 行）"
                           % (len(words), total_cards, pages,
                              layout["cols"], layout["rows_per_page"]))
        self._refresh_preview()

    def generate(self):
        words = parse_words(self.txt.get("1.0", "end"))
        if not words:
            messagebox.showwarning("提示", "请先粘贴或输入词组文字。")
            return
        out_dir = self.var_outdir.get().strip()
        if not os.path.isdir(out_dir):
            messagebox.showerror("错误", "输出目录不存在：%s" % out_dir)
            return
        img_dir = self.var_imgdir.get().strip()
        if img_dir and not os.path.isdir(img_dir):
            messagebox.showerror("错误", "图片文件夹不存在：%s" % img_dir)
            return
        w, h, _ = self.get_card_size()
        try:
            rpp = None if self.var_auto_layout.get() else max(1, int(self.var_rpp.get()))
        except Exception:
            rpp = None

        custom_name = None if self.var_auto_name.get() else self.var_fname.get()
        out_path = make_output_path(out_dir, custom_name)
        self.update_name_preview()

        ok, msg = fill_template(
            words, out_path,
            card_width_cm=w, card_height_cm=h,
            rows_per_page=rpp,
            keep_tail_placeholder=self.var_keep.get(),
            with_picture=self.var_with_pic.get(),
            image_folder=img_dir,
            spacer_enabled=self.var_spacer.get(),
            fit_mode=self.var_fit.get(),
            margin_cm=self.get_margin_cm(),
            align_mode=self.get_align_mode())
        if ok:
            self.var_status.set(msg.replace("\n", "  "))
            self.save_config()
            if messagebox.askyesno("完成", msg + "\n\n是否立即打开文档所在文件夹？"):
                self.open_outdir()
        else:
            self.var_status.set("失败：" + msg)
            messagebox.showerror("失败", msg)

    def open_outdir(self):
        d = self.var_outdir.get().strip()
        if os.path.isdir(d):
            if sys.platform.startswith("win"):
                os.startfile(d)  # noqa
            elif sys.platform == "darwin":
                subprocess.Popen(["open", d])
            else:
                subprocess.Popen(["xdg-open", d])


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
