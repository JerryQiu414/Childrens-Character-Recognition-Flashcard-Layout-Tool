# -*- coding: utf-8 -*-
"""自动为 卡片批量生成器.py 注入版本/作者/版权信息，避免手工缩进出错。"""
import re

path = r'C:/Users/QZE/WorkBuddy/20260929195933/卡片批量生成器.py'
with open(path, encoding='utf-8') as f:
    src = f.read()

# 1. 修改 docstring 头部，追加版本/作者/版权/免责声明
old_docstring = '''# -*- coding: utf-8 -*-
"""
卡片批量生成器 v4.4（修复JPEG尺寸解析偏移）
======================================
功能：'''
new_docstring = '''# -*- coding: utf-8 -*-
"""
卡片批量生成器 v4.4（修复JPEG尺寸解析偏移）
======================================
产品名称：卡片批量生成器（Card Batch Generator）
版    本：4.4.0.0
作    者：邱卓尔
版权所有：Copyright (c) 2026 邱卓尔. 保留所有权利.
软件说明：本软件用于批量生成图文卡片 Word 文档，支持多种卡片
          规格（3/4/5/6寸、横竖版、自定义）、图片自动匹配与
          三种适配模式（拉伸填充/铺满裁剪/完整显示）。
免责声明：本软件按"现状"提供，不对因使用本软件产生的任何直接
          或间接损失承担责任。

功能：'''
assert old_docstring in src, 'docstring 头部未找到'
src = src.replace(old_docstring, new_docstring, 1)

# 2. 在导入语句后、配置常量前插入 APP 版本/作者/版权常量
const_block = '''# --------------------------- 版本与版权信息 ---------------------------
APP_NAME = "卡片批量生成器"
APP_NAME_EN = "Card Batch Generator"
APP_VERSION = "4.4.0.0"
APP_AUTHOR = "邱卓尔"
APP_COMPANY = "邱卓尔"
APP_COPYRIGHT = "Copyright (c) 2026 邱卓尔. 保留所有权利."
APP_DESCRIPTION = ("批量生成图文卡片 Word 文档：支持 3/4/5/6 寸与自定义规格、"
                   "横竖版、图片自动匹配与拉伸填充/铺满裁剪/完整显示三种适配模式。")

'''
old_config = '''# ----------------------------- 配置 -----------------------------

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "卡片生成器配置.json")'''
new_config = '''# ----------------------------- 配置 -----------------------------

# 打包成 EXE（PyInstaller onedir）时，配置文件保存在 EXE 所在目录，
# 而不是脚本/临时解压目录，否则配置会丢失。
if getattr(sys, "frozen", False):
    _APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_APP_DIR, "卡片生成器配置.json")'''
assert old_config in src, '配置常量未找到'
src = src.replace(old_config, const_block + new_config, 1)

# 3. 修改 root.title
old_title = 'root.title("卡片批量生成器 v4.4（修复JPEG尺寸解析偏移）")'
new_title = 'root.title("%s v%s（修复JPEG尺寸解析偏移）" % (APP_NAME, APP_VERSION))'
assert old_title in src, 'root.title 未找到'
src = src.replace(old_title, new_title, 1)

# 4. 在 __init__ 末尾 root.minsize 之后插入菜单
old_init_tail = '''        root.minsize(740, 660)

        cfg = self.load_config()'''
new_init_tail = '''        root.minsize(740, 660)

        # ---- 菜单栏（帮助 / 关于）----
        menubar = tk.Menu(root)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于 卡片批量生成器", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)
        root.config(menu=menubar)

        cfg = self.load_config()'''
assert old_init_tail in src, '__init__ 初始化尾部未找到'
src = src.replace(old_init_tail, new_init_tail, 1)

# 5. 在 class App 中 load_config 方法之前插入 show_about 方法
about_method = '''
    def show_about(self):
        messagebox.showinfo(
            "关于 %s" % APP_NAME,
            "%s（%s）\\n\\n"
            "版本：%s\\n"
            "作者：%s\\n"
            "版权：%s\\n\\n"
            "%s\\n\\n"
            "本软件仅使用 Python 标准库编写。" % (
                APP_NAME, APP_NAME_EN, APP_VERSION, APP_AUTHOR,
                APP_COPYRIGHT, APP_DESCRIPTION,
            ),
            parent=self.root,
        )

'''
old_load_config = '    def load_config(self):'
marker = '    # ---------- 规格相关 ----------\n\n    def get_card_size(self):'
assert marker in src, '规格相关分隔线未找到'
src = src.replace(marker, about_method + marker, 1)

with open(path, 'w', encoding='utf-8') as f:
    f.write(src)
print('已完成应用信息注入，请执行 py_compile 验证。')
