# -*- coding: utf-8 -*-
"""
案例分析 PPT 生成脚本：赵一鸣零食 —— 战略地图与战略分析（基于平衡计分卡 BSC）
依赖：python-pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ----------------------------- 设计系统 -----------------------------
PRIMARY      = RGBColor(0xC0, 0x1A, 0x22)   # 赵一鸣品牌红
PRIMARY_DK   = RGBColor(0x8C, 0x12, 0x18)
ACCENT       = RGBColor(0xF2, 0x99, 0x00)   # 暖橙
INK          = RGBColor(0x26, 0x26, 0x26)
GRAY         = RGBColor(0x66, 0x66, 0x66)
LIGHT        = RGBColor(0xFC, 0xF4, 0xF1)
LINE_GRAY    = RGBColor(0xDD, 0xD5, 0xD2)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)

# BSC 四维度配色
FIN   = RGBColor(0xC0, 0x1A, 0x22)   # 财务 红
CUS   = RGBColor(0xE3, 0x7E, 0x1B)   # 客户 橙
PROC  = RGBColor(0x2E, 0x8B, 0x76)   # 内部流程 青绿
LEARN = RGBColor(0x2C, 0x5F, 0x8A)   # 学习成长 蓝

FONT = "微软雅黑"
FONT_EN = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ----------------------------- 工具函数 -----------------------------
def slide():
    return prs.slides.add_slide(BLANK)


def _set_font(run, size, color, bold=False, italic=False, font=FONT):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    # 中文字体
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)


def rect(s, l, t, w, h, fill=None, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, l, t, w, h)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    sp.shadow.inherit = False
    return sp


def txt(s, l, t, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        wrap=True):
    """lines: list of dicts {text,size,color,bold,italic,font,space_after,align}"""
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2)
    tf.margin_top = tf.margin_bottom = Pt(2)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ln.get('align', align)
        if ln.get('space_after') is not None:
            p.space_after = Pt(ln['space_after'])
        if ln.get('space_before') is not None:
            p.space_before = Pt(ln['space_before'])
        if ln.get('line_spacing'):
            p.line_spacing = ln['line_spacing']
        segs = ln['text'] if isinstance(ln['text'], list) else [ln]
        for seg in segs:
            r = p.add_run()
            r.text = seg['text']
            _set_font(r, seg.get('size', ln.get('size', 14)),
                      seg.get('color', ln.get('color', INK)),
                      seg.get('bold', ln.get('bold', False)),
                      seg.get('italic', ln.get('italic', False)),
                      seg.get('font', ln.get('font', FONT)))
    return tb


def header(s, no, title_cn, title_en=None):
    """标准内容页页眉"""
    rect(s, 0, 0, SW, Inches(1.05), fill=WHITE)
    rect(s, Inches(0.55), Inches(0.34), Inches(0.14), Inches(0.42), fill=PRIMARY)
    runs = [{'text': title_cn, 'size': 25, 'color': INK, 'bold': True}]
    txt(s, Inches(0.78), Inches(0.28), Inches(9.5), Inches(0.6),
        [{'text': runs, 'size': 25}], anchor=MSO_ANCHOR.MIDDLE)
    if no:
        txt(s, Inches(11.6), Inches(0.2), Inches(1.4), Inches(0.7),
            [{'text': no, 'size': 30, 'color': RGBColor(0xEA, 0xD9, 0xD4), 'bold': True,
              'font': FONT_EN, 'align': PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(0.55), Inches(0.95), Inches(12.2), Pt(2), fill=LINE_GRAY)


def footer(s, idx):
    txt(s, Inches(0.55), Inches(7.05), Inches(8), Inches(0.35),
        [{'text': '案例分析 · 赵一鸣零食战略地图与平衡计分卡', 'size': 9, 'color': GRAY}])
    txt(s, Inches(11.5), Inches(7.05), Inches(1.3), Inches(0.35),
        [{'text': str(idx), 'size': 9, 'color': GRAY, 'align': PP_ALIGN.RIGHT}])


def bg(s, color=WHITE):
    rect(s, 0, 0, SW, SH, fill=color)


PAGE = [0]


def page(s):
    PAGE[0] += 1
    footer(s, PAGE[0])


# =====================================================================
# 1. 封面
# =====================================================================
s = slide()
bg(s, WHITE)
rect(s, 0, 0, SW, SH, fill=PRIMARY)
rect(s, 0, Inches(5.0), SW, Inches(2.5), fill=PRIMARY_DK)
# 装饰条
rect(s, Inches(0.9), Inches(2.05), Inches(0.9), Pt(5), fill=ACCENT)
txt(s, Inches(0.9), Inches(1.0), Inches(11.5), Inches(0.8),
    [{'text': '管理会计 / 战略管理 · 案例分析', 'size': 18, 'color': RGBColor(0xF3, 0xCF, 0xC9),
      'bold': False}])
txt(s, Inches(0.85), Inches(2.35), Inches(11.6), Inches(1.8),
    [{'text': '赵一鸣零食', 'size': 60, 'color': WHITE, 'bold': True}])
txt(s, Inches(0.9), Inches(3.55), Inches(11.6), Inches(1.0),
    [{'text': '战略地图与战略执行分析', 'size': 34, 'color': WHITE, 'bold': True},
     {'text': '——基于平衡计分卡（BSC）的四维度视角', 'size': 18,
      'color': RGBColor(0xF3, 0xCF, 0xC9), 'space_before': 8}])
txt(s, Inches(0.9), Inches(5.45), Inches(11.6), Inches(1.4),
    [{'text': [{'text': '汇报小组：', 'size': 14, 'color': RGBColor(0xF0, 0xC8, 0xC2)},
               {'text': '第 X 组（5 人）', 'size': 14, 'color': WHITE, 'bold': True}]},
     {'text': [{'text': '分析框架：', 'size': 14, 'color': RGBColor(0xF0, 0xC8, 0xC2)},
               {'text': 'PEST · 波特五力 · SWOT · 战略地图 · BSC', 'size': 14, 'color': WHITE}],
      'space_before': 6}])
PAGE[0] = 1  # 封面计为第1页但不显示页脚

# =====================================================================
# 2. 目录
# =====================================================================
s = slide()
bg(s, LIGHT)
rect(s, 0, 0, Inches(4.4), SH, fill=PRIMARY)
rect(s, Inches(0.55), Inches(1.0), Inches(0.9), Pt(5), fill=ACCENT)
txt(s, Inches(0.55), Inches(1.25), Inches(3.4), Inches(2.0),
    [{'text': 'CONTENTS', 'size': 22, 'color': RGBColor(0xF0, 0xC8, 0xC2), 'font': FONT_EN, 'bold': True},
     {'text': '目  录', 'size': 40, 'color': WHITE, 'bold': True, 'space_before': 6}])
items = [
    ("01", "案例介绍", "企业背景 · 发展历程 · 商业模式"),
    ("02", "竞争战略分析", "PEST · 波特五力模型 · SWOT"),
    ("03", "战略类型与战略推导", "总体战略 · 三大战略主轴"),
    ("04", "战略地图（BSC 四维度）", "财务 · 客户 · 内部流程 · 学习成长"),
    ("05", "分析与讨论", "战略执行评价 · 偏差原因分析 · 启示"),
]
y = 1.05
for no, t1, t2 in items:
    txt(s, Inches(4.95), Inches(y), Inches(1.0), Inches(0.9),
        [{'text': no, 'size': 34, 'color': PRIMARY, 'bold': True, 'font': FONT_EN}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(6.0), Inches(y), Inches(6.8), Inches(0.9),
        [{'text': t1, 'size': 21, 'color': INK, 'bold': True},
         {'text': t2, 'size': 12.5, 'color': GRAY, 'space_before': 3}],
        anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(5.0), Inches(y + 1.02), Inches(7.7), Pt(1), fill=LINE_GRAY)
    y += 1.15
page(s)


# =====================================================================
# 章节分隔页
# =====================================================================
def section(no, cn, en, sub):
    s = slide()
    bg(s, PRIMARY)
    rect(s, 0, Inches(2.7), SW, Inches(2.1), fill=PRIMARY_DK)
    txt(s, Inches(0.9), Inches(1.5), Inches(6), Inches(1.5),
        [{'text': no, 'size': 96, 'color': RGBColor(0xE0, 0x6A, 0x60), 'bold': True, 'font': FONT_EN}])
    rect(s, Inches(0.95), Inches(3.05), Inches(0.8), Pt(5), fill=ACCENT)
    txt(s, Inches(0.9), Inches(3.25), Inches(11), Inches(1.3),
        [{'text': cn, 'size': 44, 'color': WHITE, 'bold': True}])
    txt(s, Inches(0.95), Inches(4.4), Inches(11), Inches(0.6),
        [{'text': en, 'size': 16, 'color': RGBColor(0xF0, 0xC8, 0xC2), 'font': FONT_EN}])
    txt(s, Inches(0.95), Inches(5.15), Inches(11), Inches(0.6),
        [{'text': sub, 'size': 15, 'color': RGBColor(0xF7, 0xE3, 0xDF)}])
    page(s)


# =====================================================================
# 01 案例介绍
# =====================================================================
section("01", "案例介绍", "Company Profile", "企业背景 · 创始人故事 · 股权结构 · 商业模式")

# --- 企业简介 ---
s = slide(); bg(s)
header(s, "01", "企业简介：从宜春炒货店到全国零食量贩龙头")
# 左：核心介绍
rect(s, Inches(0.55), Inches(1.35), Inches(6.4), Inches(4.0), fill=LIGHT)
rect(s, Inches(0.55), Inches(1.35), Inches(0.1), Inches(4.0), fill=PRIMARY)
txt(s, Inches(0.85), Inches(1.55), Inches(5.9), Inches(3.7),
    [{'text': '品牌定位', 'size': 15, 'color': PRIMARY, 'bold': True, 'space_after': 4},
     {'text': '诞生于江西宜春的量贩零食连锁品牌，以「不玩套路真便宜，超值超省超好吃」为核心价值主张。',
      'size': 13, 'color': INK, 'space_after': 12, 'line_spacing': 1.25},
     {'text': '规模现状', 'size': 15, 'color': PRIMARY, 'bold': True, 'space_after': 4},
     {'text': '精选 2000+ 种特色零食，月月上新 100+ 新品；全国门店数量已突破 5000 家，'
              '是炙手可热的量贩零食头部品牌。', 'size': 13, 'color': INK, 'space_after': 12,
      'line_spacing': 1.25},
     {'text': '集团归属', 'size': 15, 'color': PRIMARY, 'bold': True, 'space_after': 4},
     {'text': '现为「鸣鸣很忙集团」核心品牌（与「零食很忙」并列），集团 2024 年营收 163.28 亿元，'
              '2025 年已推进上市，赵一鸣为集团第一大收入贡献品牌（约 58.4%）。',
      'size': 13, 'color': INK, 'line_spacing': 1.25}])
# 右：使命/价值观/客群卡片
cards = [
    ("企业使命", "帮消费者省钱 · 帮加盟商赚钱 · 帮员工更值钱", ACCENT),
    ("企业价值观", "分享一鸣 · 信任共鸣 · 感恩和鸣 · 创造争鸣", PROC),
    ("目标客群", "三四线及下沉市场家庭、年轻群体；主打社区便民消费场景", LEARN),
]
y = 1.35
for ti, tx, col in cards:
    rect(s, Inches(7.2), Inches(y), Inches(5.55), Inches(1.2), fill=WHITE, line=LINE_GRAY, line_w=Pt(1))
    rect(s, Inches(7.2), Inches(y), Inches(0.1), Inches(1.2), fill=col)
    txt(s, Inches(7.45), Inches(y + 0.12), Inches(5.2), Inches(1.0),
        [{'text': ti, 'size': 15, 'color': col, 'bold': True, 'space_after': 4},
         {'text': tx, 'size': 13, 'color': INK, 'line_spacing': 1.2}],
        anchor=MSO_ANCHOR.MIDDLE)
    y += 1.4
page(s)

# --- 创始人故事 & 发展历程 ---
s = slide(); bg(s)
header(s, "01", "创始人故事与发展历程")
txt(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(0.9),
    [{'text': [{'text': '创始人赵定（1989年生，江西宜春人）：', 'size': 14, 'color': PRIMARY, 'bold': True},
               {'text': '高中辍学赴沪追逐摄影梦，两度受挫后回乡深耕零食。以儿子「赵一鸣」命名品牌，'
                        '寓意像对待儿子一样经营品牌；以蜜雪冰城为榜样开启「狂飙」扩张，成为量贩零食圈黑马。',
                'size': 14, 'color': INK}], 'line_spacing': 1.3}])
# 时间轴
milestones = [
    ("2008", "返乡创业", "回到宜春从事炒货生意，开出第一家炒货零食店，赚得第一桶金"),
    ("2019", "品牌创立", "首家直营店落地，正式推出「赵一鸣」品牌，开启量贩零食连锁化"),
    ("2023", "并入集团", "与「零食很忙」合并组建鸣鸣很忙集团，毛利率由 5.6%→9.7%"),
    ("2024", "规模登顶", "营收 163.28 亿元，门店超 5000 家，成集团第一大收入品牌"),
    ("2025", "上市扩张", "集团推进上市，借资本与供应链优势加速全国化、深耕北方市场"),
]
n = len(milestones)
x0, x1 = 1.0, 12.4
gap = (x1 - x0) / n
ty = 4.55
rect(s, Inches(x0), Inches(ty), Inches(x1 - x0), Pt(3), fill=PRIMARY)
for i, (yr, t1, t2) in enumerate(milestones):
    cx = x0 + gap * i + gap / 2
    # 节点
    rect(s, Inches(cx - 0.11), Inches(ty - 0.07), Inches(0.22), Inches(0.22),
         fill=ACCENT, shape=MSO_SHAPE.OVAL)
    txt(s, Inches(cx - gap / 2 + 0.1), Inches(ty - 1.0), Inches(gap - 0.2), Inches(0.7),
        [{'text': yr, 'size': 22, 'color': PRIMARY, 'bold': True, 'font': FONT_EN,
          'align': PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.BOTTOM)
    rect(s, Inches(cx - gap / 2 + 0.15), Inches(ty + 0.45), Inches(gap - 0.3), Inches(1.85),
         fill=LIGHT)
    txt(s, Inches(cx - gap / 2 + 0.25), Inches(ty + 0.55), Inches(gap - 0.5), Inches(1.7),
        [{'text': t1, 'size': 14, 'color': INK, 'bold': True, 'align': PP_ALIGN.CENTER,
          'space_after': 5},
         {'text': t2, 'size': 11, 'color': GRAY, 'align': PP_ALIGN.CENTER, 'line_spacing': 1.15}])
page(s)

# --- 商业模式 / 商业画布要点 ---
s = slide(); bg(s)
header(s, "01", "商业模式：量贩零食「多、快、好、省」")
boxes = [
    ("多", "SKU 丰富", "2000+ 种零食，月上新 100+；合作定制 SKU 占比超 34%", CUS),
    ("快", "履约高效", "56 个区域仓配中心，核心经济圈 300km 内 24 小时达", PROC),
    ("好", "体验升级", "欢乐舒适的购物体验 + 集团统一 IP 营销（零食音乐节等）", LEARN),
    ("省", "极致质价比", "源头直采去经销商化（92.3%），砍掉中间溢价让利消费者", PRIMARY),
]
bw = 2.95; gap2 = 0.13; x = 0.55; y = 1.55
for letter, t1, t2, col in boxes:
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(3.6), fill=WHITE, line=LINE_GRAY, line_w=Pt(1))
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(1.3), fill=col)
    txt(s, Inches(x), Inches(y + 0.12), Inches(bw), Inches(1.1),
        [{'text': letter, 'size': 44, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(x + 0.2), Inches(y + 1.55), Inches(bw - 0.4), Inches(1.9),
        [{'text': t1, 'size': 18, 'color': col, 'bold': True, 'align': PP_ALIGN.CENTER,
          'space_after': 10},
         {'text': t2, 'size': 13, 'color': INK, 'align': PP_ALIGN.CENTER, 'line_spacing': 1.3}])
    x += bw + gap2
rect(s, Inches(0.55), Inches(5.5), Inches(12.2), Inches(1.1), fill=LIGHT)
rect(s, Inches(0.55), Inches(5.5), Inches(0.1), Inches(1.1), fill=ACCENT)
txt(s, Inches(0.85), Inches(5.62), Inches(11.8), Inches(0.9),
    [{'text': [{'text': '盈利逻辑：', 'size': 14, 'color': PRIMARY, 'bold': True},
               {'text': '加盟连锁 + 供应链直采。公司向加盟商供货收入（约 18 万/月目标口径）→ 门店 GMV '
                        '（约 35–40 万/月）；公司层面以「薄毛利、大规模、高周转」取胜，'
                        '加盟商端毛利 18–20%，自有品牌毛利可达 30–50%。',
                'size': 13, 'color': INK}], 'line_spacing': 1.3}],
    anchor=MSO_ANCHOR.MIDDLE)
page(s)


# =====================================================================
# 02 竞争战略分析
# =====================================================================
section("02", "竞争战略分析", "Competitive Strategy Analysis",
        "PEST 宏观环境 · 波特五力模型 · SWOT 分析")

# --- PEST ---
s = slide(); bg(s)
header(s, "02", "PEST 宏观环境分析")
pest = [
    ("P", "政治 Political", PRIMARY, [
        "契合「县域商业体系建设」与「乡村振兴」战略，下沉市场政策支持力度大",
        "需在反垄断申报、特许经营合规、食品安全属地监管上与集团保持统一高标准"]),
    ("E", "经济 Economic", CUS, [
        "消费分级，大众追求「质价比」，价格定位精准吸引价格敏感型客群",
        "并入集团后毛利率 5.6%→9.7%，盈利能力增强，为独立扩张提供财务基础"]),
    ("S", "社会 Social", PROC, [
        "年轻客群与家庭客群追求即时性、娱乐化零食消费，「欢乐舒适」体验受青睐",
        "华中、华北原优势区域消费者认知深厚，集团统一 IP 营销提升全国知名度"]),
    ("T", "技术 Technological", LEARN, [
        "直接复用集团全流程数字化能力（选品/采购/仓储/门店管理），无需重复建设",
        "库存周转 11.6 天、物流 24 小时达，行业领先，支撑大规模精细化运营"]),
]
bw = 6.0; bh = 2.45; xs = [0.55, 6.78]; ys = [1.35, 3.95]
for i, (L, name, col, pts) in enumerate(pest):
    x = xs[i % 2]; y = ys[i // 2]
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(bh), fill=WHITE, line=LINE_GRAY, line_w=Pt(1))
    rect(s, Inches(x), Inches(y), Inches(1.05), Inches(bh), fill=col)
    txt(s, Inches(x), Inches(y), Inches(1.05), Inches(bh),
        [{'text': L, 'size': 40, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER,
          'font': FONT_EN}], anchor=MSO_ANCHOR.MIDDLE)
    lines = [{'text': name, 'size': 16, 'color': col, 'bold': True, 'space_after': 6}]
    for p in pts:
        lines.append({'text': [{'text': '▪ ', 'size': 12, 'color': col},
                               {'text': p, 'size': 12, 'color': INK}],
                      'space_after': 5, 'line_spacing': 1.2})
    txt(s, Inches(x + 1.25), Inches(y + 0.15), Inches(bw - 1.45), Inches(bh - 0.3), lines)
page(s)

# --- 波特五力 ---
s = slide(); bg(s)
header(s, "02", "波特五力模型：行业竞争结构")
cx, cy = 6.55, 4.05
# 中心：行业内竞争
rect(s, Inches(cx - 1.7), Inches(cy - 0.7), Inches(3.4), Inches(1.4), fill=PRIMARY)
txt(s, Inches(cx - 1.7), Inches(cy - 0.7), Inches(3.4), Inches(1.4),
    [{'text': '现有竞争者', 'size': 18, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER},
     {'text': '【极强】', 'size': 13, 'color': RGBColor(0xFF, 0xE0, 0x80), 'bold': True,
      'align': PP_ALIGN.CENTER, 'space_before': 2},
     {'text': '零食很忙·好想来·来优品；价格战 + 门店加密 + 加盟商争夺',
      'size': 10.5, 'color': RGBColor(0xF7, 0xDD, 0xDA), 'align': PP_ALIGN.CENTER,
      'space_before': 2, 'line_spacing': 1.1}], anchor=MSO_ANCHOR.MIDDLE)
forces = [
    (cx - 1.7, 1.35, 3.4, 1.15, "潜在进入者", "中", PROC,
     "赛道进入整合期，新品牌独立突围难；威胁来自上游制造商向下整合"),
    (cx - 1.7, 5.55, 3.4, 1.15, "替代品", "中", LEARN,
     "传统商超/便利店/社区团购/电商替代；「多快好省」锁定年轻客群"),
    (0.55, cy - 0.95, 4.0, 1.9, "供应商议价力", "弱", CUS,
     "集团第二大品牌 + 联合采购规模 → 极强议价力；定制化产品深度绑定供应商"),
    (8.78, cy - 0.95, 4.0, 1.9, "购买者议价力", "中/弱", ACCENT,
     "消费者价格敏感但忠诚度建立中；加盟商高度依赖品牌、单体无议价力"),
]
for x, y, w, h, name, lv, col, desc in forces:
    rect(s, Inches(x), Inches(y), Inches(w), Inches(h), fill=WHITE, line=col, line_w=Pt(2))
    rect(s, Inches(x), Inches(y), Inches(w), Inches(0.42), fill=col)
    txt(s, Inches(x), Inches(y), Inches(w), Inches(0.42),
        [{'text': [{'text': name + '  ', 'size': 14, 'color': WHITE, 'bold': True},
                   {'text': '【' + lv + '】', 'size': 12, 'color': RGBColor(0xFF, 0xF0, 0xC0), 'bold': True}],
          'align': PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(x + 0.15), Inches(y + 0.5), Inches(w - 0.3), Inches(h - 0.55),
        [{'text': desc, 'size': 11, 'color': INK, 'align': PP_ALIGN.CENTER, 'line_spacing': 1.2}],
        anchor=MSO_ANCHOR.MIDDLE)
# 箭头
def arrow(s, x1, y1, x2, y2):
    cn = s.shapes.add_connector(2, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cn.line.color.rgb = GRAY; cn.line.width = Pt(1.5)
    le = cn.line._get_or_add_ln()
    tail = le.makeelement(qn('a:tailEnd'), {'type': 'triangle'})
    le.append(tail)
arrow(s, cx, 2.5, cx, 3.32)
arrow(s, cx, 4.78, cx, 5.5)
arrow(s, 4.55, cy, 4.85, cy)
arrow(s, 8.45, cy, 8.78, cy)
page(s)

# --- SWOT ---
s = slide(); bg(s)
header(s, "02", "SWOT 分析")
swot = [
    ("S 优势", PROC, [
        "集团赋能：供应链/数字化/资本/品牌营销，增速远超独立发展",
        "独立品牌认知：华北、江西消费者基础与加盟网络深厚",
        "盈利跃升：毛利率 5.6%→9.7%，单店模型优化",
        "地域互补：与零食很忙形成南北错位，减少内耗"]),
    ("W 劣势", CUS, [
        "全国品牌力弱于零食很忙（如湖南大本营影响力较弱）",
        "历史管理遗留：早期「非独立加盟商」反映加盟体系不规范",
        "战略/供应链/数字化完全依赖集团，自主权有限"]),
    ("O 机会", LEARN, [
        "全国化扩张：借上市资金向华东/华南/西南空白市场拓展",
        "承接集团资源倾斜（占集团收入 58.4%）",
        "北方下沉市场竞争较缓，可快速加密构筑护城河"]),
    ("T 威胁", PRIMARY, [
        "与零食很忙未来在同区域直接竞争不可避免",
        "加盟商忠诚度考验：盈利不及预期或倒戈竞品",
        "门店庞大且以加盟为主，单店食安事件易波及品牌声誉"]),
]
bw = 6.0; bh = 2.45; xs = [0.55, 6.78]; ys = [1.35, 3.95]
for i, (name, col, pts) in enumerate(swot):
    x = xs[i % 2]; y = ys[i // 2]
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(bh), fill=WHITE, line=LINE_GRAY, line_w=Pt(1))
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(0.5), fill=col)
    txt(s, Inches(x + 0.2), Inches(y), Inches(bw - 0.4), Inches(0.5),
        [{'text': name, 'size': 16, 'color': WHITE, 'bold': True}], anchor=MSO_ANCHOR.MIDDLE)
    lines = []
    for j, p in enumerate(pts):
        lines.append({'text': [{'text': '• ', 'size': 11.5, 'color': col, 'bold': True},
                               {'text': p, 'size': 11.5, 'color': INK}],
                      'space_after': 4, 'line_spacing': 1.15})
    txt(s, Inches(x + 0.25), Inches(y + 0.62), Inches(bw - 0.5), Inches(bh - 0.7), lines)
page(s)


# =====================================================================
# 03 战略类型与推导
# =====================================================================
section("03", "战略类型与战略推导", "Strategy Formulation",
        "竞争战略定位 · 总体战略 · 三大战略主轴")

# --- 战略类型 ---
s = slide(); bg(s)
header(s, "03", "战略类型：成本领先为主的总成本领先战略")
txt(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(0.8),
    [{'text': [{'text': '基于波特竞争战略框架，', 'size': 14, 'color': INK},
               {'text': '赵一鸣采取「总成本领先战略」', 'size': 14, 'color': PRIMARY, 'bold': True},
               {'text': '——以源头直采、规模采购、高效供应链与数字化运营压低成本，'
                        '向价格敏感型下沉客群提供「质价比」零食；并辅以适度差异化（自有品牌、定制爆款）。',
                'size': 14, 'color': INK}], 'line_spacing': 1.3}])
cols = [
    ("成本领先抓手", PRIMARY, ["源头直采去经销商化（92.3%）", "联合集团规模采购压低进价",
                          "56 个仓配中心 + 物流成本 1.7%", "数字化选品/补货降低损耗"]),
    ("差异化补充", CUS, ["自有品牌毛利 30–50%", "定制爆款（麻酱味素毛肚 1 亿+件）",
                      "统一 IP 营销塑造品牌", "欢乐舒适门店体验"]),
    ("目标聚焦", LEARN, ["聚焦三四线及下沉市场", "深耕北方华北/西北空白市场",
                     "社区便民消费场景", "家庭与年轻客群"]),
]
x = 0.55; bw = 4.0; gap2 = 0.1
for name, col, pts in cols:
    rect(s, Inches(x), Inches(2.35), Inches(bw), Inches(3.9), fill=WHITE, line=LINE_GRAY, line_w=Pt(1))
    rect(s, Inches(x), Inches(2.35), Inches(bw), Inches(0.7), fill=col)
    txt(s, Inches(x), Inches(2.35), Inches(bw), Inches(0.7),
        [{'text': name, 'size': 18, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE)
    lines = []
    for p in pts:
        lines.append({'text': [{'text': '✓ ', 'size': 13, 'color': col, 'bold': True},
                               {'text': p, 'size': 13, 'color': INK}],
                      'space_after': 11, 'line_spacing': 1.2})
    txt(s, Inches(x + 0.3), Inches(2.45 + 0.7), Inches(bw - 0.6), Inches(3.0), lines,
        anchor=MSO_ANCHOR.MIDDLE)
    x += bw + gap2
page(s)

# --- 总体战略 & 三主轴 ---
s = slide(); bg(s)
header(s, "03", "总体战略与三大战略主轴")
rect(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(1.35), fill=PRIMARY)
txt(s, Inches(0.8), Inches(1.42), Inches(11.7), Inches(1.15),
    [{'text': '总体战略', 'size': 13, 'color': RGBColor(0xF7, 0xD0, 0xCB), 'bold': True, 'align': PP_ALIGN.CENTER},
     {'text': '规模深耕 + 品牌提质 + 供应链提效 + 数字化赋能', 'size': 24, 'color': WHITE, 'bold': True,
      'align': PP_ALIGN.CENTER, 'space_before': 2},
     {'text': '打造全国领先的零食量贩龙头品牌 ｜ 2027 目标：营收突破 250 亿元',
      'size': 14, 'color': RGBColor(0xF7, 0xE3, 0xDF), 'align': PP_ALIGN.CENTER, 'space_before': 4}])
axes = [
    ("巩固核心", PRIMARY, "稳固基本盘",
     "持续强化供应链整合优势，牢牢占据下沉市场领导地位，稳固量贩零食基本盘，夯实品牌护城河。"),
    ("精细运营", PROC, "效率驱动转型",
     "战略重心从门店扩张转向单店效益提升，以数字化加强加盟商赋能与管控，从「规模驱动」转向「效率驱动」。"),
    ("多元探索", LEARN, "培育第二曲线",
     "布局健康化、功能化产品赛道，孵化「新鲜零食」等创新业态，挖掘消费新需求，保障长期生命力。"),
]
x = 0.55; bw = 4.0; gap2 = 0.1; y = 2.95
for name, col, tag, desc in axes:
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(3.4), fill=LIGHT, line=col, line_w=Pt(1.5))
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(0.95), fill=col)
    txt(s, Inches(x), Inches(y + 0.08), Inches(bw), Inches(0.85),
        [{'text': name, 'size': 22, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER},
         {'text': tag, 'size': 12, 'color': RGBColor(0xFF, 0xF0, 0xE0), 'align': PP_ALIGN.CENTER,
          'space_before': 2}], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(x + 0.3), Inches(y + 1.15), Inches(bw - 0.6), Inches(2.1),
        [{'text': desc, 'size': 14, 'color': INK, 'line_spacing': 1.35}], anchor=MSO_ANCHOR.MIDDLE)
    x += bw + gap2
txt(s, Inches(0.55), Inches(6.55), Inches(12.2), Inches(0.4),
    [{'text': '战略愿景：通过「核心稳固、运营提效、业态创新」三位一体，构建全链路、可持续的商业增长闭环。',
      'size': 13, 'color': PRIMARY, 'bold': True, 'align': PP_ALIGN.CENTER}])
page(s)


# =====================================================================
# 04 战略地图（核心页）
# =====================================================================
section("04", "战略地图（BSC 四维度）", "Strategy Map",
        "财务 · 客户 · 内部流程 · 学习与成长 —— 因果驱动闭环")

# --- 战略地图主图 ---
s = slide(); bg(s, LIGHT)
header(s, "04", "赵一鸣零食 BSC 战略地图")
dims = [
    ("财务维度", FIN, "扩大营收规模、优化利润结构、提升整体盈利能力",
     ["单店月均供货收入→18万元", "物流成本→<1.5%", "源头直采占比≥95%", "线上营收占比 5%→15%"]),
    ("客户维度", CUS, "稳固客群、拓展市场、提升口碑与复购",
     ["优化品牌形象", "降低客诉率", "会员消费占比→70%", "北方门店占比→35%"]),
    ("内部流程维度", PROC, "完善供应链、品控、门店管理，提升运营效率",
     ["搭建区域仓储配送中心", "商品动销率→88%", "自有品牌占比→30%", "加盟管理标准化"]),
    ("学习与成长维度", LEARN, "强化数字化、人才、组织能力，为长期发展打底",
     ["搭建一体化数字系统", "员工/加盟商培训", "库存准确率→99%", "优化组织架构适配全国扩张"]),
]
top = 1.25; band_h = 1.32; gap_b = 0.08
left = 0.55; band_w = 12.2
label_w = 2.55
for i, (name, col, theme, kpis) in enumerate(dims):
    y = top + i * (band_h + gap_b)
    # 左侧维度标签
    rect(s, Inches(left), Inches(y), Inches(label_w), Inches(band_h), fill=col)
    txt(s, Inches(left + 0.1), Inches(y), Inches(label_w - 0.2), Inches(band_h),
        [{'text': name, 'size': 18, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER},
         {'text': theme, 'size': 10, 'color': RGBColor(0xFF, 0xF0, 0xEC), 'align': PP_ALIGN.CENTER,
          'space_before': 4, 'line_spacing': 1.1}], anchor=MSO_ANCHOR.MIDDLE)
    # 右侧 KPI 卡片
    cx = left + label_w + 0.12
    cw = (band_w - label_w - 0.12 - 0.36) / 4
    for j, k in enumerate(kpis):
        kx = cx + j * (cw + 0.12)
        rect(s, Inches(kx), Inches(y + 0.12), Inches(cw), Inches(band_h - 0.24),
             fill=WHITE, line=col, line_w=Pt(1.25))
        txt(s, Inches(kx + 0.06), Inches(y + 0.12), Inches(cw - 0.12), Inches(band_h - 0.24),
            [{'text': k, 'size': 11.5, 'color': INK, 'align': PP_ALIGN.CENTER, 'line_spacing': 1.1}],
            anchor=MSO_ANCHOR.MIDDLE)
# 右侧因果箭头（自下而上驱动）
ax = left + band_w + 0.05
arr = s.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(12.78), Inches(top),
                         Inches(0.42), Inches(band_h * 4 + gap_b * 3))
arr.fill.solid(); arr.fill.fore_color.rgb = ACCENT; arr.line.fill.background()
arr.shadow.inherit = False
tf = arr.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "因\n果\n驱\n动"
_set_font(r, 11, WHITE, True)
txt(s, Inches(0.55), Inches(6.75), Inches(12.2), Inches(0.4),
    [{'text': '驱动逻辑：学习成长 → 内部流程 → 客户 → 财务，底层能力支撑上层目标，形成战略闭环。',
      'size': 12, 'color': GRAY, 'bold': True, 'align': PP_ALIGN.CENTER}])
page(s)


# --- 四维度详解（每页一个/合并）---
def dim_detail(name, col, theme, rows, note):
    s = slide(); bg(s)
    header(s, "04", f"战略地图 · {name}")
    rect(s, Inches(0.55), Inches(1.3), Inches(12.2), Inches(0.75), fill=col)
    txt(s, Inches(0.85), Inches(1.3), Inches(11.6), Inches(0.75),
        [{'text': [{'text': '战略主题：', 'size': 14, 'color': RGBColor(0xFF, 0xEC, 0xE6), 'bold': True},
                   {'text': theme, 'size': 15, 'color': WHITE, 'bold': True}]}],
        anchor=MSO_ANCHOR.MIDDLE)
    # 表头
    ty = 2.25
    headers = ["战略目标", "关键指标 (KPI)", "目标值 / 现状", "支撑举措"]
    widths = [2.6, 3.0, 2.7, 3.9]
    x = 0.55
    rect(s, Inches(0.55), Inches(ty), Inches(12.2), Inches(0.5), fill=PRIMARY_DK)
    for hh, w in zip(headers, widths):
        txt(s, Inches(x + 0.1), Inches(ty), Inches(w - 0.2), Inches(0.5),
            [{'text': hh, 'size': 13, 'color': WHITE, 'bold': True}], anchor=MSO_ANCHOR.MIDDLE)
        x += w
    ry = ty + 0.5
    rh = 0.86
    for ri, row in enumerate(rows):
        bgc = WHITE if ri % 2 == 0 else LIGHT
        rect(s, Inches(0.55), Inches(ry), Inches(12.2), Inches(rh), fill=bgc, line=LINE_GRAY, line_w=Pt(0.5))
        x = 0.55
        for ci, (cell, w) in enumerate(zip(row, widths)):
            bold = ci == 0
            color = col if ci == 0 else INK
            txt(s, Inches(x + 0.12), Inches(ry), Inches(w - 0.24), Inches(rh),
                [{'text': cell, 'size': 11.5, 'color': color, 'bold': bold, 'line_spacing': 1.12}],
                anchor=MSO_ANCHOR.MIDDLE)
            x += w
        ry += rh
    rect(s, Inches(0.55), Inches(ry + 0.1), Inches(12.2), Inches(0.7), fill=LIGHT)
    rect(s, Inches(0.55), Inches(ry + 0.1), Inches(0.1), Inches(0.7), fill=ACCENT)
    txt(s, Inches(0.8), Inches(ry + 0.1), Inches(11.8), Inches(0.7),
        [{'text': [{'text': '小结：', 'size': 12, 'color': PRIMARY, 'bold': True},
                   {'text': note, 'size': 12, 'color': INK}], 'line_spacing': 1.2}],
        anchor=MSO_ANCHOR.MIDDLE)
    page(s)


dim_detail("财务维度", FIN, "扩大营收规模、优化利润结构、提升整体盈利能力", [
    ["营收规模增长", "总营收 / 单店月供货收入", "250亿(2027) / 18万元", "20% CAGR；规模化扩张 + 单店提效"],
    ["供应链成本优化", "源头直采占比", "≥95%（现 92.3%）", "完全去经销商化，砍中间溢价"],
    ["物流费用管控", "物流成本占收入比", "<1.5%（现 1.7%）", "56 仓配中心 + 数字化网格配送"],
    ["渠道结构升级", "线上(即时零售)营收占比", "5%→15%", "门店变前置仓，接入美团/饿了么"],
], "财务目标以「营收+毛利」双增长为核心；直采与物流成本已具领先优势，毛利提升关键在自有品牌放量。")

dim_detail("客户维度", CUS, "稳固客群、拓展市场、提升用户口碑与复购", [
    ["提升品牌口碑", "品牌形象 / 客诉率", "优化形象、客诉率下降", "统一 IP 营销 + 服务标准化"],
    ["深挖用户价值", "会员消费占比", "→70%（复购率已达 75%）", "会员分层权益 + 精准营销"],
    ["拓展区域市场", "北方门店占比", "→35%（现约 15–18%）", "10 亿专项资金深耕华北/西北"],
    ["扩大客群基数", "门店数量", "北方 3000+ 家", "优先安徽/河北/山东，辐射山陕"],
], "复购率已超额达成（75%>50%），重心转向「从复购到高价值消费」；北方市场为最大增量空间。")

dim_detail("内部流程维度", PROC, "完善供应链、品控、门店管理，提升运营效率", [
    ["供应链网络建设", "区域仓配中心数量", "已建 56 个，持续加密", "网格化配送，300km/24h 达"],
    ["商品周转效率", "商品动销率", "→88%（现约 75%）", "数字化选品 + 全链路品效优化"],
    ["产品力升级", "自有品牌占比", "→30%（现约 12%）", "联合定制 + 打造 10 个亿元大单品"],
    ["门店运营标准化", "加盟管理标准化率", "标准化全覆盖", "WMS/TMS/ERP 一体化管控"],
], "动销率与自有品牌是利润与差异化关键；需从「结果导向」转向「能力导向」（选品周期、库存周转）。")

dim_detail("学习与成长维度", LEARN, "强化数字化、人才、组织能力，为长期发展打底", [
    ["数字化底座建设", "数字化投入 / 系统一体化", "累计 5 亿+；2026 一体化", "WMS/TMS/ERP 打通全链路数据"],
    ["库存精准管理", "库存准确率", "→99%（目标 99.8%）", "智能补货 + 数据协同闭环"],
    ["加盟商赋能", "培训覆盖率 / 新店存活率", "覆盖 100% / 存活率 95%+", "标准化课程 + 实战演练"],
    ["组织能力适配", "组织架构 / 人才梯队", "适配全国扩张", "数据分析/算法/业务复合人才"],
], "组织能力建设需匹配业务扩张速度；当前数字化偏「系统建设」，应深化为「流程再造」与「组织转型」。")


# =====================================================================
# 05 分析与讨论
# =====================================================================
section("05", "分析与讨论", "Evaluation & Discussion",
        "战略执行评价 · 偏差原因分析 · 数据口径修正 · 结论启示")

# --- 战略执行评价（达成 vs 目标）---
s = slide(); bg(s)
header(s, "05", "战略执行评价：目标达成对照")
ty = 1.35
headers = ["维度", "关键指标", "目标值", "现状(最新)", "达成评价"]
widths = [1.7, 3.3, 2.0, 2.0, 3.2]
rect(s, Inches(0.55), Inches(ty), Inches(12.2), Inches(0.5), fill=PRIMARY)
x = 0.55
for hh, w in zip(headers, widths):
    txt(s, Inches(x + 0.1), Inches(ty), Inches(w - 0.2), Inches(0.5),
        [{'text': hh, 'size': 13, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER if hh != '关键指标' else PP_ALIGN.LEFT}],
        anchor=MSO_ANCHOR.MIDDLE)
    x += w
rows = [
    ("财务", "源头直采占比", "≥95% / 90%", "92.3%", "超额*", PROC),
    ("财务", "物流成本占比", "<1.5%", "1.7%", "接近", CUS),
    ("财务", "毛利率(公司层面)", "32%", "7.6%", "差距大", PRIMARY),
    ("客户", "会员复购率", "50%→提升", "75%", "超额达成", PROC),
    ("客户", "北方门店占比", "35%", "约15–18%", "推进中", CUS),
    ("内部流程", "商品动销率", "88%", "约75%", "推进中", CUS),
    ("内部流程", "自有品牌占比", "30%", "约12%(定制34%)", "差距较大", PRIMARY),
    ("学习成长", "库存准确率", "99%", "目标99.8%", "建设中", CUS),
]
ry = ty + 0.5
rh = 0.55
for ri, (d, k, tg, cur, ev, evc) in enumerate(rows):
    bgc = WHITE if ri % 2 == 0 else LIGHT
    rect(s, Inches(0.55), Inches(ry), Inches(12.2), Inches(rh), fill=bgc, line=LINE_GRAY, line_w=Pt(0.5))
    vals = [d, k, tg, cur]
    x = 0.55
    for ci, (cell, w) in enumerate(zip(vals, widths[:4])):
        al = PP_ALIGN.CENTER if ci != 1 else PP_ALIGN.LEFT
        txt(s, Inches(x + 0.1), Inches(ry), Inches(w - 0.2), Inches(rh),
            [{'text': cell, 'size': 11.5, 'color': INK, 'bold': ci == 0, 'align': al}],
            anchor=MSO_ANCHOR.MIDDLE)
        x += w
    # 评价 chip
    chip_w = 1.6
    rect(s, Inches(x + 0.1), Inches(ry + 0.1), Inches(chip_w), Inches(rh - 0.2), fill=evc)
    txt(s, Inches(x + 0.1), Inches(ry + 0.1), Inches(chip_w), Inches(rh - 0.2),
        [{'text': ev, 'size': 11, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE)
    ry += rh
txt(s, Inches(0.55), Inches(ry + 0.08), Inches(12.2), Inches(0.5),
    [{'text': '*注：直采占比、会员复购率已超额完成原定目标；毛利率、自有品牌占比、动销率等仍存在较大差距，'
              '部分源于目标设定偏激进（详见下页偏差原因分析）。', 'size': 10.5, 'color': GRAY,
      'line_spacing': 1.2}])
page(s)

# --- 偏差原因分析 ---
s = slide(); bg(s)
header(s, "05", "偏差原因分析：四维度诊断")
dev = [
    ("财务维度", FIN, "毛利率目标(32%)过高",
     "目标过度依赖行业对标（良品铺子26%、三只松鼠24%）而非自身能力与历史趋势；"
     "公司层面实际仅7.6%；物流成本目标未考虑全国扩张带来的边际压力。"),
    ("客户维度", CUS, "缺乏客户分层管理",
     "「复购→高价值消费」转化路径不清；KPI 偏后验结果（占比）缺过程指标（新客转化、客单价）；"
     "10 亿专项资金偏开店与品牌，未匹配客户洞察与区域产品定制。"),
    ("内部流程", PROC, "能力未拆解、数字化落地滞后",
     "动销率(88%)、自有品牌(30%)目标偏「结果导向」而非「能力导向」；自有品牌缺与供应链/研发/门店协同机制；"
     "数字化仍以功能模块为主，缺流程再造（智能补货、区域差异化选品未落地）。"),
    ("学习成长", LEARN, "组织能力滞后于业务扩张",
     "加盟商培训缺覆盖率、培训后绩效等关键指标；数字化投入偏系统建设而非组织转型；"
     "人才发展与战略目标映射模糊（如自有品牌开发能力对应岗位/技能未明确）。"),
]
bw = 6.0; bh = 2.45; xs = [0.55, 6.78]; ys = [1.35, 3.95]
for i, (name, col, issue, reason) in enumerate(dev):
    x = xs[i % 2]; y = ys[i // 2]
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(bh), fill=WHITE, line=LINE_GRAY, line_w=Pt(1))
    rect(s, Inches(x), Inches(y), Inches(bw), Inches(0.5), fill=col)
    txt(s, Inches(x + 0.2), Inches(y), Inches(bw - 0.4), Inches(0.5),
        [{'text': [{'text': name + '  ｜  ', 'size': 14, 'color': WHITE, 'bold': True},
                   {'text': issue, 'size': 13, 'color': RGBColor(0xFF, 0xEC, 0xE6), 'bold': True}]}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(x + 0.25), Inches(y + 0.62), Inches(bw - 0.5), Inches(bh - 0.7),
        [{'text': reason, 'size': 12, 'color': INK, 'line_spacing': 1.3}])
page(s)

# --- 数据口径差异与修正建议 ---
s = slide(); bg(s)
header(s, "05", "数据口径差异说明与修正建议")
ty = 1.4
headers = ["指标", "原文档数据", "实际/最新数据", "修正建议"]
widths = [2.6, 3.0, 3.0, 3.6]
rect(s, Inches(0.55), Inches(ty), Inches(12.2), Inches(0.5), fill=PRIMARY)
x = 0.55
for hh, w in zip(headers, widths):
    txt(s, Inches(x + 0.1), Inches(ty), Inches(w - 0.2), Inches(0.5),
        [{'text': hh, 'size': 13, 'color': WHITE, 'bold': True}], anchor=MSO_ANCHOR.MIDDLE)
    x += w
rows = [
    ("单店月均营收", "12万元(供货口径)", "25–30万元(门店GMV)", "明确区分「公司供货收入」与「门店 GMV」两口径"),
    ("毛利率核算", "28%(行业/自有品牌)", "公司7.6% / 加盟商18–20%", "分层披露公司层面与加盟商层面毛利率"),
    ("会员年复购率", "30%(2023单品牌)", "75%(2024集团合并后)", "更新为合并后最新 75%，体现会员整合优势"),
    ("供应链直采占比", "80%(2023历史)", "92.3%(2025最新)", "采用最新 92.3%，突出供应链效率持续优化"),
]
ry = ty + 0.5; rh = 0.92
for ri, row in enumerate(rows):
    bgc = WHITE if ri % 2 == 0 else LIGHT
    rect(s, Inches(0.55), Inches(ry), Inches(12.2), Inches(rh), fill=bgc, line=LINE_GRAY, line_w=Pt(0.5))
    x = 0.55
    for ci, (cell, w) in enumerate(zip(row, widths)):
        color = PRIMARY if ci == 0 else INK
        txt(s, Inches(x + 0.12), Inches(ry), Inches(w - 0.24), Inches(rh),
            [{'text': cell, 'size': 12, 'color': color, 'bold': ci == 0, 'line_spacing': 1.15}],
            anchor=MSO_ANCHOR.MIDDLE)
        x += w
    ry += rh
rect(s, Inches(0.55), Inches(ry + 0.12), Inches(12.2), Inches(0.8), fill=LIGHT)
rect(s, Inches(0.55), Inches(ry + 0.12), Inches(0.1), Inches(0.8), fill=ACCENT)
txt(s, Inches(0.8), Inches(ry + 0.12), Inches(11.8), Inches(0.8),
    [{'text': [{'text': '核心数据来源：', 'size': 12, 'color': PRIMARY, 'bold': True},
               {'text': '鸣鸣很忙集团 2024 招股书及年报、赵一鸣官网、GeoQ 智图行业蓝皮书、美团闪购行业白皮书。'
                        '关键口径：单店营收区分供货收入/GMV；毛利率分公司层/加盟商层/自有品牌；会员数据按集团合并口径。',
                'size': 11.5, 'color': INK}], 'line_spacing': 1.25}], anchor=MSO_ANCHOR.MIDDLE)
page(s)

# --- 结论与启示 ---
s = slide(); bg(s)
header(s, "05", "结论与启示")
concl = [
    ("战略成效", PROC, "并入集团后，赵一鸣借供应链、资本与数字化赋能，从区域强势品牌跃升为全国头部，"
                   "收入规模超越「零食很忙」，成本领先战略成效显著。"),
    ("执行短板", CUS, "毛利率、自有品牌占比、动销率等目标与现状差距较大；KPI 偏「结果导向」、缺过程性与能力性指标，"
                  "目标设定偏向行业对标而非自身能力。"),
    ("BSC 启示", LEARN, "战略地图需强化四维度因果传导：以「学习成长(数字化/人才)→内部流程(选品/周转)"
                    "→客户(分层/客单价)→财务(毛利/营收)」的能力链条支撑目标，避免「一刀切」。"),
    ("发展建议", PRIMARY, "①目标分层、设过程指标；②自有品牌建协同开发机制；③客户分区域精细运营；"
                    "④数字化由「系统建设」转向「流程再造+组织转型」；⑤平衡双品牌区域竞合。"),
]
y = 1.4
for name, col, txt_c in concl:
    rect(s, Inches(0.55), Inches(y), Inches(12.2), Inches(1.18), fill=LIGHT)
    rect(s, Inches(0.55), Inches(y), Inches(2.3), Inches(1.18), fill=col)
    txt(s, Inches(0.6), Inches(y), Inches(2.2), Inches(1.18),
        [{'text': name, 'size': 18, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(3.05), Inches(y + 0.1), Inches(9.55), Inches(1.0),
        [{'text': txt_c, 'size': 13, 'color': INK, 'line_spacing': 1.3}], anchor=MSO_ANCHOR.MIDDLE)
    y += 1.32
page(s)

# --- 谢谢 ---
s = slide(); bg(s, PRIMARY)
rect(s, 0, Inches(2.9), SW, Inches(1.7), fill=PRIMARY_DK)
txt(s, Inches(0), Inches(2.55), SW, Inches(1.2),
    [{'text': '感 谢 聆 听', 'size': 54, 'color': WHITE, 'bold': True, 'align': PP_ALIGN.CENTER}])
txt(s, Inches(0), Inches(3.85), SW, Inches(0.6),
    [{'text': 'THANKS  ·  欢迎点评与讨论', 'size': 18, 'color': RGBColor(0xF3, 0xCF, 0xC9),
      'align': PP_ALIGN.CENTER, 'font': FONT_EN}])
txt(s, Inches(0), Inches(5.0), SW, Inches(0.5),
    [{'text': '深入洞察战略蓝图，共绘零食行业新未来', 'size': 14,
      'color': RGBColor(0xF7, 0xE3, 0xDF), 'align': PP_ALIGN.CENTER}])

import os
out = os.path.join(os.path.dirname(__file__), "赵一鸣零食_战略地图案例分析.pptx")
prs.save(out)
print("Saved:", out, "| slides:", len(prs.slides))
