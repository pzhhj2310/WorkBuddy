# -*- coding: utf-8 -*-
"""生成《1000 元二手主机跑本地大模型 · 最低配置方案》Word 文档"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ACCENT = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x59, 0x59, 0x59)
FONT_CN = "微软雅黑"

doc = Document()

# ---------- 全局样式 ----------
sec = doc.sections[0]
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.4)
sec.right_margin = Cm(2.4)

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CN)
normal.paragraph_format.line_spacing = 1.45
normal.paragraph_format.space_after = Pt(6)


def set_cn(run, size=None, bold=None, color=None, name=FONT_CN):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def para(text="", size=10.5, bold=False, color=None, space_before=0, space_after=6,
         align=None, indent=0, line=1.45, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if align:
        p.alignment = align
    if text:
        r = p.add_run(text)
        set_cn(r, size, bold, color)
        r.italic = italic
    return p


def rich(parts, size=10.5, space_before=0, space_after=6, indent=0, line=1.45):
    """parts: [(text, bold, color) ...]"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    for text, bold, color in parts:
        r = p.add_run(text)
        set_cn(r, size, bold, color)
    return p


def h1(text, space_before=16):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    set_cn(r, 14, True, ACCENT)
    # 下边框
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "1F3864")
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p


def bullet(text, level=0, size=10.5):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.left_indent = Cm(0.75 + 0.6 * level)
    r = p.add_run(text)
    set_cn(r, size)
    return p


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def cell_text(cell, text, size=9.5, bold=False, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.2
    if align:
        p.alignment = align
    r = p.add_run(text)
    set_cn(r, size, bold, color)


def table(header, rows, widths, header_fill="1F3864", zebra="EEF2F8", size=9.5,
          aligns=None):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.width = Cm(widths[i])
        shade(c, header_fill)
        cell_text(c, h, size, True, RGBColor(0xFF, 0xFF, 0xFF),
                  WD_ALIGN_PARAGRAPH.CENTER if i else None)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].width = Cm(widths[i])
            if zebra and ri % 2 == 1:
                shade(cells[i], zebra)
            al = None
            if aligns and aligns[i] == "c":
                al = WD_ALIGN_PARAGRAPH.CENTER
            bold = i == len(row) - 1 and header[-1] in ("金额（元）", "合计")
            cell_text(cells[i], str(val), size, bold, None, al)
    # 表头重复
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def box(lines, fill="F2F5FA", bar="1F3864"):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    cell = t.rows[0].cells[0]
    cell.width = Cm(16.2)
    shade(cell, fill)
    cell.text = ""
    first = True
    for text, bold, size, color in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.4
        r = p.add_run(text)
        set_cn(r, size, bold, color)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ================= 封面/标题 =================
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("1000 元以内在闲鱼搭一台能跑本地大模型的电脑")
set_cn(r, 19, True, ACCENT)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("最低配置要求 · 二手配件清单 · 实测能跑多大模型")
set_cn(r, 11, False, GREY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(14)
r = p.add_run("行情基准日：2026 年 11 月 8 日（闲鱼 / 转转成交价）")
set_cn(r, 9, False, GREY)

box([
    ("一句话结论", True, 11, ACCENT),
    ("1000 元这一档，唯一能落地的路线是「二手至强 E5 平台 + 6GB 矿卡」。"
     "主力方案总价约 970 元：E5-2680 v4（14 核 28 线程）+ X99 + 32G ECC 内存 + P106-100 6G。"
     "能流畅跑 4B～8B 的 Q4 量化模型（15～30 字/秒），14B 以上跑不动。", False, 10, None),
    ("如果预算能到 1100 元，把显卡换成 8GB 的 P104-100，就刚好摸到 14B 的门槛——这 150 元是整份单子里性价比最高的一笔。",
     False, 10, None),
])

# ================= 一、先搞清楚最低要求 =================
h1("一、本地跑模型，到底卡的是什么")

para("很多人以为要「一台几万块的服务器」，其实不是。本地推理的性能由四件事决定，按重要程度排：", 10.5)

table(
    ["部件", "为什么关键", "最低要求（跑 7～8B）", "本方案给到"],
    [
        ["显存（显卡）", "模型权重放得进显存就全速跑，放不进就溢出到内存，速度直接掉 3～5 倍", "6GB", "6GB（P106-100）"],
        ["内存", "系统 + 模型加载缓冲 + CPU 分担层", "16GB（推荐 32GB）", "32GB ECC"],
        ["硬盘", "模型文件单个 3～20GB，机械盘加载要几分钟", "256GB 固态", "256GB SATA SSD"],
        ["CPU", "需支持 AVX2 指令集；显存不够时靠它兜底", "2013 年后 Intel / 支持 AVX2", "E5-2680 v4 支持 AVX2"],
        ["电源", "单卡整机峰值 250～300W", "额定 400～500W", "500W 二手品牌"],
    ],
    [2.2, 5.6, 4.4, 4.0],
)

rich([("显存是唯一的硬门槛。", True, ACCENT),
      ("6GB 是 2026 年跑 Q4 量化 7～8B 模型的及格线；想碰 14B（Q4 约 8.5GB），必须 8GB 起。"
       "显存容量决定「能不能跑」，显存带宽决定「跑多快」——这两条比 CPU 主频重要得多。", False, None)],
     10.5, indent=0)

# ================= 二、方案一 =================
h1("二、方案一（首选）：970 元主力机")

para("思路：用服务器平台「论颗卖」的二手至强换核心数和内存带宽，把省下来的钱全部给显卡。"
     "X99 平台是四通道 DDR4，内存带宽 25.6GB/s，是普通台式双通道的两倍——纯 CPU 兜底时的优势很明显。", 10.5)

table(
    ["部件", "型号 / 规格", "二手价（元）", "选购要点"],
    [
        ["CPU", "Intel 至强 E5-2680 v4（14 核 28 线程 / 35MB 缓存 / 120W）", "150", "只要「拆机正式片」，别碰 ES/QS 工程片"],
        ["主板", "X99 寨板（华南、精粤、铭瑄等）", "200", "必须确认 BIOS 已刷好支持 v4；要 CPU-Z 主板页截图，芯片组应为 C612"],
        ["内存", "32GB DDR4 ECC REG（4×8G，插满四通道）", "160", "只买 REG 条，不能与普通 UDIMM 混插"],
        ["显卡", "P106-100 6GB（GP106 核心，等同 GTX 1060 6G 去视频输出）", "170", "要求 GPU-Z 截图 + FurMark 烤机 15 分钟温度图"],
        ["亮机卡", "GT 710 / HD 7450（P106 无视频输出，必须配）", "40", "随便一张能出画面的卡即可"],
        ["硬盘", "256GB SATA 固态（拆机亦可）", "60", "SATA 就行，X99 寨板的 M.2 多为残血"],
        ["电源", "额定 500W 品牌二手（航嘉 / 长城 / 台达）", "90", "别买杂牌，炸了会带走内存和硬盘"],
        ["散热", "六热管塔式 + LGA2011-3 扣具", "60", "120W 的 U，原装小风扇压不住，满载会摸到 95℃"],
        ["机箱", "二手 ATX 中塔", "40", "能塞下塔式散热即可"],
        ["合计", "—", "970", "预算内完成"],
    ],
    [1.9, 6.6, 2.4, 5.3],
    aligns=["c", None, "c", None],
)

rich([("关于 P106-100 要说清楚：", True, ACCENT),
      ("它是矿卡，没有视频输出口，所以必须再插一张亮机卡接显示器；Windows 下需要魔改驱动（改注册表 + 指定老版驱动，且官方支持到 580 分支为止，之后不再更新），"
       "折腾成本不低。但如果在 Ubuntu 上跑，标准 NVIDIA 驱动可直接识别，无需改驱动——", False, None),
      ("强烈建议直接装 Ubuntu 22.04 LTS，这是这条路线最大的省心点。", True, None)],
     10.5)

# ================= 三、方案二 =================
h1("三、方案二（+150 元升级）：1100 元，摸到 14B")

para("其余配件完全不变，只把显卡从 P106-100（6GB / 192GB/s）换成 P104-100（8GB / 256bit / 256GB/s）。"
     "8GB 显存刚好卡住 14B 模型 Q4 量化（约 8.5GB），能全部放进显存，速度 8～10 字/秒，日常对话可用。", 10.5)

table(
    ["项目", "P106-100（6GB）", "P104-100（8GB）"],
    [
        ["二手价", "150～200 元", "280～350 元"],
        ["显存 / 带宽", "6GB GDDR5 / 192GB/s", "8GB GDDR5 / 256GB/s"],
        ["能跑上限", "7～8B（Q4）", "14B（Q4，卡线）"],
        ["7B 速度", "15～18 字/秒", "20 字/秒左右"],
        ["14B 速度", "不可用（爆显存，掉到 3～5）", "8～10 字/秒"],
        ["视频输出", "无，需亮机卡", "无，需亮机卡"],
    ],
    [3.4, 6.4, 6.4],
    aligns=["c", "c", "c"],
)

para("结论：预算真能到 1100，就别省这 150 元。6GB 到 8GB 是「能不能跑 14B」的分水岭，"
     "而 14B 模型的回答质量比 7B 有明显提升。", 10.5)

# ================= 四、方案三 =================
h1("四、方案三：纯 CPU 路线（930 元，不插显卡）")

para("完全放弃显卡，把钱全砸在内存上：64GB ECC 内存，靠四通道带宽硬跑。适合对速度不敏感、"
     "只需要「能跑通、能离线、不花钱」的场景（批量文本处理、挂机总结、给局域网当个 API）。", 10.5)

table(
    ["部件", "型号 / 规格", "二手价（元）"],
    [
        ["CPU", "E5-2680 v4", "150"],
        ["主板", "X99 寨板", "200"],
        ["内存", "64GB DDR4 ECC REG（4×16G 四通道）", "300"],
        ["显卡", "亮机卡（主板无核显，必须有）", "40"],
        ["硬盘", "256GB SATA 固态", "60"],
        ["电源", "额定 450W 二手品牌", "80"],
        ["散热", "六热管塔式 + LGA2011-3 扣具", "60"],
        ["机箱", "二手 ATX 中塔", "40"],
        ["合计", "—", "930"],
    ],
    [2.0, 10.0, 4.2],
    aligns=[None, None, "c"],
)

para("缺点很明确：7B 模型只有 4～5 字/秒，14B 约 2.5～3 字/秒，长回答要等半分钟以上。"
     "它更适合当「后台批量任务机」，不适合当聊天助手。", 10.5)

# ================= 五、能跑什么 =================
h1("五、到底能跑什么模型、多快")

para("下表速度基于 Ollama + GGUF Q4_K_M 量化的社区实测均值，实际会因模型版本、上下文长度、驱动版本有 ±30% 浮动。", 9.5, color=GREY)

table(
    ["模型规模（Q4 量化）", "占用显存", "P106-100 6G", "P104-100 8G", "纯 CPU 64G"],
    [
        ["4B 级（如 Qwen3.5-4B）", "约 2.8GB", "30 字/秒", "30 字/秒", "6～8 字/秒"],
        ["7～8B 级（如 Qwen3.5-7B / 8B）", "约 4.5GB", "15～18 字/秒", "20 字/秒", "4～5 字/秒"],
        ["推理强化 7B（如 DeepSeek-R1-7B）", "约 5.2GB", "10～12 字/秒", "15 字/秒", "3～4 字/秒"],
        ["12～14B 级", "约 7～8.5GB", "跑不动", "8～10 字/秒", "2.5～3 字/秒"],
        ["30B 以上 / 多模态", "10GB+", "跑不动", "跑不动", "跑不动"],
    ],
    [5.0, 2.6, 2.9, 2.9, 2.8],
    aligns=[None, "c", "c", "c", "c"],
)

rich([("还有两个容易被忽略的体验问题：", True, ACCENT)], 10.5, space_after=3)
bullet("预填充慢。P106/P104 都是 Pascal 架构，没有 Tensor Core，处理长文档输入（比如丢进去一篇 1 万字）时，"
       "等模型「读完」的时间比生成答案还长。做单轮闲聊没感觉，做长文档总结会很难受。")
bullet("上下文越长越吃显存。6GB 卡把上下文开到 8K 以上就会开始爆显存，日常建议锁在 4K。")

# ================= 六、落地步骤 =================
h1("六、系统怎么装（30 分钟能跑起来）")

table(
    ["步骤", "操作", "备注"],
    [
        ["1", "装 Ubuntu 22.04 LTS（U 盘启动盘）", "别装 Windows，P106 在 Win 下要魔改驱动"],
        ["2", "装 NVIDIA 驱动：sudo apt install nvidia-driver-580", "P106 可用 580 分支，用 nvidia-smi 验证能看到 6GB"],
        ["3", "装 Ollama：curl -fsSL https://ollama.com/install.sh | sh", "自动接管 GPU，无需手动配 CUDA"],
        ["4", "拉模型：ollama run qwen3.5:7b", "首次会下载约 4.5GB，注意国内外网速"],
        ["5", "（可选）开放 API：默认监听 11434 端口", "局域网其他设备可直接调，等于自建一个 AI 服务"],
    ],
    [1.3, 9.5, 5.4],
    aligns=["c", None, None],
)

para("如果想要图形界面，可以再装 Open WebUI（一条 Docker 命令），或者直接用 LM Studio（但仅限 Windows/macOS，"
     "本方案不推荐 Windows 路线的矿卡）。", 10.5)

# ================= 七、避坑 =================
h1("七、买之前必须核对的清单")

para("这套配置几乎全是「无质保、坏了自负」的散件，翻车概率不低。下面每一条都是踩过的坑，下单前逐条问卖家：", 10.5)

items = [
    ("主板", "要 CPU-Z 主板页截图，芯片组必须是 C612（标 X99 却只有双通道的坑板会砍掉一半带宽）；确认 BIOS 已刷好支持 v4，刷崩会变砖。"),
    ("CPU", "要 CPU-Z 的「步进/修订」页，必须是 Broadwell 正式版；ES/QS 工程片一律 pass。"),
    ("内存", "只买 ECC REG（RDIMM），插普通 UDIMM 点不亮；四通道要四条满插，只插两条就是双通道。"),
    ("显卡", "要 GPU-Z 截图 + FurMark 烤机 15 分钟温度图，核心上 85℃ 一律不要；看 PCB 背面有没有发黄、电容鼓包、显存补焊痕迹。"),
    ("电源", "只买航嘉/长城/台达等品牌额定功率，杂牌标「800W」实际 300W，炸电源会带走主板和内存。"),
    ("交易", "退换期限写进闲鱼聊天记录，口头说「稳定跑」不作数；全程走平台，凡是让加微信转账的都是坑。"),
]
for k, v in items:
    rich([(k + "：", True, ACCENT), (v, False, None)], 10.5, space_after=4, indent=0.2)

# ================= 八、隐性成本 =================
h1("八、几个没人提前告诉你的成本")

table(
    ["项目", "实际情况"],
    [
        ["电费", "待机 80～90W，满载 250～300W。按 0.55 元/度、每天满载 4 小时算，月电费约 45～50 元，一年 500 元以上"],
        ["噪音", "LGA2011 平台 + 塔式散热 + 矿卡风扇，满载时明显吵，不适合放卧室"],
        ["质保", "全机 0 质保。任何一件坏了都要重新买，实际故障率比新机高 3～5 倍"],
        ["升级空间", "没有。LGA2011-3 是 2016 年的平台，DDR4 ECC 到顶就 64GB，显卡最多上到 8GB 老卡"],
        ["时间", "蹲货、验机、刷 BIOS、装系统，第一台通常要花掉一个周末"],
    ],
    [3.0, 13.2],
)

# ================= 九、建议 =================
h1("九、我的建议")

rich([("如果你的目标是「先低成本体验本地模型」，", False, None),
      ("按方案一（970 元）或方案二（1100 元）执行", True, ACCENT),
      ("，别再多花一分钱——这个平台的价值就是「花小钱把本地推理跑通」。", False, None)], 10.5)

rich([("如果你是想长期用，", False, None),
      ("这条路不建议走。", True, ACCENT),
      ("预算到 1300～1500 时，二手 RTX 3060 12G（约 900～1200 元）+ i3-12100F/H610 平台（约 700 元）贵不了多少，"
       "但显存 12GB 能跑 14B 不卡线、有视频输出不用亮机卡、驱动完整支持、整机还有质保——", False, None),
      ("这是 2026 年 1500 元档最合理的组合。", True, None)], 10.5)

para("最后一句实话：本地跑模型最贵的不是硬件，是你的时间。一张卡省 800 元、速度慢 5 倍、"
     "每天多等几个小时——这笔账比配置单更值得先算一遍。", 10.5, italic=True, color=GREY)

# 页脚
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(18)
r = p.add_run("价格均为 2026 年 11 月闲鱼/转转成交价区间中位值，实际成交随成色浮动 ±20%。")
set_cn(r, 8.5, False, GREY)

out = r"D:/workbuddy/Claw/1000元二手主机跑本地大模型_最低配置方案.docx"
doc.save(out)
print("saved:", out)
