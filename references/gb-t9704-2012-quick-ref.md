# GB/T 9704-2012 党政机关公文格式 — 快速参考

## 页面设置 (A4, 210mm × 297mm)

| 属性 | 数值 | 备注 |
|------|------|------|
| 版心尺寸 | 156mm × 225mm | 文字区域 |
| 上白边 (天头) | 37mm ± 1mm | section.top_margin = Cm(3.7) |
| 下白边 | 35mm ± 1mm | section.bottom_margin = Cm(3.5) |
| 左白边 (订口) | 28mm ± 1mm | section.left_margin = Cm(2.8) |
| 右白边 (翻口) | 26mm ± 1mm | section.right_margin = Cm(2.6) |
| 行数 | 22行/面 | 28pt固定行距 |
| 字数 | 28字/行 | 三号字 |

## 字体字号对照表

| 要素 | 字体 | 字号 | pt值 | 粗体 |
|------|------|------|------|------|
| 发文机关标志(红头) | 方正小标宋简体 | — | 高度≤22mm | 否 |
| 密级/保密期限 | 黑体 | 三号 | 16pt | 是 |
| 紧急程度 | 黑体 | 三号 | 16pt | 是 |
| 发文字号 | 仿宋_GB2312 | 三号 | 16pt | 否 |
| 签发人姓名 | 楷体_GB2312 | 三号 | 16pt | 否 |
| **标题** | **方正小标宋简体** | **二号** | **22pt** | 否 |
| 主送机关 | 仿宋_GB2312 | 三号 | 16pt | 否 |
| **正文** | **仿宋_GB2312** | **三号** | **16pt** | 否 |
| 一级标题 (一、) | 黑体 | 三号 | 16pt | 是 |
| 二级标题 ((一)) | 楷体_GB2312 | 三号 | 16pt | 否 |
| 三级标题 (1.) | 仿宋_GB2312 | 三号 | 16pt | 否 |
| 四级标题 ((1)) | 仿宋_GB2312 | 三号 | 16pt | 否 |
| 附件说明 | 仿宋_GB2312 | 三号 | 16pt | 否 |
| 发文机关署名 | 仿宋_GB2312 | 三号 | 16pt | 否 |
| 成文日期 | 仿宋_GB2312 | 三号 | 16pt | 否 |
| 抄送机关 | 仿宋_GB2312 | 三号 | 16pt | 否 |
| **页码** | **宋体** | **四号** | **14pt** | 否 |

## 字号与pt换算

| 字号 | pt |
|------|-----|
| 二号 | 22pt |
| 三号 | 16pt |
| 小三号 | 15pt |
| 四号 | 14pt |
| 小四号 | 12pt |
| 五号 | 10.5pt |

## 行距与间距

| 要素 | 行距/间距 | python-docx 实现 |
|------|-----------|------------------|
| 正文 | 固定值28pt | pf.line_spacing = Pt(28) |
| 标题与正文间距 | 下空一行 | — |
| 标题与主送间距 | 上空一行 | — |
| 正文段落间 | 无空行（不设段前段后距） | pf.space_before = Pt(0), pf.space_after = Pt(0) |
| 正文缩进 | 首行缩进2字符 ≈ 16pt × 2 | pf.first_line_indent = Cm(0.85) |

## 标题层级规则

| 层级 | 编号格式 | 字体 | 对齐 |
|------|---------|------|------|
| 一级 | 一、二、三、 | 黑体 三号 | 顶格 |
| 二级 | （一）（二）（三） | 楷体_GB2312 三号 | 顶格 |
| 三级 | 1. 2. 3. | 仿宋_GB2312 三号 | 首行缩进2字符 |
| 四级 | (1) (2) (3) | 仿宋_GB2312 三号 | 首行缩进2字符 |

## 红头规范

- 发文机关标志居中，上边缘至版心上边缘 35mm
- 红色（RGB: 255,0,0），方正小标宋简体
- 最大高度 ≤ 22mm
- 下方为红色分隔线（红色反线），宽度为版心宽度，厚度 0.35mm ~ 0.5mm

## 页码规范

- 宋体 四号（14pt）半角阿拉伯数字
- 页码数字左右各放一条一字线，如 "- 1 -"
- 单页码居右，双页码居左
- 版心下边缘之下 7mm

## 发文字号格式

机关代字 + 〔年份〕+ 序号

示例：X政发〔2024〕1号

- 年份用六角括号〔〕，不能用[]或【】
- 序号不编虚位（即1不编为001）

## 成文日期

- 用阿拉伯数字标全年、月、日
- 示例：2024年1月15日（不写"二零二四年一月十五日"）

## python-docx 关键代码

```python
from docx import Document
from docx.shared import Pt, Cm, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# 页面设置
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(3.7)
section.bottom_margin = Cm(3.5)
section.left_margin = Cm(2.8)
section.right_margin = Cm(2.6)

# 设置东亚字体与西文字体
def set_run_font(run, font_cn, font_en="Times New Roman", size=Pt(16), bold=False):
    run.font.size = size
    run.bold = bold
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_cn)
    rFonts.set(qn("w:ascii"), font_en)
    rFonts.set(qn("w:hAnsi"), font_en)

# 固定行距28pt
from docx.shared import Pt as PtSpacing
pf.line_spacing = Pt(28)
# 或者通过OxmlElement精确设置:
spacing = OxmlElement("w:spacing")
spacing.set(qn("w:line"), "560")  # 28pt * 20 twips/pt
spacing.set(qn("w:lineRule"), "exact")
```

## 字体回退策略

当系统缺少指定字体时：

| 缺失字体 | 回退字体 | 影响 |
|---------|---------|------|
| 方正小标宋简体 | 宋体 + 警告 | 红头/标题非标准 |
| 仿宋_GB2312 | 仿宋/FangSong | 正文非标准 |
| 楷体_GB2312 | 楷体/KaiTi | 二级标题非标准 |
| 黑体 | SimHei | 一级标题非标准 |

> 注：方正小标宋简体 是方正商业字体，Windows 系统默认不安装。正式公文需确保该字体已安装在系统中。
