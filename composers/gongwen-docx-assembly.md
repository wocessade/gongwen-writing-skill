# Composer: gongwen-docx-assembly

**Purpose:** Generate a Chinese government official document (.docx) conforming to GB/T 9704-2012 (党政机关公文格式) using a python-docx script.

**Used by:** gongwen-writing G2 (drafting stage)

**Parameters:** Config dict with fields: `authority`, `document_type`, `document_number`, `title`, `main_recipient`, `body_sections`, `attachments`, `issuing_org`, `date`, `cc_organs`, optional `security`, `urgency`, `signatory`.

**CRITICAL RULE:** Must generate via Python script execution (`{pipeline_root}/scripts/generate_gongwen_docx.py`), not manually write prose in context. Configure the shipped script via a config dict, then execute it.

## Instructions

### 1. Pre-Writing Setup

Before generating the script call, confirm with the user (STOP-AND-ASK):
- 发文机关 (authority/issuing organ name)
- 文种 (document type: 通知, 报告, 请示, 批复, 函, 纪要, etc.)
- 标题 (document title)
- 主送机关 (main recipient)
- 正文要点 (key content points for body sections)
- 是否含附件 (attachments, if any)
- 成文日期 (date)
- 密级/紧急程度 (security classification / urgency, if any)

### 2. Python Script Template

The following functions are implemented in `{pipeline_root}/scripts/generate_gongwen_docx.py`. Reference this template when calling the script. The primary entry point is `generate_gongwen(config)`.

```python
from docx import Document
from docx.shared import Pt, Cm, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import warnings


def set_run_font(run, font_cn, font_en="Times New Roman", size=Pt(16), bold=False):
    """Set both East-Asian (eastAsia) and Latin (ascii + hAnsi) fonts on a run.

    Args:
        run: python-docx Run object
        font_cn: Chinese font name for w:eastAsia (e.g. '仿宋_GB2312', '黑体')
        font_en: Latin font name for w:ascii and w:hAnsi (default 'Times New Roman')
        size: font size in Pt (default Pt(16) = 三号)
        bold: bold flag (default False)
    """
    run.font.size = size
    run.bold = bold
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_cn)
    rFonts.set(qn('w:ascii'), font_en)
    rFonts.set(qn('w:hAnsi'), font_en)


def set_page_setup(section):
    """Configure A4 page margins per GB/T 9704-2012.

    - Top: 37mm, Bottom: 35mm, Left: 28mm, Right: 26mm
    - A4 (210mm x 297mm)
    """
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3.7)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)


def try_set_font_name(font_name):
    """Check if a font is available in the system; return font name or fallback.

    Args:
        font_name: desired font name
    Returns:
        font_name if likely available, fallback name with warning otherwise
    """
    # Common fonts that are almost always available
    common_fonts = {'宋体', '仿宋', '黑体', '楷体', 'SimSun', 'SimHei', 'FangSong', 'KaiTi'}
    # Fonts that require separate installation
    special_fonts = {
        '方正小标宋简体': ('宋体', '方正小标宋简体 未安装，回退到 宋体'),
        '仿宋_GB2312': ('仿宋', '仿宋_GB2312 未安装，回退到 仿宋'),
        '楷体_GB2312': ('楷体', '楷体_GB2312 未安装，回退到 楷体'),
    }
    if font_name in common_fonts:
        return font_name
    if font_name in special_fonts:
        fallback, msg = special_fonts[font_name]
        warnings.warn(msg)
        return fallback
    return font_name


def add_red_header(doc, authority_name):
    """Add the red document header (发文机关标志) per GB/T 9704-2012.

    - Red (RGB 255,0,0) 方正小标宋简体 (with fallback)
    - Centered
    - Maximum font height <= 22mm (use ~72pt for standard-length names)
    - Followed by a red horizontal divider line (红色反线, ~0.4mm thick)

    Args:
        doc: python-docx Document
        authority_name: full name of issuing authority (e.g. 'XX省人民政府办公厅')
    """
    header_font = try_set_font_name('方正小标宋简体')
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.0

    run = para.add_run(authority_name)
    set_run_font(run, header_font, font_en='SimSun', size=Pt(72), bold=False)
    run.font.color.rgb = RGBColor(255, 0, 0)  # Red

    # Red horizontal divider line (bottom border on the paragraph)
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '10')   # ~0.42mm (in 1/8 pt units)
    bottom.set(qn('w:color'), 'FF0000')
    bottom.set(qn('w:space'), '4')
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_urgency(doc, urgency_text):
    """Add urgency level (紧急程度) if specified.

    黑体 三号(16pt) bold, left-aligned at top of document.
    Placed between red header and document number.

    Args:
        doc: python-docx Document
        urgency_text: e.g. '特急' or '加急'
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    run = para.add_run(urgency_text)
    set_run_font(run, try_set_font_name('黑体'), size=Pt(16), bold=True)


def add_security(doc, security_text):
    """Add security classification (密级) if specified.

    黑体 三号(16pt) bold, left-aligned.

    Args:
        doc: python-docx Document
        security_text: e.g. '绝密', '机密', '秘密' (optionally with期限)
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    run = para.add_run(security_text)
    set_run_font(run, try_set_font_name('黑体'), size=Pt(16), bold=True)


def add_document_number(doc, number_text):
    """Add document number (发文字号) — 仿宋_GB2312 三号(16pt), centered.

    Args:
        doc: python-docx Document
        number_text: e.g. 'X政发〔2024〕1号'
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    run = para.add_run(number_text)
    set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))


def add_title(doc, text):
    """Add document title — 方正小标宋简体 二号(22pt), centered.

    Placed one line below the divider/red header area.

    Args:
        doc: python-docx Document
        text: document title (no punctuation at end)
    """
    # Blank line before title (spacer)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = Pt(28)

    title_font = try_set_font_name('方正小标宋简体')
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    run = para.add_run(text)
    set_run_font(run, title_font, font_en='SimSun', size=Pt(22))


def add_main_recipient(doc, text):
    """Add main recipient (主送机关) — 仿宋_GB2312 三号(16pt), left-aligned.

    Args:
        doc: python-docx Document
        text: recipient name(s), e.g. '各省、自治区、直辖市人民政府：'
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    run = para.add_run(text)
    set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))


def add_body_para(doc, text, indent=True):
    """Add a body paragraph — 仿宋_GB2312 三号(16pt), 28pt fixed line spacing.

    Args:
        doc: python-docx Document
        text: paragraph text
        indent: whether to apply first-line indent of 2 chars (default True)
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    if indent:
        pf.first_line_indent = Cm(0.85)  # ~2 chars for 三号(16pt)
    run = para.add_run(text)
    set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))


def add_heading_level1(doc, text):
    """Add level-1 heading (一级标题) — 黑体 三号(16pt), bold, e.g. '一、'.

    Args:
        doc: python-docx Document
        text: heading text with numbering, e.g. '一、提高政治站位'
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    run = para.add_run(text)
    set_run_font(run, try_set_font_name('黑体'), size=Pt(16), bold=True)


def add_heading_level2(doc, text):
    """Add level-2 heading (二级标题) — 楷体_GB2312 三号(16pt), e.g. '（一）'.

    Args:
        doc: python-docx Document
        text: heading text with numbering, e.g. '（一）加强组织领导'
    """
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    run = para.add_run(text)
    set_run_font(run, try_set_font_name('楷体_GB2312'), size=Pt(16))


def add_attachment(doc, attachments):
    """Add attachment description (附件说明) — 仿宋_GB2312 三号(16pt).

    Placed after the body, before signature. Format:
    附件：1. XXXXX
           2. XXXXX

    Args:
        doc: python-docx Document
        attachments: list of attachment description strings
    """
    if not attachments:
        return
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    pf.first_line_indent = Cm(0.85)
    prefix = '附件：' + ' '.join(f'{i+1}. {a}' for i, a in enumerate(attachments))
    run = para.add_run(prefix)
    set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))


def add_signature_and_date(doc, org_name, date_str):
    """Add issuing organ signature and date (发文机关署名 + 成文日期).

    Right-aligned, 右空四字 (approximately 4 characters from right margin).

    Args:
        doc: python-docx Document
        org_name: issuing organ name
        date_str: date in format '2024年1月15日'
    """
    # Blank spacer line before signature
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(0)
    spacer.paragraph_format.line_spacing = Pt(28)

    # Org name — right-aligned with right indent of ~4 chars
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    pf.right_indent = Cm(2.0)  # ~右空四字
    run = para.add_run(org_name)
    set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))

    # Date — same alignment
    para2 = doc.add_paragraph()
    pf2 = para2.paragraph_format
    pf2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf2.space_before = Pt(0)
    pf2.space_after = Pt(0)
    pf2.line_spacing = Pt(28)
    pf2.right_indent = Cm(2.0)
    run2 = para2.add_run(date_str)
    set_run_font(run2, try_set_font_name('仿宋_GB2312'), size=Pt(16))


def add_cc_organs(doc, cc_list):
    """Add CC organs (抄送) — 仿宋_GB2312 三号(16pt).

    Placed after signature and date, separated by a horizontal line.

    Args:
        doc: python-docx Document
        cc_list: list of organ names to CC
    """
    if not cc_list:
        return
    # Horizontal separator line
    para_sep = doc.add_paragraph()
    pf_sep = para_sep.paragraph_format
    pf_sep.space_before = Pt(0)
    pf_sep.space_after = Pt(0)
    pf_sep.line_spacing = Pt(28)
    pPr = para_sep._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '6')
    top.set(qn('w:color'), '000000')
    top.set(qn('w:space'), '1')
    pBdr.append(top)
    pPr.append(pBdr)

    # CC text
    cc_text = '抄送：' + '、'.join(cc_list)
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(28)
    run = para.add_run(cc_text)
    set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))


def add_page_number(section):
    """Add page number in standard gongwen format — 宋体 四号(14pt), '- 1 -' style.

    Uses odd/even page differentiation per GB/T 9704-2012:
    - Odd pages: right-aligned
    - Even pages: left-aligned

    Args:
        section: python-docx Section object
    """
    # Enable different odd/even headers
    sectPr = section._sectPr
    titlePg = OxmlElement('w:titlePg')
    sectPr.append(titlePg)

    # Even page footer — left-aligned
    even_footer = section.even_page_footer
    even_footer.is_linked_to_previous = False
    even_p = even_footer.paragraphs[0]
    even_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Odd page footer — right-aligned
    odd_footer = section.footer
    odd_footer.is_linked_to_previous = False
    odd_p = odd_footer.paragraphs[0]
    odd_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for p in [even_p, odd_p]:
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        # Add dash, page number field, dash:  - 1 -
        run_dash1 = p.add_run('— ')   # em dash + space (一字线)
        set_run_font(run_dash1, try_set_font_name('宋体'), font_en='Times New Roman', size=Pt(14))

        # PAGE field
        run_field = p.add_run()
        fldChar_begin = OxmlElement('w:fldChar')
        fldChar_begin.set(qn('w:fldCharType'), 'begin')
        run_field._r.append(fldChar_begin)

        run_instr = p.add_run()
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = ' PAGE '
        run_instr._r.append(instrText)

        run_sep = p.add_run()
        fldChar_sep = OxmlElement('w:fldChar')
        fldChar_sep.set(qn('w:fldCharType'), 'separate')
        run_sep._r.append(fldChar_sep)

        run_page = p.add_run()
        set_run_font(run_page, try_set_font_name('宋体'), font_en='Times New Roman', size=Pt(14))

        run_end = p.add_run()
        fldChar_end = OxmlElement('w:fldChar')
        fldChar_end.set(qn('w:fldCharType'), 'end')
        run_end._r.append(fldChar_end)

        run_dash2 = p.add_run(' —')   # space + em dash
        set_run_font(run_dash2, try_set_font_name('宋体'), font_en='Times New Roman', size=Pt(14))


def generate_gongwen(config):
    """Main entry point — generate a GB/T 9704-2012 compliant .docx.

    Args:
        config: dict with the following keys:
            authority (str): 发文机关全称
            document_type (str): 文种 (informational, used for filename)
            document_number (str): 发文字号, e.g. 'X政发〔2024〕1号'
            title (str): 公文标题
            main_recipient (str): 主送机关
            body_sections (list): body content, each is dict:
                - 'level': 0 (一级), 1 (二级), or None (正文)
                - 'text': content string
            attachments (list, optional): list of attachment descriptions
            issuing_org (str): 发文机关署名
            date (str): 成文日期, e.g. '2024年1月15日'
            cc_organs (list, optional): 抄送机关列表
            security (str, optional): 密级, e.g. '机密'
            urgency (str, optional): 紧急程度, e.g. '特急'
            output_path (str, optional): output file path
            signatory (str, optional): 签发人 (上行文用)

    Returns:
        str: output file path of the generated .docx
    """
    doc = Document()

    # -- Default style --
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    style.font.size = Pt(16)

    # -- Page Setup --
    section = doc.sections[0]
    set_page_setup(section)

    # -- Security classification (if specified) --
    if config.get('security'):
        add_security(doc, config['security'])

    # -- Urgency (if specified) --
    if config.get('urgency'):
        add_urgency(doc, config['urgency'])

    # -- Red Header --
    add_red_header(doc, config['authority'])

    # -- Document number --
    add_document_number(doc, config['document_number'])

    # -- Title --
    add_title(doc, config['title'])

    # -- Main recipient --
    add_main_recipient(doc, config['main_recipient'])

    # -- Body sections --
    for section_item in config.get('body_sections', []):
        level = section_item.get('level')
        text = section_item.get('text', '')
        if level == 0:
            add_heading_level1(doc, text)
        elif level == 1:
            add_heading_level2(doc, text)
        else:
            add_body_para(doc, text, indent=True)

    # -- Signatory (上行文签发人, after body) --
    if config.get('signatory'):
        para = doc.add_paragraph()
        pf = para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(28)
        run = para.add_run(f'签发人：{config["signatory"]}')
        set_run_font(run, try_set_font_name('仿宋_GB2312'), size=Pt(16))

    # -- Attachments --
    add_attachment(doc, config.get('attachments', []))

    # -- Signature and date --
    add_signature_and_date(doc, config['issuing_org'], config['date'])

    # -- CC organs --
    add_cc_organs(doc, config.get('cc_organs', []))

    # -- Page numbers --
    add_page_number(section)

    # -- Save --
    output_path = config.get('output_path', os.path.join(
        os.getcwd(),
        f'{config["authority"]}_{config["document_type"]}.docx'
    ))
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    doc.save(output_path)
    return output_path
```

### 3. Execution Pattern

```python
# Step 1: Build config dict
config = {
    "authority": "XX省人民政府办公厅",
    "document_type": "通知",
    "document_number": "X政发〔2024〕1号",
    "title": "关于做好2024年元旦春节期间有关工作的通知",
    "main_recipient": "各市、县人民政府，省政府各部门：",
    "body_sections": [
        {"level": 0, "text": "一、提高思想认识"},
        {"level": 1, "text": "（一）充分认识做好元旦春节期间工作的重要性"},
        {"level": None, "text": "正文内容……"},
    ],
    "attachments": ["1. 任务分工表"],
    "issuing_org": "XX省人民政府办公厅",
    "date": "2024年1月15日",
    "cc_organs": ["省委办公厅", "省人大常委会办公厅", "省政协办公厅"],
}

# Step 2: Import and run
import sys
sys.path.insert(0, '{pipeline_root}/scripts')
from generate_gongwen_docx import generate_gongwen

output_path = generate_gongwen(config)
print(f"Document saved to: {output_path}")
```

### 4. Document Structure (in order)

| Element | Notes |
|---------|-------|
| 密级 (optional) | 黑体 三号 bold, left-aligned |
| 紧急程度 (optional) | 黑体 三号 bold, left-aligned |
| 发文机关标志(红头) | 红色 方正小标宋简体, centered |
| 红色分隔线 | Red horizontal line |
| 发文字号 | 仿宋_GB2312 三号, centered |
| 标题 | 方正小标宋简体 二号(22pt), centered |
| 主送机关 | 仿宋_GB2312 三号, left-aligned |
| 正文(多段) | 仿宋_GB2312 三号, 28pt line spacing, first-line indent |
| 一级标题 | 黑体 三号 bold, 一、格式 |
| 二级标题 | 楷体_GB2312 三号, （一）格式 |
| 附件说明 (optional) | 仿宋_GB2312 三号 |
| 签发人 (上行文optional) | 仿宋_GB2312 三号 |
| 发文机关署名 | 仿宋_GB2312 三号, right-aligned, 右空四字 |
| 成文日期 | 仿宋_GB2312 三号, right-aligned |
| 抄送 (optional) | 仿宋_GB2312 三号, after horizontal separator |

### 5. Execute & Verify

1. Run the Python script to generate the .docx file
2. Verify:
   - File exists and is > 10KB
   - Open with python-docx to check paragraph count
   - Red header text is present (font color RGB 255,0,0)
   - Page margins are set correctly (3.7, 3.5, 2.8, 2.6 cm)
   - No placeholder text ("[待填写]", "TODO", etc.)
   - Title font is 方正小标宋简体 二号(22pt)
   - Body font is 仿宋_GB2312 三号(16pt)
   - Line spacing is fixed 28pt

## Verification

- [ ] Config dict built with all required fields
- [ ] `generate_gongwen()` entry point called
- [ ] Page margins: top 3.7cm, bottom 3.5cm, left 2.8cm, right 2.6cm
- [ ] Red header: red (255,0,0), centered, with red divider line
- [ ] Title: 方正小标宋简体 二号(22pt), centered
- [ ] Body: 仿宋_GB2312 三号(16pt), 28pt fixed line spacing
- [ ] Level-1 headings: 黑体 三号 bold, '一、' format
- [ ] Level-2 headings: 楷体_GB2312 三号, '（一）' format
- [ ] Recipient: left-aligned, 仿宋_GB2312 三号
- [ ] Signature: right-aligned, 右空四字
- [ ] Page numbers: 宋体 四号(14pt), '- 1 -' format, odd/right even/left
- [ ] File > 10KB, all sections present
- [ ] No placeholder text
- [ ] Font fallback warnings issued for missing special fonts

## Common Pitfalls

- **Font availability:** 方正小标宋简体, 仿宋_GB2312, 楷体_GB2312 are NOT installed by default on Windows. The script provides fallbacks with warnings, but the user should install these fonts for proper formatting.
- **Pt(0) is falsy in Python:** Use `is not None`, not `if sb:` when checking paragraph spacing. Pt(0) evaluates to False in a boolean context.
- **Red divider thickness:** The `w:sz` attribute on paragraph borders uses 1/8 pt units. sz="10" = 1.25pt ≈ 0.44mm. Adjust between sz="8" and sz="12" for 0.35-0.5mm.
- **Page number positioning:** Per GB/T 9704-2012, single pages right-aligned, double pages left-aligned. This is done via `section.even_page_footer` and `section.footer` (odd).
- **行距 must be fixed 28pt:** Always set `pf.line_spacing = Pt(28)` for body text. Do not use 1.0/1.5/2.0 multipliers — the standard requires fixed value.
- **Right indent for signature:** Use `pf.right_indent = Cm(2.0)` to approximate 右空四字. Adjust if needed.
- **Em dashes in page numbers:** Use Unicode U+2014 (em dash) or U+2015 (horizontal bar) for the 一字线 around page numbers.
