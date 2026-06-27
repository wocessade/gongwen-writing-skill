#!/usr/bin/env python3
"""Parse an existing Chinese official document into a template JSON.

Analyzes a user-provided document (text or .docx), identifies its structure
(title, headings, body paragraphs), and outputs a template JSON matching the
format documented in templates/README.md.

Usage:
    # From text file
    python create_template.py --input document.txt

    # From stdin (paste text)
    python create_template.py < pasted_doc.txt

    # From .docx file
    python create_template.py --input document.docx

    # Save to file
    python create_template.py --input doc.txt --output templates/custom-template.json
"""

import argparse
import json
import os
import re
import sys


# ── Heading Pattern Registry ──────────────────────────────────────────────

# Level-1 heading (一级标题): 一、二、三、... 十、
_L0_RE = re.compile(r'^[一二三四五六七八九十]+、')

# Level-2 heading (二级标题): （一）（二）（三）... (1) (2) (3)
_L1_RE = re.compile(r'^（[一二三四五六七八九十]+）')

# Title pattern: 关于...的XX
_TITLE_RE = re.compile(r'关于.*的(通知|报告|请示|批复|函|纪要|通报|通告|公告|决定|决议|命令|公报|意见|议案)$')

# Main recipient pattern: lines ending with fullwidth colon
_RECIPIENT_RE = re.compile(r'.+：$')

# Document type keywords in title
_TYPE_KEYWORDS = {
    '通知': '通知', '报告': '报告', '请示': '请示', '批复': '批复',
    '函': '函', '纪要': '纪要', '通报': '通报', '通告': '通告',
    '公告': '公告', '决定': '决定', '决议': '决议', '命令': '命令',
    '意见': '意见', '议案': '议案',
}


# ── Parsers ───────────────────────────────────────────────────────────────

def infer_document_type(paragraphs):
    """Infer document type from title or content keywords."""
    for p in paragraphs[:10]:
        m = _TITLE_RE.search(p)
        if m:
            return m.group(1)
    for p in paragraphs[:10]:
        for kw, dtype in _TYPE_KEYWORDS.items():
            if kw in p:
                return dtype
    return '通知'


def _detect_level(text):
    """Detect heading level from text content. Returns 0, 1, or None."""
    if _L0_RE.match(text):
        return 0
    if _L1_RE.match(text):
        return 1
    return None


def _is_title(text, prev_level):
    """Heuristic: first line that looks like a document title."""
    if _TITLE_RE.search(text):
        return True
    if not prev_level and len(text) >= 6 and '关于' in text:
        return True
    return False


def _generate_template_id(title, paragraphs):
    """Generate a unique template ID from the document content."""
    base = 'custom'
    m = _TITLE_RE.search(title)
    if m:
        # Extract key topic from 关于...的XX
        topic_match = re.search(r'关于(.+?)的', title)
        if topic_match:
            topic = topic_match.group(1)
            # Take first meaningful part
            parts = re.split(r'[，、]', topic)
            key = parts[0][:12]
            base = f'{m.group(1).lower()}-{key}'
    # Fallback: use first few content words
    if base == 'custom':
        for p in paragraphs[1:6]:
            words = re.findall(r'[一-鿿]{2,6}', p)
            if words:
                base = f'template-{words[0]}'
                break
    # Sanitize: only allow safe chars
    base = re.sub(r'[^a-z0-9一-鿿_-]', '', base)
    base = re.sub(r'_+', '_', base)
    base = base[:40].strip('_-')
    return base or 'custom-template'


def parse_text(text):
    """Parse plain text into template structure.

    Args:
        text: Full document text (multiline string).

    Returns:
        dict with template_id, name, description, document_type,
        source_refs, defaults, body_sections, hints.
    """
    raw_lines = text.strip().split('\n')
    paragraphs = [p.strip() for p in raw_lines if p.strip()]

    if not paragraphs:
        raise ValueError("Empty input — no paragraphs found")

    # ── Identify title (search up to first 15 paragraphs) ──
    title = paragraphs[0]
    title_idx = 0
    for i, p in enumerate(paragraphs[:15]):
        if _is_title(p, i):
            title = p
            title_idx = i
            break
    # If no clear title found, use the longest first-line that looks like one
    if title_idx == 0 and not _is_title(paragraphs[0], 0):
        # Fallback: first line > 8 chars that contains Chinese content
        for i, p in enumerate(paragraphs[:15]):
            if len(p) >= 10 and any('一' <= c <= '鿿' for c in p[:5]):
                title = p
                title_idx = i
                break

    doc_type = infer_document_type(paragraphs)
    template_id = _generate_template_id(title, paragraphs)
    name = title[:30] + ('...' if len(title) > 30 else '')

    # ── Parse body sections ──
    body_sections = []
    for idx, p in enumerate(paragraphs):
        text = p
        # Skip everything before the title (red header, doc number, etc.)
        if idx < title_idx:
            continue
        # Skip the title paragraph itself (unless it looks like 主送机关)
        if idx == title_idx:
            if _RECIPIENT_RE.match(text) and len(text) < 40:
                body_sections.append({"level": None, "text": text})
            continue

        level = _detect_level(text)
        if level is not None:
            body_sections.append({"level": level, "text": text})
        else:
            body_sections.append({"level": None, "text": text})

    # ── Generate description ──
    m = _TITLE_RE.search(title)
    if m:
        topic = ''
        topic_match = re.search(r'关于(.+?)的', title)
        if topic_match:
            topic = topic_match.group(1)[:20]
        description = f'从"{title}"自动提取的{doc_type}模板'
    else:
        description = f'从用户提供的{doc_type}文档自动提取的模板'

    # ── Generate hints ──
    # Collect unique heading labels for hints
    level0_labels = [s['text'] for s in body_sections if s['level'] == 0]
    hints = {
        'title_template': title[:40] + ('...' if len(title) > 40 else ''),
        'body_guide': '；'.join(level0_labels[:5]) if level0_labels else f'参考原始文档结构替换内容'
    }

    # ── Try to find main_recipient ──
    defaults = {}
    for p in paragraphs[1:6]:
        if _RECIPIENT_RE.match(p) and len(p) < 80:
            defaults['main_recipient'] = p
            break

    if doc_type == '纪要':
        defaults.pop('main_recipient', None)

    # ── Assemble ──
    template = {
        "template_id": template_id,
        "name": name,
        "description": description,
        "document_type": doc_type,
        "source_refs": [
            {
                "type": "example",
                "note": f"从用户提供的{doc_type}文档自动提取，原始标题：{title[:40]}..."
            }
        ],
        "defaults": defaults,
        "body_sections": body_sections,
        "hints": hints,
    }

    return template


def parse_docx(path):
    """Parse a .docx file into template structure.

    Uses python-docx to read paragraphs and font information.
    Font-based heuristic: 黑体 paragraphs → level 0, 楷体 → level 1.
    Falls back to text-based pattern detection.

    Args:
        path: Path to .docx file.

    Returns:
        dict with template data (same structure as parse_text).
    """
    try:
        from docx import Document
    except ImportError:
        print("[WARN] python-docx not installed; falling back to raw text extraction.", file=sys.stderr)
        # Read as raw text
        import zipfile
        try:
            with zipfile.ZipFile(path) as z:
                # Extract document.xml and read text
                import xml.etree.ElementTree as ET
                xml_content = z.read('word/document.xml')
                root = ET.fromstring(xml_content)
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                texts = []
                for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                    p_texts = []
                    for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                        if t.text:
                            p_texts.append(t.text)
                    texts.append(''.join(p_texts))
                return parse_text('\n'.join(texts))
        except Exception:
            raise ValueError(f"Cannot read {path}: not a valid .docx or text file")

    doc = Document(path)
    paragraphs_data = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        # Try font-based heading detection
        font_name = None
        for run in para.runs:
            if run.font.name:
                font_name = run.font.name.strip()
                break

        level = _detect_level(text)

        # Font-based override
        if level is None:
            if font_name and ('黑体' in font_name):
                level = 0
            elif font_name and ('楷体' in font_name):
                level = 1

        paragraphs_data.append({
            'text': text,
            'level': level,
            'font': font_name,
        })

    # Reconstruct text for parsing
    text_lines = [p['text'] for p in paragraphs_data if p['text']]
    template = parse_text('\n'.join(text_lines))

    # Refine with font info: mark level 0/1 for paragraphs whose font
    # matches heading fonts but whose text pattern doesn't
    for i, section in enumerate(template['body_sections']):
        text = section['text']
        # Find matching paragraph data
        for pd in paragraphs_data:
            if pd['text'] == text:
                if pd['level'] is not None and section['level'] is None:
                    section['level'] = pd['level']
                break

    return template


# ── CLI ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Parse a Chinese official document into a template JSON file.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'Examples:\n'
            '  cat doc.txt | python create_template.py\n'
            '  python create_template.py -i document.docx -o my-template.json\n'
            '  python create_template.py --input doc.txt --preview\n'
        )
    )
    parser.add_argument('--input', '-i', type=str,
                        help='Path to input document (.txt or .docx). If omitted, reads from stdin.')
    parser.add_argument('--output', '-o', type=str,
                        help='Output template JSON path. If omitted, prints to stdout.')
    parser.add_argument('--preview', '-p', action='store_true',
                        help='Show summary preview (template id, name, section count) instead of full JSON.')
    parser.add_argument('--quiet', '-q', action='store_true',
                        help='Suppress info messages; only output JSON.')
    args = parser.parse_args()

    # ── Read input ──
    if args.input:
        ext = os.path.splitext(args.input)[1].lower()
        if ext == '.docx':
            if not args.quiet:
                print(f"[INFO] Parsing .docx: {args.input}", file=sys.stderr)
            template = parse_docx(args.input)
        else:
            with open(args.input, 'r', encoding='utf-8') as f:
                text = f.read()
            template = parse_text(text)
    else:
        # stdin
        if not sys.stdin.isatty() or args.quiet:
            text = sys.stdin.read()
        else:
            print("Paste document text (Ctrl+Z then Enter to end, or Ctrl+C to cancel):",
                  file=sys.stderr)
            text = sys.stdin.read()
        template = parse_text(text)

    # ── Output ──
    if args.preview:
        print(f"\nTemplate ID:    {template['template_id']}")
        print(f"Name:           {template['name']}")
        print(f"Document type:  {template['document_type']}")
        print(f"Sections:       {len(template['body_sections'])}")
        levels = {}
        for s in template['body_sections']:
            lvl = s['level'] if s['level'] is not None else 'body'
            levels[lvl] = levels.get(lvl, 0) + 1
        print(f"Breakdown:      {', '.join(f'{k}={v}' for k, v in sorted(levels.items(), key=lambda x: str(x[0])))}")
        if template.get('defaults', {}).get('main_recipient'):
            print(f"Recipient:      {template['defaults']['main_recipient']}")
        print(f"\nSource:         {template['source_refs'][0]['note'][:60]}...")
        return

    json_str = json.dumps(template, ensure_ascii=False, indent=2)

    if args.output:
        os.makedirs(os.path.dirname(args.output) or '.', exist_ok=True)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(json_str)
        if not args.quiet:
            print(f"[OK] Template saved to: {os.path.abspath(args.output)}", file=sys.stderr)
            print(f"     template_id: {template['template_id']}", file=sys.stderr)
            print(f"     sections: {len(template['body_sections'])}", file=sys.stderr)
    else:
        print(json_str)


if __name__ == '__main__':
    main()
