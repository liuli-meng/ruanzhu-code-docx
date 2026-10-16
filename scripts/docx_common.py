# -*- coding: utf-8 -*-
"""软著源代码 docx 的读写公共逻辑。

关键约定（踩坑总结，勿改）：
- 这类文档通常只有 **1 个 <w:p> 段落**，全文靠成千上万个 <w:br/> 换行。
- 缩进不是字面 TAB 字符，而是 <w:tab/> 元素。提取时**必须**处理 tab，
  否则会误判"缩进全部丢失"。
- 写回时 <w:br/> 一律**后置**（文本在前、换行在后），末行不带 br；
  前置会在整篇开头凭空多出一个空行。
"""

import os
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
DOC_XML = 'word/document.xml'

# 默认字体：宋体五号（sz 值为半磅，21 = 10.5pt）
FONT_INNER = ('<w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:cs="宋体"/>'
              '<w:sz w:val="21"/><w:szCs w:val="21"/>')


# ---------------------------------------------------------------- 读取

def read_doc_xml(path):
    """读出 document.xml 文本。"""
    with zipfile.ZipFile(path) as z:
        return z.read(DOC_XML).decode('utf-8')


def extract_lines(doc_xml):
    """把 document.xml 还原成逻辑行列表（含缩进的 \\t）。"""
    body = ET.fromstring(doc_xml).find(W + 'body')
    if body is None:
        raise ValueError('document.xml 里找不到 <w:body>')
    lines, cur = [], []
    for r in body.iter(W + 'r'):
        for ch in r:
            tag = ch.tag.replace(W, '')
            if tag in ('br', 'cr'):
                lines.append(''.join(cur))
                cur = []
            elif tag == 't':
                cur.append(ch.text or '')
            elif tag == 'tab':
                cur.append('\t')
    lines.append(''.join(cur))
    return lines


def doc_info(path):
    """返回 dict：段落数、br 数、sectPr 数。"""
    d = read_doc_xml(path)
    return {
        'paras': len(__import__('re').findall(r'<w:p[ />]', d)),
        'brs': d.count('<w:br'),
        'sectPr': d.count('<w:sectPr'),
    }


def part_text(path, part):
    """取某个部件（如 word/header2.xml）里的纯文本（自动剔除 PAGE 域）。

    注意两点：
    1. 正则必须写 `<w:t(?:\\s[^>]*)?>`，写成 `<w:t[^>]*>` 会误匹配 `<w:tabs>`。
    2. 必须剔除域（fldChar begin…end）段，否则页码的缓存结果会被当成
       页眉文字的一部分，页眉看起来像 "XX系统 V1.01"。
    """
    with zipfile.ZipFile(path) as z:
        try:
            xml = z.read(part).decode('utf-8')
        except KeyError:
            return None
    xml = _strip_fields(xml)
    import re
    return ''.join(re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', xml, re.S))


def _strip_fields(xml):
    """移除 fldChar begin…end 之间的整段（含缓存结果）。"""
    import re
    while True:
        b = xml.find('fldCharType="begin"')
        if b < 0:
            break
        e = xml.find('fldCharType="end"', b)
        if e < 0:
            break
        rs = max(xml.rfind('<w:r ', 0, b), xml.rfind('<w:r>', 0, b))
        re_end = xml.find('</w:r>', e)
        if rs < 0 or re_end < 0:
            break
        xml = xml[:rs] + xml[re_end + 6:]
    return xml


def part_exists(path, part):
    with zipfile.ZipFile(path) as z:
        return part in z.namelist()


def header_footer_map(path):
    """解析 sectPr 里的 header/footer 引用 -> 具体部件名。

    返回 {'header': {'default': 'word/header2.xml', ...}, 'footer': {...}}
    """
    import re
    with zipfile.ZipFile(path) as z:
        doc = z.read(DOC_XML).decode('utf-8')
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
        settings = z.read('word/settings.xml').decode('utf-8')

    rid2target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    out = {'header': {}, 'footer': {}}
    for m in re.finditer(r'<w:(header|footer)Reference w:type="(\w+)" r:id="(rId\d+)"/>', doc):
        kind, typ, rid = m.group(1), m.group(2), m.group(3)
        target = rid2target.get(rid)
        if target and not target.startswith('/'):
            target = 'word/' + target.lstrip('/')
        out[kind][typ] = target
    out['evenAndOddHeaders'] = '<w:evenAndOddHeaders' in settings
    out['titlePg'] = '<w:titlePg' in doc
    return out


# ---------------------------------------------------------------- 写入

def build_para(lines, font_inner=FONT_INNER, line_spacing=None):
    """把行列表构造成单个 <w:p>。

    line_spacing: 固定行距，单位 twips（20 twips = 1pt）。None 表示不设置。
    """
    rpr = '<w:rPr>%s</w:rPr>' % font_inner
    runs = []
    for i, l in enumerate(lines):
        br = '<w:br/>' if i < len(lines) - 1 else ''
        runs.append('<w:r>%s<w:t xml:space="preserve">%s</w:t>%s</w:r>'
                    % (rpr, escape(l), br))
    sp = ''
    if line_spacing:
        sp = ('<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/>'
              % int(line_spacing))
    return '<w:p><w:pPr>%s%s</w:pPr>%s</w:p>' % (sp, rpr, ''.join(runs))


def replace_body(doc_xml, para_xml):
    """用 para_xml 替换 document.xml 的 body 内容，保留 sectPr。"""
    import re
    m = re.search(r'(<w:body>)(.*?)(<w:sectPr)', doc_xml, re.S)
    if not m:
        raise ValueError('定位 <w:body>…<w:sectPr> 失败')
    return doc_xml[:m.end(1)] + para_xml + doc_xml[m.start(3):]


def save_with(src_path, dst_path, parts):
    """以 src 为模板打包出 dst，parts 形如 {'word/document.xml': '<xml>'}。"""
    zin = zipfile.ZipFile(src_path)
    zout = zipfile.ZipFile(dst_path, 'w', zipfile.ZIP_DEFLATED)
    try:
        for it in zin.infolist():
            data = parts.get(it.filename)
            if data is None:
                data = zin.read(it.filename)
            elif isinstance(data, str):
                data = data.encode('utf-8')
            zout.writestr(it, data)
    finally:
        zout.close()
        zin.close()
    return dst_path


# ---------------------------------------------------------------- 页面几何

def page_geometry(doc_xml):
    """从 sectPr 读页宽/页高/页边距（twips）。"""
    import re
    sect = re.search(r'<w:sectPr.*?</w:sectPr>', doc_xml, re.S)
    sect = sect.group(0) if sect else doc_xml
    def grab(tag, attr):
        m = re.search(r'<w:%s[^>]*w:%s="(-?\d+)"' % (tag, attr), sect)
        return int(m.group(1)) if m else None
    return {
        'pw': grab('pgSz', 'w'), 'ph': grab('pgSz', 'h'),
        'top': grab('pgMar', 'top'), 'bottom': grab('pgMar', 'bottom'),
        'left': grab('pgMar', 'left'), 'right': grab('pgMar', 'right'),
    }


def has_line_numbers(doc_xml):
    return '<w:lnNumType' in doc_xml


# ---------------------------------------------------------------- Word COM

def _word_app():
    import win32com.client as win32
    app = win32.Dispatch('Word.Application')
    app.Visible = False
    app.DisplayAlerts = 0
    return app


def word_stats(path, pdf_path=None):
    """用 Word 实测页数/行数/行距等。需要 pywin32 与本机 Word。

    统计与导出 PDF 在**同一个 Word 会话**内完成——分两次 Dispatch 时，
    上一次 Quit 尚未结束会让新实例取到正在销毁的对象，
    表现为 `Open.ComputeStatistics` / `Open.ExportAsFixedFormat` 属性不存在。
    """
    app = None
    doc = None
    try:
        app = _word_app()
        doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True)
        ps = doc.Sections(1).PageSetup
        pf = doc.Paragraphs(1).Format
        st = {
            'pages': doc.ComputeStatistics(2),
            'lines': doc.ComputeStatistics(1),
            'chars_with_space': doc.ComputeStatistics(5),
            'paragraphs': doc.Paragraphs.Count,
            'line_numbering_active': int(ps.LineNumbering.Active),
            'line_spacing': round(float(pf.LineSpacing), 2),
            'line_spacing_rule': int(pf.LineSpacingRule),   # 4 = 固定值
            'top_cm': round(ps.TopMargin / 72 * 2.54, 2),
            'bottom_cm': round(ps.BottomMargin / 72 * 2.54, 2),
            'left_cm': round(ps.LeftMargin / 72 * 2.54, 2),
            'right_cm': round(ps.RightMargin / 72 * 2.54, 2),
            'header_text': doc.Sections(1).Headers(1).Range.Text.replace('\r', ''),
            'footer_text': doc.Sections(1).Footers(1).Range.Text.replace('\r', ''),
        }
        if pdf_path:
            doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)  # 17 = PDF
        return st
    finally:
        try:
            if doc is not None:
                doc.Close(False)
            if app is not None:
                app.Quit()
        except Exception:
            pass


def word_export_pdf(path, pdf_path):
    """单独导出 PDF（内部复用 word_stats 的单会话逻辑）。"""
    word_stats(path, pdf_path=pdf_path)
    return pdf_path


def kill_word():
    """兜底清掉残留的 WINWORD 进程（COM 偶发退不干净，会留 ~$ 锁文件）。"""
    os.system('taskkill /F /IM WINWORD.EXE > nul 2>&1')
