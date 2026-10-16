# -*- coding: utf-8 -*-
"""页眉 / 页脚重建。

踩坑记录（务必遵守）：
- `<w:tabs>` 的子元素是 `<w:tab w:val="center" w:pos="N"/>`，
  **不是** `<w:center>` / `<w:right>`。写成后者 Word 会静默忽略，
  页眉内容会贴着左边距排，且不报错——只能靠导出 PDF 看 x 坐标发现。
- tab stop 的 pos 以正文宽度（页宽 − 左右边距）为基准：
    居中位 = 正文宽 / 2      右对齐位 = 正文宽
- 页眉内容高度不变才不会挤压正文行数；因此把页码放在**同一行**，
  不要新增段落。
"""

import re

DEFAULT_HF_FONT = ('<w:rFonts w:ascii="微软雅黑" w:eastAsia="微软雅黑" '
                   'w:hAnsi="微软雅黑" w:cs="微软雅黑"/>'
                   '<w:sz w:val="16"/><w:szCs w:val="16"/>'
                   '<w:lang w:eastAsia="zh-CN"/>')

PAGE_FIELD = (
    '<w:r>{r}<w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r>{r}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r>{r}<w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r>{r}<w:t>1</w:t></w:r>'
    '<w:r>{r}<w:fldChar w:fldCharType="end"/></w:r>'
)


def _paragraphs(xml):
    """匹配段落。必须同时覆盖 `<w:p .../>`（自闭合，无 </w:p>），
    否则正则会把自闭合段落一路吞到下一个 </w:p>，误删中间的段落。"""
    return [(m.start(), m.end())
            for m in re.finditer(r'<w:p\b[^>]*/>|<w:p\b[^>]*>.*?</w:p>', xml, re.S)]


def _content_bounds(xml):
    """返回 (根标签结束位置, 根闭合标签起始位置)。"""
    m = re.search(r'<(?:w:hdr|w:ftr)\b[^>]*>', xml)
    if not m:
        raise ValueError('不是 header/footer 部件')
    c = re.search(r'</(?:w:hdr|w:ftr)>\s*$', xml)
    end = c.start() if c else len(xml)
    return m.end(), end


def _run_start(xml, pos):
    """pos 之前最近的 <w:r> / <w:r ...> 起始位置。"""
    best = -1
    for m in re.finditer(r'<w:r(?=[ >])', xml[:pos]):
        best = m.start()
    return best


def _replace_last_text_paragraph(xml, new_p):
    """替换**最后一个含文本的段落**，保留其前面的段落。

    页眉里常有一个前置空段落撑高度，直接删会改变正文可用高度、进而改变页数；
    只替换含文本的那段最安全。
    """
    spans = _paragraphs(xml)
    target = None
    for s, e in reversed(spans):
        if '<w:t' in xml[s:e]:
            target = (s, e)
            break
    if target is None:
        if not spans:
            return xml
        target = spans[-1]
    s, e = target
    return xml[:s] + new_p + xml[e:]


def build_header(name, body_width_twips, page_at_right=True,
                 font=DEFAULT_HF_FONT, bottom_border=True):
    """软件名居中（+ 页码右上），单段落、单行。"""
    r = '<w:rPr>%s</w:rPr>' % font
    tabs = ''
    inner = ''
    if page_at_right:
        tabs = ('<w:tabs><w:tab w:val="center" w:pos="%d"/>'
                '<w:tab w:val="right" w:pos="%d"/></w:tabs>'
                % (body_width_twips // 2, body_width_twips))
        inner = '<w:r>%s<w:tab/></w:r>' % r
    else:
        tabs = ('<w:tabs><w:tab w:val="center" w:pos="%d"/></w:tabs>'
                % (body_width_twips // 2))
        inner = '<w:r>%s<w:tab/></w:r>' % r

    border = ''
    if bottom_border:
        border = ('<w:pBdr><w:bottom w:val="single" w:sz="1" w:space="1" '
                  'w:color="000000"/></w:pBdr>')

    parts = ['<w:p><w:pPr>', border, tabs, '<w:rPr>%s</w:rPr>' % font, '</w:pPr>',
             inner,
             '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (r, _esc(name))]
    if page_at_right:
        parts.append('<w:r>%s<w:tab/></w:r>' % r)
        parts.append(PAGE_FIELD.format(r=r))
    parts.append('</w:p>')
    return ''.join(parts)


def build_footer_empty():
    return '<w:p><w:pPr><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:pPr></w:p>'


def build_footer_page(font=DEFAULT_HF_FONT):
    r = '<w:rPr>%s</w:rPr>' % font
    return ('<w:p><w:pPr><w:jc w:val="center"/><w:rPr>%s</w:rPr></w:pPr>%s</w:p>'
            % (font, PAGE_FIELD.format(r=r)))


def set_header(xml, name, body_width_twips, page_at_right=True):
    xml = strip_page_field(xml)          # 先清掉页眉里可能已有的页码域
    return _replace_last_text_paragraph(
        xml, build_header(name, body_width_twips, page_at_right))


def set_footer(xml, page_number=False):
    """整段重建页脚内容，**只留一个段落**。

    不能只替换「最后一个段落」：Word 常把页码包在内容控件
    `<w:sdt><w:sdtContent>…</w:sdtContent></w:sdt>` 里，这样会替换到
    边上的空段落，原本的页码留在 SDT 中 → 同一页出现两个重叠的页码。
    直接重建整个 body 才能保证有且只有一个页码。
    """
    s, e = _content_bounds(xml)
    body = build_footer_page() if page_number else build_footer_empty()
    return xml[:s] + body + xml[e:]


def strip_page_field(xml):
    """移除完整的 PAGE 域（begin…end，含中间的缓存结果 run）。"""
    while True:
        b = xml.find('fldCharType="begin"')
        if b < 0:
            break
        e = xml.find('fldCharType="end"', b)
        if e < 0:
            break
        rs = _run_start(xml, b)
        re_end = xml.find('</w:r>', e)
        if rs < 0 or re_end < 0:
            break
        xml = xml[:rs] + xml[re_end + 6:]
    return xml


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
