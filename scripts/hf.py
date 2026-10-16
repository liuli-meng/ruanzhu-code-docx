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


def _replace_last_paragraph(xml, new_p):
    spans = _paragraphs(xml)
    if not spans:
        return xml.replace('</w:hdr>', new_p + '</w:hdr>').replace(
            '</w:ftr>', new_p + '</w:ftr>')
    s, e = spans[-1]
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
    return _replace_last_paragraph(
        xml, build_header(name, body_width_twips, page_at_right))


def set_footer(xml, page_number=False):
    return _replace_last_paragraph(
        xml, build_footer_page() if page_number else build_footer_empty())


def strip_page_field(xml):
    """安全地移除页脚里的 PAGE 域（保留段落）。"""
    xml = re.sub(r'<w:r>(?:(?!</w:r>).)*?<w:fldChar[^>]*/>(?:(?!</w:r>).)*?</w:r>', '', xml, flags=re.S)
    xml = re.sub(r'<w:r>(?:(?!</w:r>).)*?<w:instrText[^>]*>(?:(?!</w:r>).)*?</w:r>', '', xml, flags=re.S)
    return xml


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
