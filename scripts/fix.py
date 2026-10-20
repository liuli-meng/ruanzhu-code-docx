#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""对已有软著源代码文档做定点修复（不动其余内容）。

用法示例：
    python fix.py 软著代码.docx --out 软著代码_优化版.docx \
        --strip-text 2097632843@qq.com --strip-date-comments \
        --no-line-numbers --page-number header-right --trim-blank-lines \
        --fit-pages 60

设计原则：所有改动都在 XML 层做**定点替换/删除**，不重建正文，
因此除你指定的内容外，源码内容一字不动。
"""

import argparse
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_common as C   # noqa: E402
import hf                 # noqa: E402

DATE_COMMENT = re.compile(r'^\s*(?://|#|\*|/\*+)\s*\d{4}-\d{2}-\d{2}\s*\*?/?\s*$')


# ------------------------------------------------------------ 定点操作

def replace_in_texts(xml, old, new):
    """在 <w:t> 内替换文本。返回 (xml, 命中次数)。"""
    n = [0]

    def sub(m):
        txt = m.group(2)
        if old in txt:
            n[0] += txt.count(old)
            txt = txt.replace(old, new)
        return m.group(1) + txt + m.group(3)
    xml = re.sub(r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)', sub, xml, flags=re.S)
    return xml, n[0]


def delete_run_containing(xml, needle, warn=None):
    """删除包含 needle 的那个 <w:r>（假定一行一个 run）。"""
    i = xml.find(needle)
    if i < 0:
        return xml, False
    rs = max(xml.rfind('<w:r ', 0, i), xml.rfind('<w:r>', 0, i))
    if rs < 0:
        return xml, False
    re_end = xml.find('</w:r>', i)
    if re_end < 0:
        return xml, False
    run = xml[rs:re_end + 6]
    others = re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', run, re.S)
    joined = ''.join(others)
    if needle not in joined or joined.strip() != needle.strip():
        if warn:
            warn('该 run 除目标内容外还包含其他文本，已跳过：%r' % joined[:60])
        return xml, False
    return xml[:rs] + xml[re_end + 6:], True


def drop_trailing_blank_runs(xml, limit=50):
    """删除段落末尾只含 <w:br/> 的 run（即末尾空行）。"""
    removed = 0
    for _ in range(limit):
        m = None
        for mm in re.finditer(r'<w:r\b[^>]*>(?:(?!</w:r>).)*?</w:r>(?=\s*</w:p>)', xml, re.S):
            m = mm
        if not m:
            break
        run = m.group(0)
        if '<w:t' in run or '<w:tab' in run:
            break
        if '<w:br/>' not in run and '<w:br' not in run:
            break
        xml = xml[:m.start()] + xml[m.end():]
        removed += 1
    return xml, removed


def remove_line_numbers(xml):
    n = len(re.findall(r'<w:lnNumType[^>]*/>', xml))
    return re.sub(r'<w:lnNumType[^>]*/>', '', xml), n


def uniform_lines(xml):
    """关闭全部段落的孤行控制（widowControl），让每页行数一致。

    现象：每页本应 51 行，却有个别页只有 50 行。
    根因：这些页末尾是**折行的长代码行**，Word 的「孤行控制」会把整段
    （2 个视觉行）一起推到下一页，于是本页少一行、下一页多一行。
    关掉 widowControl 后，折行可以自由跨页，每页行数就统一了。

    实测：某 54 页文档 {50×2, 51×51} -> {51×53}（末页 33 行），页数不变。
    注意：对代码清单来说断行跨页完全可接受，反而更整齐。

    返回 (新 xml, 修改的段落数)。
    """
    cnt = [0]

    def repl(m):
        inner = m.group(1)
        if 'widowControl' in inner:
            return m.group(0)
        cnt[0] += 1
        ins = '<w:widowControl w:val="0"/>'
        # 按 OOXML 的元素顺序，widowControl 应排在 spacing / ind / jc / rPr 之前
        for anchor in ('<w:spacing', '<w:ind', '<w:jc', '<w:rPr>'):
            i = inner.find(anchor)
            if i >= 0:
                return '<w:pPr>' + inner[:i] + ins + inner[i:] + '</w:pPr>'
        return '<w:pPr>' + ins + inner + '</w:pPr>'

    return re.sub(r'<w:pPr>(.*?)</w:pPr>', repl, xml, flags=re.S), cnt[0]


def set_tabs(xml, tab_twips, count=24):
    """把段落制表位固定成指定宽度（twips 一个制表位）。

    为什么需要：这类源码文档的缩进是 `<w:tab/>`，而中文 Word 的**默认制表位
    是每 4 个字符一跳**。8pt 时 12 个 tab 只占 384pt（放得下），
    换成五号 10.5pt 就变成 504pt——直接顶满整行，Word 会把这一排 tab
    挤成**一个空行**，正文反而掉到下一行左侧。表现为每页凭空少掉几行，
    跌破「每页 50 行」，且肉眼看上去是诡异的空白块。
    显式给每 2 个字符一个制表位，缩进比例与 8pt 时基本一致，也不会再挤爆。
    """
    tabs = '<w:tabs>%s</w:tabs>' % ''.join(
        '<w:tab w:val="left" w:pos="%d"/>' % int(tab_twips * (i + 1))
        for i in range(count))
    xml = re.sub(r'<w:tabs>.*?</w:tabs>', '', xml, flags=re.S)
    m = re.search(r'<w:pPr\b[^>]*?/>', xml)
    if m:                                    # 自闭合 pPr → 展开
        head, tail = xml[:m.start()], xml[m.end():]
        return head + '<w:pPr>%s</w:pPr>' % tabs + tail
    m = re.search(r'<w:pPr\b[^>]*>', xml)
    if not m:
        return xml
    # tabs 在 pPr 里必须排在 spacing / ind / jc / rPr 之前
    return xml[:m.end()] + tabs + xml[m.end():]


def set_spacing(xml, pt):
    """设定正文段落的固定行距（twips = pt*20）。

    注意：段落属性常被写成自闭合的 `<w:pPr />`（从 HTML/文本导入的软著文档
    几乎都是这样，正文只有 1 个段落）。原实现只找 `<w:pPr>`，
    遇到自闭合写法会**静默不改**——行距没生效，页数自然也不会变。
    """
    tw = int(round(pt * 20))
    new_sp = ('<w:spacing w:before="0" w:after="0" w:line="%d" w:lineRule="exact"/>' % tw)
    xml = re.sub(r'<w:spacing[^>]*w:line="\d+"[^>]*w:lineRule="exact"\s*/>', '', xml)

    # 自闭合 <w:pPr ... />：直接展开成带 spacing 的完整 pPr
    m = re.search(r'<w:pPr\b([^>]*?)\s*/>', xml)
    if m:
        head, tail = xml[:m.start()], xml[m.end():]
        return head + '<w:pPr%s>%s</w:pPr>' % (m.group(1), new_sp) + tail

    m = re.search(r'<w:pPr\b[^>]*>', xml)
    if not m:
        return xml
    inner_start = m.end()
    inner_end = xml.find('</w:pPr>', inner_start)
    if inner_end < 0:
        return xml
    inner = xml[inner_start:inner_end]
    # spacing 的合法位置：pBdr / tabs / shd 之后，ind / jc / rPr 之前
    for anchor in ('<w:ind', '<w:jc', '<w:rPr>'):
        i = inner.find(anchor)
        if i >= 0:
            return xml[:inner_start] + inner[:i] + new_sp + inner[i:] + xml[inner_end:]
    return xml[:inner_start] + inner + new_sp + xml[inner_end:]


# 注意 `<w:sz` 后面必须紧跟空白或 / >，否则会连 `<w:szCs` 一起吃进去
_SZ = r'<w:sz(?=[\s/>])[^>]*?(?:/>|>\s*</w:sz>)'
_SZCS = r'<w:szCs(?=[\s/>])[^>]*?(?:/>|>\s*</w:szCs>)'
_RFONTS = r'<w:rFonts\b[^>]*?(?:/>|>\s*</w:rFonts>)'


def set_font(xml, font, size_pt=None):
    """统一字体 / 字号。返回 (新 xml, {改动统计})。

    size_pt 用磅；写进 XML 时 ×2（半磅）。`w:sz` 与 `w:szCs` 都要写，
    只写一个在换软件/换打印驱动时会变形。
    """
    stat = {}
    if font:
        new = ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/>'
               % ((font,) * 4))
        xml, stat['rFonts'] = re.subn(_RFONTS, new, xml)
    if size_pt:
        half = int(round(size_pt * 2))
        xml, n1 = re.subn(_SZ, '<w:sz w:val="%d"/>' % half, xml)
        xml, n2 = re.subn(_SZCS, '<w:szCs w:val="%d"/>' % half, xml)
        stat['sz'] = n1
        stat['szCs'] = n2
    return xml, stat


# ------------------------------------------------------------ 主流程

def fix(inp, out, strip_texts=(), strip_dates=False, no_line_numbers=False,
        page_number=None, header_name=None, trim_blank=False,
        line_spacing=None, fit_pages=None, uniform=False, font=None,
        font_size=None, tab_chars=None, verbose=True):
    with zipfile.ZipFile(inp) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    doc = parts[C.DOC_XML].decode('utf-8')
    changes = []
    warns = []

    def warn(msg):
        warns.append(msg)
        if verbose:
            print('  ! ' + msg)

    if line_spacing:
        doc = set_spacing(doc, line_spacing)
        changes.append('固定行距 %.1fpt' % line_spacing)

    if tab_chars and font_size:
        tw = int(round(tab_chars * font_size * 20))
        doc = set_tabs(doc, tw)
        changes.append('制表位固定为 %g 字符（%d twips），避免缩进把整行挤爆'
                       % (tab_chars, tw))

    if font or font_size:
        doc, st = set_font(doc, font, font_size)
        changes.append('统一正文字体 %s%s（rFonts %d 处 / sz %d 处 / szCs %d 处）'
                       % (font or '（不变）',
                          ' %gpt' % font_size if font_size else '',
                          st.get('rFonts', 0), st.get('sz', 0), st.get('szCs', 0)))

    for s in strip_texts:
        doc, n = replace_in_texts(doc, s, '')
        if n:
            changes.append('清除文本 %r（%d 处）' % (s, n))
        else:
            warn('未找到文本 %r' % s)

    if strip_dates:
        lines = C.extract_lines(doc)
        hits = [l.strip() for l in lines if DATE_COMMENT.match(l)]
        done = 0
        for h in set(hits):
            doc, ok = delete_run_containing(doc, h, warn)
            done += 1 if ok else 0
        if done:
            changes.append('删除日期注释行 %d 个' % done)

    if trim_blank:
        doc, n = drop_trailing_blank_runs(doc)
        if n:
            changes.append('删除末尾空行 %d 个' % n)

    if no_line_numbers:
        doc, n = remove_line_numbers(doc)
        changes.append('移除行号设置 %d 处' % n)

    if uniform:
        doc, n = uniform_lines(doc)
        changes.append('关闭孤行控制（%d 个段落），每页行数统一' % n)

    parts[C.DOC_XML] = doc.encode('utf-8')

    # 页眉 / 页脚
    hfm = C.header_footer_map(inp)
    geo = C.page_geometry(doc)
    body_w = (geo['pw'] or 11906) - (geo['left'] or 1440) - (geo['right'] or 1440)

    # 页眉 / 页脚。注意：page_number 指定的是「页码应该在哪儿」，
    # 因此必须把**另一侧**的页码一并清掉，否则会出现同页两个页码。
    hp = hfm['header'].get('default')
    fp = hfm['footer'].get('default')
    if header_name:
        hdr_name = header_name
    elif hp:
        hdr_name = (C.part_text(inp, hp) or '').strip()
    else:
        hdr_name = ''

    if page_number and hp:
        hx = parts[hp].decode('utf-8')
        right = (page_number == 'header-right')
        # 只在「需要动页眉」时才重建：
        #   - 用户指定了 --header-name，或
        #   - 页眉里确实有页码域需要清掉，或
        #   - 目标是把页码放到页眉右上角
        # 否则重建属于多余改动，会改掉原有的字体/字号（实测把宋体 10.5pt
        # 改成了微软雅黑 8pt）。
        need = bool(header_name) or right or ('fldCharType="begin"' in hx)
        if need:
            parts[hp] = hf.set_header(hx, hdr_name, body_w, page_at_right=right).encode('utf-8')
            changes.append('页眉重建：%s（%s，保留原标题样式）'
                           % (hdr_name, '含右上角页码' if right else '仅软件名，无页码'))
        else:
            changes.append('页眉未改动（本来就没有页码，保留原样式）')
    elif header_name and not hp:
        warn('没有默认页眉部件，无法写入软件名称')

    if page_number == 'header-right':
        if fp:
            parts[fp] = hf.set_footer(parts[fp].decode('utf-8'), page_number=False).encode('utf-8')
            changes.append('清空页脚页码')
    elif page_number == 'footer-center':
        if fp:
            parts[fp] = hf.set_footer(parts[fp].decode('utf-8'), page_number=True).encode('utf-8')
            changes.append('页脚居中放页码')
        else:
            warn('没有默认页脚部件，无法放置页码')
    elif page_number == 'none':
        if fp:
            parts[fp] = hf.set_footer(parts[fp].decode('utf-8'), page_number=False).encode('utf-8')
            changes.append('清空页脚页码')

    os.makedirs(os.path.dirname(os.path.abspath(out)) or '.', exist_ok=True)

    # 页眉/页脚里的软件名、页码也统一字体（否则会出现「正文宋体五号、
    # 页眉微软雅黑 8pt」的混搭）。放在页眉页脚重建之后执行。
    if font or font_size:
        for name in list(parts):
            if not re.match(r'^word/(header|footer)\d*\.xml$', name):
                continue
            x, st = set_font(parts[name].decode('utf-8'), font, font_size)
            parts[name] = x.encode('utf-8')

    # 自动反推行距，保证页数不超过 fit_pages 且每页 >= 50 行
    if fit_pages:
        best = _fit_spacing(inp, parts, fit_pages, verbose)
        if best is not None:
            tw, pages, per = best
            doc = parts[C.DOC_XML].decode('utf-8')
            doc = set_spacing(doc, tw / 20.0)
            parts[C.DOC_XML] = doc.encode('utf-8')
            changes.append('自动行距 %.1fpt → %d 页（每页 %d~%d 行）'
                           % (tw / 20.0, pages, min(per), max(per)))
        else:
            warn('自动匹配行距失败，沿用当前行距')

    C.save_with(inp, out, {k: v for k, v in parts.items()})
    if verbose:
        print('\n改动：')
        for c in changes:
            print('  · %s' % c)
        print('\n已生成：%s' % out)
    return out


def _fit_spacing(tpl, parts, max_pages, verbose=True):
    """反推行距：从「最宽松」往「最紧凑」试，取第一个满足
      总页数 <= max_pages 且每页 >= 50 行的行距。

    必须从大到小试——从小到大会挑到最挤的行距（行数虽多但排版难读）。
    """
    orig = parts[C.DOC_XML].decode('utf-8')
    import tempfile
    tmp = tempfile.mkdtemp(prefix='rz_fit_')
    cands = [15.5, 15.0, 14.5, 14.0, 13.5, 13.0, 12.5, 12.0, 11.5]
    try:
        for pt in cands:
            doc = set_spacing(orig, pt)
            p2 = dict(parts)
            p2[C.DOC_XML] = doc.encode('utf-8')
            probe = os.path.join(tmp, 'p.docx')
            C.save_with(tpl, probe, p2)
            try:
                st = C.word_stats(probe)
                pages = st['pages']
                per = None
                try:
                    import pymupdf
                    pdf = os.path.join(tmp, 'p.pdf')
                    C.word_export_pdf(probe, pdf)
                    per = _per_page(pdf, st['header_text'])
                    os.remove(pdf)
                except Exception:
                    pass
                if verbose:
                    print('  试 %.1fpt -> %d 页%s' % (
                        pt, pages,
                        '，每页 %d~%d 行' % (min(per), max(per)) if per else ''))
                ok_lines = (per is None) or (min(per[:-1] or per) >= 50)
                if pages <= max_pages and ok_lines:
                    return int(pt * 20), pages, (per or [0])
            except Exception:
                continue
        return None
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
        C.kill_word()


def _per_page(pdf_path, header_text):
    import pymupdf
    d = pymupdf.open(pdf_path)
    out = []
    for page in d:
        n = 0
        for blk in page.get_text('dict')['blocks']:
            if blk['type'] != 0:
                continue
            for ln in blk['lines']:
                s = ''.join(sp['text'] for sp in ln['spans']).strip()
                if not s:
                    continue
                if header_text and header_text.strip() and header_text.strip() in s:
                    continue
                if re.fullmatch(r'\d+', s) and ln['bbox'][0] < 90:
                    continue
                n += 1
        out.append(n)
    d.close()
    return out


def main():
    ap = argparse.ArgumentParser(description='软著源代码文档定点修复')
    ap.add_argument('docx')
    ap.add_argument('--out', required=True)
    ap.add_argument('--strip-text', action='append', default=[],
                    help='清除指定文本（可多次），如邮箱')
    ap.add_argument('--strip-date-comments', action='store_true',
                    help='删除形如 /** 2025-09-14 */ 的孤立日期注释行')
    ap.add_argument('--no-line-numbers', action='store_true')
    ap.add_argument('--page-number', choices=['header-right', 'footer-center', 'keep'])
    ap.add_argument('--header-name', help='写入/覆盖页眉的软件名称')
    ap.add_argument('--trim-blank-lines', action='store_true')
    ap.add_argument('--uniform-lines', action='store_true',
                    help='关闭孤行控制，使每页行数一致（解决个别页少一行的折行问题）')
    ap.add_argument('--line-spacing', type=float, help='固定行距(pt)')
    ap.add_argument('--font', help='统一正文字体（含页眉页脚），如 宋体')
    ap.add_argument('--font-size', type=float,
                    help='统一正文字号(pt)，软著源码文档一般为 10.5（五号）')
    ap.add_argument('--tab-chars', type=float,
                    help='把 <w:tab/> 制表位固定成 N 个字符宽（常用 2）。'
                         '字号调大后缩进会挤爆整行、每页凭空少行，用这个修')
    ap.add_argument('--fit-pages', type=int,
                    help='自动反推行距，使总页数不超过该值且每页 >= 50 行')
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print('文件不存在：%s' % args.docx)
        return 2
    fix(args.docx, args.out, strip_texts=args.strip_text,
        strip_dates=args.strip_date_comments, no_line_numbers=args.no_line_numbers,
        page_number=args.page_number, header_name=args.header_name,
        trim_blank=args.trim_blank_lines, line_spacing=args.line_spacing,
        fit_pages=args.fit_pages, uniform=args.uniform_lines,
        font=args.font, font_size=args.font_size, tab_chars=args.tab_chars)
    print('\n下一步：python audit.py "%s"' % args.out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
