#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""软著源代码文档合规体检。

用法：
    python audit.py <软著代码.docx> [--src 源码目录] [--no-word] [--json]

输出：按 FAIL / WARN / INFO 分级的问题清单 + 实测数据。
退出码：0 = 无 FAIL；1 = 存在 FAIL。
"""

import argparse
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_common as C   # noqa: E402

# 敏感信息模式：(正则, 名称, 级别)
PATTERNS = [
    (r'[\w.+-]+@[\w-]+\.[\w.]+', '邮箱地址', 'WARN'),
    (r'\b1[3-9]\d{9}\b', '手机号', 'WARN'),
    (r'[A-Za-z]:\\\\[^\s"\'<>]*', 'Windows 本地路径', 'WARN'),
    (r'(?:版权|版权所有|Copyright|\(c\)\s*\d{4})', '版权声明', 'WARN'),
    (r'@author|@Author|作者[：:]', '作者署名标记', 'WARN'),
    (r'(?:仅供|内部资料|保密|机密)', '保密字样', 'WARN'),
    (r'sk-[A-Za-z0-9]{16,}', '疑似 API Key', 'WARN'),
    (r'Bearer\s+[A-Za-z0-9._-]{24,}', '疑似 Token', 'WARN'),
    (r'\bhttps?://[^\s"\'<>)},]+', 'URL', 'INFO'),
    (r'\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', 'IP 地址', 'INFO'),
]

# 代码注释里的孤立日期（编辑器自动写入的噪音）
DATE_COMMENT = re.compile(r'^\s*(?://|#|\*|/\*+)\s*\d{4}-\d{2}-\d{2}\s*\*?/?\s*$')


def sev(v):
    return {'FAIL': 0, 'WARN': 1, 'INFO': 2}.get(v, 9)


def audit(path, src_dir=None, use_word=True):
    issues = []          # (level, title, detail)
    data = {}

    def add(level, title, detail=''):
        issues.append((level, title, detail))

    # ------------------------------ 包完整性
    try:
        z = zipfile.ZipFile(path)
        if z.testzip() is not None:
            add('FAIL', 'docx 包损坏', 'ZIP 内部校验失败')
        names = z.namelist()
        if C.DOC_XML not in names:
            add('FAIL', '缺少 word/document.xml', '不是有效的 docx')
            return issues, data
        doc = z.read(C.DOC_XML).decode('utf-8')
        try:
            C.EXTRA_CHECK = None
            import xml.etree.ElementTree as ET
            ET.fromstring(doc)
        except Exception as e:
            add('FAIL', 'document.xml 解析失败', str(e))
            return issues, data
    except Exception as e:
        add('FAIL', '无法读取文件', str(e))
        return issues, data

    # ------------------------------ 内容
    lines = C.extract_lines(doc)
    nonblank = [l for l in lines if l.strip()]
    data['lines_total'] = len(lines)
    data['lines_nonblank'] = len(nonblank)
    data['lines_blank'] = len(lines) - len(nonblank)

    if data['lines_blank'] > 0:
        tail_blank = 0
        for l in reversed(lines):
            if l.strip():
                break
            tail_blank += 1
        add('WARN', '文档含空行 %d 个' % data['lines_blank'],
            '空行不计有效代码，会稀释每页行数；末尾空行 %d 个（应删）' % tail_blank)
    if lines and not lines[0].strip():
        add('WARN', '首行为空', '开头存在空行，观感差且浪费版面')

    data['lines_indented'] = sum(1 for l in nonblank if l[:1] in (' ', '\t'))
    if nonblank and data['lines_indented'] == 0:
        add('FAIL', '缩进疑似全部丢失', '没有任何行带前导空白，源码可读性/真实性存疑')

    longest = max((len(l) for l in lines), default=0)
    data['longest_line'] = longest
    if longest > 2000:
        add('WARN', '存在超长行（最长 %d 字符）' % longest,
            '压缩/生成的单行文件（min.js、GeoJSON 等）会渲染成上百页乱码，'
            '建议从源码目录排除后再生成')

    # HTML / 模板残留
    txt = '\n'.join(lines)
    if '&lt;!DOCTYPE' in doc or '<!DOCTYPE html>' in txt:
        add('WARN', '含 <!DOCTYPE html>', '确认这是源码本身，而非提取时混入的脚手架')
    if txt.count('</style>') > 2:
        add('INFO', '含 %d 处 </style>' % txt.count('</style>'), 'Vue/SFC 组件的样式块属正常代码')

    # ------------------------------ 排版
    geo = C.page_geometry(doc)
    data['page_setup'] = geo
    marg = [v for v in (geo['top'], geo['left']) if v]
    if marg and min(marg) < 850:   # < 1.5cm
        add('WARN', '页边距偏窄（上下 %.2fcm / 左右 %.2fcm）'
            % ((geo['top'] or 0) / 1440 * 2.54, (geo['left'] or 0) / 1440 * 2.54),
            '过窄易判「文字贴边」；但注意窄边距常用于凑够每页 50 行，改动前需复测行数')

    if C.has_line_numbers(doc):
        add('WARN', '行号已开启（w:lnNumType）',
            '提交版通常应关闭行号：行号非源码内容，且随软件/打印机变化')
    else:
        data['line_numbers'] = False

    spacings = set(re.findall(r'<w:spacing[^>]*w:line="(\d+)"[^>]*w:lineRule="(\w+)"', doc))
    if not spacings:
        add('WARN', '未固定行距',
            '默认单倍行距下，换个办公软件/电脑排版会漂，页数随之变化')
    else:
        data['line_spacing'] = list(spacings)

    szs = sorted(set(re.findall(r'<w:sz w:val="(\d+)"', doc)))
    data['font_sizes_halfpt'] = szs
    if '21' not in szs:
        add('WARN', '正文字号不是五号（sz=21 / 10.5pt）', '当前字号集合：%s' % szs)
    if '<w:szCs' not in doc:
        add('WARN', '缺少 w:szCs 标签', '换软件打开可能变字号')
    if not re.search(r'<w:sz w:val="\d+"/>', doc):
        add('WARN', '缺少 w:sz 标签', '仅靠 w:szCs 控制字号不安全')

    # ------------------------------ 页眉页脚
    hf = C.header_footer_map(path)
    data['header_footer'] = hf
    default_hdr = C.part_text(path, hf['header'].get('default', '')) if hf['header'].get('default') else None
    data['header_text'] = default_hdr
    if not default_hdr or not default_hdr.strip():
        add('FAIL', '页眉为空或缺失',
            '页眉必须标注软件名称+版本号，且与申请表全称一字不差——这是唯一硬性驳回点')
    else:
        add('INFO', '页眉：%s' % default_hdr.strip(), '需人工核对与申请表全称是否完全一致')

    if hf.get('evenAndOddHeaders'):
        even_hdr = C.part_text(path, hf['header'].get('even', '')) if hf['header'].get('even') else None
        if not even_hdr or not even_hdr.strip():
            add('FAIL', '开启了「奇偶页不同」但偶数页页眉为空',
                '偶数页将没有软件名称，必须补齐或关闭 evenAndOddHeaders')

    ftr = hf['footer'].get('default')
    ftr_txt = C.part_text(path, ftr) if ftr else None
    has_page_field = False
    if ftr:
        with zipfile.ZipFile(path) as zz:
            has_page_field = b'PAGE' in zz.read(ftr)
    hdr_page_field = False
    if hf['header'].get('default'):
        with zipfile.ZipFile(path) as zz:
            hdr_page_field = b'PAGE' in zz.read(hf['header']['default'])
    data['page_number_where'] = ('header' if hdr_page_field else
                                 'footer' if has_page_field else None)
    if not has_page_field and not hdr_page_field:
        add('FAIL', '找不到页码域', '页眉或页脚必须有 PAGE 域')
    elif has_page_field:
        add('WARN', '页码在页脚',
            '官方措辞为「右上角应标注页码」；页脚页码通行但不完全对齐原文')

    # ------------------------------ 敏感信息
    for pat, name, level in PATTERNS:
        hits = re.findall(pat, txt)
        if hits:
            uniq = sorted(set(hits))[:5]
            add(level, '命中%s %d 处' % (name, len(hits)), '示例：%s' % uniq)

    date_hits = [l.strip() for l in lines if DATE_COMMENT.match(l)]
    if date_hits:
        add('WARN', '代码内散落日期注释 %d 处' % len(date_hits),
            '示例：%s（多为编辑器自动写入的噪音，建议删除）' % date_hits[:3])

    # ------------------------------ 源码覆盖度（可选）
    if src_dir and os.path.isdir(src_dir):
        present, absent = _coverage(src_dir, set(l.strip() for l in nonblank))
        data['src_files_present'] = present
        data['src_files_absent'] = absent
        add('INFO', '源码覆盖 %d 个文件' % len(present),
            '程序总页数超过 60 页时，规范要求交【前 30 页 + 后 30 页】')
        if absent:
            add('WARN', '源码覆盖不完整：%d 个文件未出现在文档中' % len(absent),
                '未包含：%s' % ', '.join(sorted(absent)[:12]))

    # ------------------------------ Word / PDF 实测
    if use_word:
        pdf = None
        try:
            import pymupdf          # noqa: F401
            import tempfile
            pdf = os.path.join(tempfile.gettempdir(), 'rz_audit_%d.pdf' % os.getpid())
        except ImportError:
            pass
        try:
            st = C.word_stats(path, pdf_path=pdf)
            data['word'] = st
            pages = st['pages']
            data['pages'] = pages
            if pages > 60:
                add('WARN', '共 %d 页，超过 60 页' % pages,
                    '规范：超过 60 页只提交【前 30 页 + 后 30 页】；不要靠缩页边距硬挤')
            if pdf and os.path.exists(pdf):
                per = _pdf_lines_per_page(pdf, st['header_text'])
                os.remove(pdf)
                data['lines_per_page'] = per
                low = [(i + 1, n) for i, n in enumerate(per) if n < 50]
                low_strict = [p for p in low if p[0] != len(per)]
                if low_strict:
                    add('FAIL', '有 %d 页不足 50 行' % len(low_strict),
                        '低于 50 行的页码：%s' % [p for p, _ in low_strict][:10])
                elif low:
                    add('INFO', '末页 %d 行（结束页可少于 50 行）' % per[-1], '')
                else:
                    add('INFO', '每页 %d~%d 行，全部达标' % (min(per), max(per)), '')
            elif not pdf:
                add('INFO', '未安装 pymupdf，跳过逐页行数核对',
                    'pip install pymupdf -i https://mirrors.aliyun.com/pypi/simple/')
        except ImportError:
            add('INFO', '未安装 pywin32，跳过 Word 实测页数',
                'pip install pywin32 -i https://mirrors.aliyun.com/pypi/simple/')
        except Exception as e:
            add('INFO', 'Word 实测失败：%s' % e, '可用 --no-word 跳过')
        finally:
            if pdf and os.path.exists(pdf):
                try:
                    os.remove(pdf)
                except OSError:
                    pass
            C.kill_word()

    return issues, data


def _coverage(src_dir, doc_set):
    """用每个文件最长的几条行做指纹，判断该文件是否出现在文档中。"""
    exts = ('.py', '.js', '.ts', '.vue', '.java', '.c', '.cpp', '.h', '.cs',
            '.go', '.rs', '.php', '.rb', '.kt', '.swift', '.css', '.html',
            '.json', '.scss', '.less', '.xml', '.sql', '.sh', '.bat', '.ps1')
    skip = ('node_modules', '.git', 'dist', 'build', 'unpackage', 'target',
            '__pycache__', '.next', '.nuxt', 'vendor')
    present, absent = [], []
    for root, dirs, files in os.walk(src_dir):
        dirs[:] = [d for d in dirs if d not in skip]
        for f in files:
            if not f.lower().endswith(exts):
                continue
            p = os.path.join(root, f)
            rel = os.path.relpath(p, src_dir).replace('\\', '/')
            try:
                ls = [x.strip() for x in open(p, encoding='utf-8', errors='ignore').read().split('\n')]
            except Exception:
                continue
            ls = [x for x in ls if x]
            if not ls:
                continue
            cand = sorted(set(x for x in ls if len(x) >= 30), key=len, reverse=True)[:3] or ls[:3]
            hit = sum(1 for x in cand if x in doc_set)
            (present if hit >= 2 else absent).append(rel)
    return present, absent


def _pdf_lines_per_page(pdf_path, header_text):
    import pymupdf
    doc = pymupdf.open(pdf_path)
    per = []
    for page in doc:
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
                # 左侧独立数字（行号残留 或 页码）
                if re.fullmatch(r'\d+', s) and ln['bbox'][0] < 90:
                    continue
                n += 1
        per.append(n)
    doc.close()
    return per


def render(issues, data):
    order = ['FAIL', 'WARN', 'INFO']
    label = {'FAIL': '必须处理', 'WARN': '建议处理', 'INFO': '提示'}
    out = []
    for lv in order:
        group = [i for i in issues if i[0] == lv]
        if not group:
            continue
        out.append('\n【%s】' % label[lv])
        for _, title, detail in sorted(group, key=lambda x: sev(x[0])):
            out.append('  · %s' % title)
            if detail:
                out.append('      %s' % detail)
    return '\n'.join(out)


def main():
    ap = argparse.ArgumentParser(description='软著源代码文档合规体检')
    ap.add_argument('docx', help='待检 docx 路径')
    ap.add_argument('--src', help='源码目录（用于比对内容覆盖度）')
    ap.add_argument('--no-word', action='store_true', help='跳过 Word 实测')
    ap.add_argument('--json', action='store_true', help='输出 JSON')
    args = ap.parse_args()

    if not os.path.isfile(args.docx):
        print('文件不存在：%s' % args.docx)
        return 2

    issues, data = audit(args.docx, args.src, use_word=not args.no_word)
    n_fail = sum(1 for i in issues if i[0] == 'FAIL')

    if args.json:
        print(json.dumps({'issues': issues, 'data': data}, ensure_ascii=False, indent=2))
    else:
        print('=' * 62)
        print('软著源代码文档体检：%s' % os.path.basename(args.docx))
        print('=' * 62)
        print('  逻辑行 %s ｜ 非空 %s ｜ 空行 %s'
              % (data.get('lines_total'), data.get('lines_nonblank'), data.get('lines_blank')))
        if data.get('pages'):
            print('  页数 %s' % data['pages'])
        if data.get('lines_per_page'):
            p = data['lines_per_page']
            print('  每页行数 %d~%d' % (min(p), max(p)))
        print(render(issues, data))
        print('\n结论：%s' % ('存在 %d 项必须处理，不可直接提交' % n_fail if n_fail
                            else '无必须处理项；请人工核对页眉软件名与申请表是否一字不差'))
    return 1 if n_fail else 0


if __name__ == '__main__':
    sys.exit(main())
