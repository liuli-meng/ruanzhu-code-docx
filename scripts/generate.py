#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""从源码目录生成软著源代码文档（docx）。

用法示例：
    # 全量提交（程序不足 60 页时）
    python generate.py --src ./AquaSentinel-MR --out 软著代码.docx \
        --name "多区域水质在线监测智能预警系统 V1.0" --mode full

    # 超过 60 页时按规范交【前 30 页 + 后 30 页】
    python generate.py --src ./proj --out 软著代码.docx \
        --name "XX系统 V1.0" --mode first30last30

    # 指定文件顺序（推荐的顺序：入口 -> api -> utils -> 组件 -> 页面）
    python generate.py --src ./proj --order main.js,App.vue,pages/a.vue

不指定 --template 时会用本机 Word 临时造一个底稿（需要 pywin32 + Word）。
建议直接复用你已排好版的历史文档作模板：--template 旧的软著代码.docx
"""

import argparse
import os
import re
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_common as C   # noqa: E402
import hf                 # noqa: E402

DEFAULT_EXTS = ('.py', '.js', '.ts', '.vue', '.java', '.c', '.cpp', '.h', '.cs',
                '.go', '.rs', '.php', '.rb', '.kt', '.swift', '.m', '.scala',
                '.css', '.scss', '.less', '.html', '.htm', '.json', '.xml',
                '.sql', '.sh', '.bat', '.ps1', '.yml', '.yaml')

SKIP_DIRS = {'node_modules', '.git', 'dist', 'build', 'unpackage', 'target',
             '__pycache__', '.next', '.nuxt', 'vendor', '.idea', '.vscode',
             'coverage', '.venv', 'venv', 'env', 'out', 'bin', 'obj'}

# 入口类文件优先，让文档从程序起点开始
ENTRY_FIRST = ['main.js', 'main.ts', 'main.py', 'app.py', 'index.js', 'index.ts',
               'app.js', 'App.vue', 'index.html', 'manage.py', 'server.js',
               'Program.cs', 'main.c', 'main.cpp', 'Main.java', 'app.vue']


def collect(src, exts, order=None, exclude=()):
    """收集文件 -> [(rel_path, abs_path)]，按 order 或默认启发式排序。"""
    found = {}
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith('.')]
        for f in files:
            if not f.lower().endswith(tuple(exts)):
                continue
            p = os.path.join(root, f)
            rel = os.path.relpath(p, src).replace('\\', '/')
            if any(_match(rel, pat) for pat in exclude):
                continue
            found[rel] = p

    if order:
        seq = []
        for want in order:
            want = want.strip().replace('\\', '/')
            if not want:
                continue
            if want in found:
                seq.append((want, found[want]))
            else:
                cand = [k for k in found if k.endswith('/' + want) or k == want]
                if cand:
                    seq.append((cand[0], found[cand[0]]))
                else:
                    print('  ! 顺序里指定的文件未找到：%s' % want)
        used = {k for k, _ in seq}
        for k in sorted(found):
            if k not in used:
                seq.append((k, found[k]))
        return seq

    def key(rel):
        base = os.path.basename(rel)
        try:
            rank = ENTRY_FIRST.index(base)
        except ValueError:
            rank = len(ENTRY_FIRST)
        return (rank, rel)
    return [(k, found[k]) for k in sorted(found, key=key)]


def read_lines(files, keep_blank=False, skip_long=5000):
    """读入所有行。

    skip_long: 若某文件存在长度超过该值的行，视为「压缩/生成的数据文件」
    （典型：单行 58 万字符的 GeoJSON、min.js），**整份跳过**。
    这类文件在软著文档里只会渲染成上百页乱码，没有审查价值。
    传 0 表示不跳过。
    返回 (lines, report, skipped)
    """
    lines, report, skipped = [], [], []
    for rel, p in files:
        try:
            txt = open(p, encoding='utf-8').read()
        except UnicodeDecodeError:
            txt = open(p, encoding='gbk', errors='ignore').read()
        ls = txt.splitlines()
        longest = max((len(x) for x in ls), default=0)
        if skip_long and longest > skip_long:
            skipped.append((rel, len(ls), longest))
            continue
        keep = [l for l in ls if l.strip()] if not keep_blank else ls
        lines.extend(keep)
        report.append((rel, len(ls), len(keep)))
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    if skipped:
        print('  已跳过 %d 个含超长行的文件（疑似压缩数据，软著无审查价值）：' % len(skipped))
        for rel, n, mx in skipped:
            print('    - %s（%d 行，单行最长 %d 字符）' % (rel, n, mx))
    return lines, report, skipped


def make_template(path, top_cm=2.54, bottom_cm=2.54, left_cm=3.17, right_cm=3.17):
    """用 Word 造一个带页眉页脚占位的空白底稿。"""
    import win32com.client as win32
    app = win32.Dispatch('Word.Application')
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Add()
        ps = doc.Sections(1).PageSetup
        ps.PageWidth = app.CentimetersToPoints(21.0)
        ps.PageHeight = app.CentimetersToPoints(29.7)
        ps.TopMargin = app.CentimetersToPoints(top_cm)
        ps.BottomMargin = app.CentimetersToPoints(bottom_cm)
        ps.LeftMargin = app.CentimetersToPoints(left_cm)
        ps.RightMargin = app.CentimetersToPoints(right_cm)
        # 建出 header/footer 部件（内容稍后由 XML 覆盖）
        doc.Sections(1).Headers(1).Range.Text = 'HDR'
        doc.Sections(1).Footers(1).Range.Text = 'FTR'
        doc.SaveAs2(os.path.abspath(path), FileFormat=16)   # wdFormatXMLDocument
        doc.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        C.kill_word()
    return path


def _match(rel, pat):
    """支持 glob（*、?）与子串匹配。"""
    import fnmatch
    pat = pat.strip()
    if not pat:
        return False
    return fnmatch.fnmatch(rel, pat) or pat in rel


def build(src, out, name, template=None, exts=DEFAULT_EXTS, order=None,
          exclude=(), skip_long=5000,
          line_spacing_pt=13.5, mode='full', keep_blank=False,
          page_number='header-right', margins=(2.54, 2.54, 3.17, 3.17),
          verbose=True):
    tmpdir = tempfile.mkdtemp(prefix='rz_')
    try:
        # 1. 收集
        files = collect(src, exts, order, exclude)
        if not files:
            raise SystemExit('没有在 %s 里找到符合扩展名的源码文件' % src)
        lines, report, skipped = read_lines(files, keep_blank=keep_blank,
                                            skip_long=skip_long)
        if verbose:
            print('收集到 %d 个文件，共 %d 行（去空行后）' % (len(files) - len(skipped), len(lines)))

        # 2. 模板
        tpl = template
        if not tpl:
            tpl = os.path.join(tmpdir, 'tpl.docx')
            make_template(tpl, *margins)

        # 3. 构造正文
        spacing_twips = int(round(line_spacing_pt * 20)) if line_spacing_pt else None
        para = C.build_para(lines, line_spacing=spacing_twips)
        doc = C.read_doc_xml(tpl)
        doc = C.replace_body(doc, para)
        doc = doc.replace('<w:lnNumType w:countBy="1"/>', '')
        doc = re.sub(r'<w:lnNumType[^>]*/>', '', doc)

        # 4. 页眉页脚
        geo = C.page_geometry(doc)
        body_w = (geo['pw'] or 11906) - (geo['left'] or 1440) - (geo['right'] or 1440)
        hfm = C.header_footer_map(tpl)
        parts = {C.DOC_XML: doc}

        hdr_part = hfm['header'].get('default')
        if page_number == 'header-right' and name:
            if not hdr_part:
                raise SystemExit('模板缺少页眉部件，请给一个带页眉的 docx 作 --template')
            hx = _read(tpl, hdr_part)
            parts[hdr_part] = hf.set_header(hx, name, body_w, page_at_right=True)
        elif name and hdr_part:
            hx = _read(tpl, hdr_part)
            parts[hdr_part] = hf.set_header(hx, name, body_w, page_at_right=False)

        ftr_part = hfm['footer'].get('default')
        if page_number == 'footer-center':
            if not ftr_part:
                raise SystemExit('模板缺少页脚部件')
            fx = _read(tpl, ftr_part)
            parts[ftr_part] = hf.set_footer(fx, page_number=True)
        elif ftr_part:
            fx = _read(tpl, ftr_part)
            parts[ftr_part] = hf.set_footer(fx, page_number=False)

        # 5. full 模式：直接产出
        full = os.path.join(tmpdir, 'full.docx')
        C.save_with(tpl, full, parts)

        if mode == 'full':
            os.makedirs(os.path.dirname(os.path.abspath(out)) or '.', exist_ok=True)
            shutil.copy2(full, out)
            if verbose:
                _report(out, lines, tpl, hdr_part)
            return out

        # 6. first30last30：实测分页后拼接
        cut = _first_last_cut(full, lines)
        if cut is None:
            raise SystemExit('无法测量分页（需要 Word + pywin32）')
        i, j = cut
        if verbose:
            print('前 30 页 = 第 1~%d 行；后 30 页 = 第 %d~%d 行；中间省 %d 行'
                  % (i, j + 1, len(lines), j - i))
        para2 = C.build_para(lines[:i] + lines[j:], line_spacing=spacing_twips)
        doc2 = C.read_doc_xml(tpl)
        doc2 = C.replace_body(doc2, para2)
        doc2 = re.sub(r'<w:lnNumType[^>]*/>', '', doc2)
        parts[C.DOC_XML] = doc2
        os.makedirs(os.path.dirname(os.path.abspath(out)) or '.', exist_ok=True)
        C.save_with(tpl, out, parts)
        if verbose:
            _report(out, lines, tpl, hdr_part)
        return out
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def _read(docx, part):
    import zipfile
    with zipfile.ZipFile(docx) as z:
        return z.read(part).decode('utf-8')


def _offsets(lines):
    """每行的起始字符位置（Word 口径：行内容 + 1 个换行符）。"""
    pos, out = 0, []
    for l in lines:
        out.append(pos)
        pos += len(l) + 1
    return out


def _first_last_cut(docx, lines, first=30, last=30):
    """在当前排版下测出前 N 页 / 后 N 页的行边界。"""
    try:
        import win32com.client as win32
    except ImportError:
        return None
    app = None
    doc = None
    try:
        app = win32.Dispatch('Word.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(os.path.abspath(docx), ReadOnly=True)
        total = doc.ComputeStatistics(2)
        if total <= first + last:
            return None
        p_mid = doc.GoTo(1, 1, first + 1).Start          # 第 31 页起点
        p_back = doc.GoTo(1, 1, total - last + 1).Start  # 倒数第 30 页起点
        offs = _offsets(lines)

        def find(cp):
            lo, hi = 0, len(offs) - 1
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if offs[mid] <= cp:
                    lo = mid
                else:
                    hi = mid - 1
            return lo
        return find(p_mid - 1) + 1, find(p_back)
    except Exception as e:
        print('  分页测量失败：%s' % e)
        return None
    finally:
        try:
            if doc is not None:
                doc.Close(False)
            if app is not None:
                app.Quit()
        except Exception:
            pass
        C.kill_word()


def _report(out, lines, tpl, hdr_part):
    try:
        st = C.word_stats(out)
        print('  页数 %s  行距 %s  字号见下' % (st['pages'], st['line_spacing']))
    except Exception:
        pass
    if hdr_part:
        print('  页眉 %r' % (C.part_text(out, hdr_part) or ''))
    print('  已生成：%s' % out)


def main():
    ap = argparse.ArgumentParser(description='从源码目录生成软著源代码文档')
    ap.add_argument('--src', required=True, help='源码目录')
    ap.add_argument('--out', required=True, help='输出 docx 路径')
    ap.add_argument('--name', default='', help='页眉上的软件名称+版本号（须与申请表全称一致）')
    ap.add_argument('--template', help='用作底稿的 docx（复用其样式/页面设置/页眉页脚）')
    ap.add_argument('--ext', help='收集的扩展名，逗号分隔，如 .py,.js')
    ap.add_argument('--order', help='显式文件顺序，逗号分隔')
    ap.add_argument('--exclude', default='',
                    help='排除的文件，逗号分隔，支持 * 通配或子串，如 "*.min.js,china.json"')
    ap.add_argument('--skip-long-lines', type=int, default=5000,
                    help='文件含超过该长度的行则整份跳过（默认 5000；传 0 关闭）')
    ap.add_argument('--line-spacing', type=float, default=13.5,
                    help='固定行距(pt)，0 表示不固定。默认 13.5')
    ap.add_argument('--mode', choices=['full', 'first30last30'], default='full',
                    help='full=全部提交；first30last30=前30页+后30页')
    ap.add_argument('--keep-blank', action='store_true', help='保留空行（默认去掉）')
    ap.add_argument('--page-number', choices=['header-right', 'footer-center', 'none'],
                    default='header-right')
    ap.add_argument('--margins', default='2.54,2.54,3.17,3.17',
                    help='上下左右边距(cm)，仅在自动造模板时生效')
    args = ap.parse_args()

    exts = DEFAULT_EXTS
    if args.ext:
        exts = tuple(x.strip() if x.strip().startswith('.') else '.' + x.strip()
                     for x in args.ext.split(',') if x.strip())
    order = [x for x in (args.order or '').split(',') if x.strip()] or None
    exclude = [x for x in (args.exclude or '').split(',') if x.strip()]
    margins = tuple(float(x) for x in args.margins.split(','))

    build(args.src, args.out, args.name, template=args.template, exts=exts,
          order=order, exclude=exclude, skip_long=args.skip_long_lines,
          line_spacing_pt=args.line_spacing, mode=args.mode,
          keep_blank=args.keep_blank, page_number=args.page_number, margins=margins)
    print('\n下一步：python audit.py "%s" --src "%s"' % (args.out, args.src))


if __name__ == '__main__':
    main()
