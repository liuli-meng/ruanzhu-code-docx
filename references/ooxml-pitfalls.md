# OOXML 操作踩坑清单

这份文档记录的都是**会让操作静默出错**的行为——不报错、但结果不对，
只能靠导出 PDF 看坐标或逐一比对内容才能发现。动脚本前先读一遍。

---

## 1. 缩进是 `<w:tab/>` 元素，不是 TAB 字符

这类源代码文档通常**只有 1 个 `<w:p>` 段落**，靠成千上万个 `<w:br/>` 换行，
缩进则用 `<w:tab/>` 元素表示（实测某文档有 5377 个 `<w:tab/>`，
而 `<w:t>` 内字面 TAB 字符为 0）。

**提取正文时必须同时处理三种子元素**：

```python
for r in body.iter(W + 'r'):
    for ch in r:
        tag = ch.tag.replace(W, '')
        if tag == 'br':          # 换行：结算当前行
            lines.append(''.join(cur)); cur = []
        elif tag == 't':
            cur.append(ch.text or '')
        elif tag == 'tab':       # ← 漏了这个就会误判「缩进全部丢失」
            cur.append('\t')
```

漏掉 `tab` 分支会得出完全相反的结论（「源码缩进被剥离，格式不规范」），
而文件其实完好。

---

## 2. `<w:br/>` 的位置：后置，不是前置

构造段落时把 `<w:br/>` 放在**文本之后**：

```xml
<!-- 正确：写完这行再换行 -->
<w:r><w:t>line content</w:t><w:br/></w:r>
```

若放在文本之前（`<w:br/><w:t>...</w:t>`），整篇会在**开头凭空多出一个空行**，
且后续每行的视觉位置看起来都正常，极难察觉。最后一行的 `<w:br/>` 要去掉，
否则末尾多一个空行。

---

## 3. 删 XML 片段必须补回 `</w:p>`

从 `<w:br/>` 一路砍到 `<w:sectPr` 时，会连带删掉段落的闭合标签
`</w:p>`，导致 `mismatched tag` 解析失败。删除后要显式补回。

---

## 4. `<w:tabs>` 的子元素是 `<w:tab>`，不是 `<w:center>` / `<w:right>`

把页眉做成「软件名居中 + 页码右上」需要 tab stop：

```xml
<!-- 正确 -->
<w:tabs>
  <w:tab w:val="center" w:pos="5233"/>
  <w:tab w:val="right"  w:pos="10466"/>
</w:tabs>

<!-- 错误：Word 静默忽略，tab 退化为默认 tab stop（每 720 twips） -->
<w:center w:val="center" w:pos="5233"/><w:right .../>
```

写错时**不报错**，页眉内容只是紧贴左边距排列。
诊断方法：导出 PDF，看页眉 span 的 x 坐标。

- 正确：名称居中在页面中心附近；页码贴在右边距。
- 错误：名称从左边距 +36pt 处开始（默认 tab stop 的位置）。

`w:pos` 的基准是**正文宽度**（页宽 − 左右边距）：

```
正文宽 = pgSz.w - pgMar.left - pgMar.right
居中位 = 正文宽 / 2        右对齐位 = 正文宽
```

例：`11906 - 720 - 720 = 10466` → 居中 5233，右对齐 10466。

---

## 5. 正则 `<w:t[^>]*>` 会误匹配 `<w:tabs>`

因为 `abs>` 能被 `[^>]*>` 吃掉。结果是「提取页眉文字」会吐回一大段 XML。

必须写成：

```python
r'<w:t(?:\s[^>]*)?>'      # 要求 <w:t 后面是空白或直接 >
```

同理，匹配段落要覆盖自闭合写法，否则会把自闭合段落一路吞到下一个 `</w:p>`，
**误删中间的段落**（例如页眉里的空占位段落，删掉后页眉变矮、正文行数变多）：

```python
r'<w:p\b[^>]*/>|<w:p\b[^>]*>.*?</w:p>'   # 正确
r'<w:p[ >].*?</w:p>|<w:p/>'              # 错误（漏掉带属性的自闭合）
```

---

## 6. 提取页眉文字要剔除 PAGE 域

`<w:t>` 匹配会把页码域的缓存结果一起抓进来，页眉看起来像
`XX系统 V1.01`（末尾多一个页码数字），容易误判软件名写错。

做法：先删掉 `fldChar begin` 到 `fldChar end` 的整段，再提取文本。

---

## 7. 奇偶页 / 首页页眉陷阱

文档可能有 header1/2/3（even/default/first）三套页眉。
**必须先查 `settings.xml` 是否含 `<w:evenAndOddHeaders/>`**：

- 开了 → 偶数页用 even 页眉。若 even 页眉为空，**偶数页全都没有软件名称**，是硬伤。
- 没开 → even/first 定义都不生效，只需管 default（`<w:titlePg/>` 控制首页）。

用 `header_footer_map()` 解析 `document.xml.rels` 里 rId → 部件的映射，
再看 `sectPr` 里的 `headerReference`。

---

## 8. 行号的隐藏副作用

`<w:lnNumType w:countBy="1"/>` 打开行号后，Word 若发现页边距不足以容纳行号，
会**压缩正文宽度**，导致更多行折行、**页数明显增加**（实测同一文档：带行号
64 页，去掉行号后 60 页）。所以「关行号」不只是美观问题，直接影响页数。

---

## 9. 固定行距会改页数——必须实测，不能算

`<w:spacing w:line="270" w:lineRule="exact"/>`（270 twips = 13.5pt）。

每页可用高度 ≠ 简单理论值：

```
理论值 = 页高 − 上下边距
实际值 = 理论值 − 页眉页脚占位挤压（通常是十几到几十 pt）
```

实测某文档：理论 770pt，实际约 725pt，差 45pt 来自页眉 2 段 + 页脚 1 段。

**流程**：改行距 → Word 实测页数 → 不对就调，别靠算。
`fix.py --fit-pages` 就是把这个循环自动化（**从宽到紧**试，
取第一个满足页数上限且每页 ≥ 50 行的值；从小到大试会挑到最挤的行距）。

---

## 10. Word COM 会话复用会导致方法「消失」

```python
app = win32.Dispatch('Word.Application')
... doc.Close(); app.Quit()
app2 = win32.Dispatch('Word.Application')   # ← 上一次 Quit 尚未结束
doc2.ExportAsFixedFormat(...)               # AttributeError: Open.ExportAsFixedFormat
```

表现为 `Open.ComputeStatistics` / `Open.ExportAsFixedFormat` 属性不存在。
**解决办法：统计与导出 PDF 放在同一个 Word 会话里完成。**

此外 COM 偶发退不干净，会残留 `~$xxx.docx` 锁文件，
需要 `taskkill /F /IM WINWORD.EXE` 兜底。

---

## 11. 超长行会炸掉页数

单行 582,263 字符的压缩 GeoJSON（`tools/china.json`）实测渲染出**一百多页**。

生成前应检测并跳过含超长行的文件（默认阈值 5000 字符），
或让用户显式 `--exclude`。体检时也要报出来。

---

## 12. 其他

- **改文件前先看占用**。文件被 Word/预览打开时写入报 `PermissionError`；
  可以先生成到临时目录再复制过去。
- **`/tmp` 在 Git Bash 下映射到 `%LOCALAPPDATA%\Temp`**，
  Python 里用相对路径比绝对 `/tmp/...` 更稳。
- **中文路径**：`doc_to_image` 之类的接口在中文路径下会报 "file not found"，
  必要时复制到纯英文路径再处理。
