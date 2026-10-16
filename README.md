# ruanzhu-code-docx

软件著作权（软著）**源代码文档**的生成、合规体检与定点修复工具。

从源码目录生成符合规范的 `.docx`，或对已有的软著代码文档做体检与清理。

## 它解决什么问题

软著提交的源代码文档有一堆不写在明面上、但会直接导致补正的坑：

- 页眉软件名与申请表全称不一致 → **硬性驳回**
- 每页不足 50 行 → 补正
- 程序超过 60 页却没按「前 30 页 + 后 30 页」提交
- 默认单倍行距，换台电脑页数就变
- 行号忘了关
- 提取时混进 HTML 脚手架
- 代码里带着真实邮箱 / 日期注释 / 版权声明
- 单个 58 万字符的压缩 GeoJSON 渲染出一百多页乱码

这套脚本把上述检查全部自动化，并且把「改行距会改页数」这类非直觉行为
交给实测循环而不是估算。

## 三个脚本

| 脚本 | 用途 |
|---|---|
| `scripts/generate.py` | 从源码目录生成源代码文档，支持全量或「前30页+后30页」 |
| `scripts/audit.py` | 对已有 docx 做合规体检，按 必须处理/建议处理/提示 三级输出 |
| `scripts/fix.py` | 定点修复（删敏感信息、关行号、页码移页眉右上角、反推行距控制页数） |

## 安装

```bash
pip install pywin32 pymupdf -i https://mirrors.aliyun.com/pypi/simple/
```

> `pywin32` 用于调用本机 Word 实测页数与导出 PDF，清华源装不上，需用阿里源。
> `pymupdf` 用于逐页统计真实行数。两者缺失时脚本会降级并说明。

需要 Windows + 本机安装 Microsoft Word。

## 用法

### 生成

```bash
python scripts/generate.py --src ./MyProject --out 软著代码.docx \
    --name "我的系统 V1.0" \
    --template 旧的软著代码.docx \
    --mode first30last30 \
    --order main.js,App.vue,pages/index.vue
```

### 体检

```bash
python scripts/audit.py 软著代码.docx --src ./MyProject
```

### 修复

```bash
python scripts/fix.py 软著代码.docx --out 软著代码_优化版.docx \
    --strip-text "someone@qq.com" \
    --strip-date-comments \
    --no-line-numbers --page-number header-right --trim-blank-lines \
    --fit-pages 60
```

## 文档

- [`SKILL.md`](SKILL.md) —— 完整工作流与规范速查
- [`references/ruanzhu-requirements.md`](references/ruanzhu-requirements.md) —— 规范逐条解读
- [`references/ooxml-pitfalls.md`](references/ooxml-pitfalls.md) —— OOXML 静默出错的坑清单

## 核心设计

- **所有修改都在 XML 层定点完成**，除指定内容外正文一字不动。
- **`--fit-pages` 从宽到紧试行距**，取第一个满足「页数 ≤ N 且每页 ≥ 50 行」的值，
  而不是靠公式估算。
- **超长行自动防护**：含单行 > 5000 字符的文件默认整份跳过。

具体实现约定见 `references/ooxml-pitfalls.md`，改动脚本前请先读。

## License

MIT
