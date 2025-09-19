# Markdown to Word Converter

一个支持数学公式渲染的 Markdown 转 Word 文档工具，部署在 GitHub Pages 上。

## 功能特性

- ✅ 支持标准 Markdown 语法
- ✅ 支持数学公式渲染（LaTeX 格式）
- ✅ 支持多种数学公式格式：
  - 行内公式：`$x^2 + y^2 = z^2$`
  - 块级公式：`$$\frac{n!}{k!(n-k)!} = \binom{n}{k}$$`
  - LaTeX 格式：`\[E = mc^2\]`
- ✅ 实时预览
- ✅ 文件上传支持（拖拽或选择文件）
- ✅ 生成并下载 Word 文档

## 在线使用

访问 GitHub Pages 部署的版本：[https://tongjiliuchongwen.github.io/md-word/](https://tongjiliuchongwen.github.io/md-word/)

## 本地运行

1. 克隆仓库：
```bash
git clone https://github.com/tongjiliuchongwen/md-word.git
cd md-word
```

2. 启动本地服务器（可以使用 Python 或 Node.js）：
```bash
# 使用 Python 3
python -m http.server 8000

# 或使用 Node.js
npx serve .
```

3. 在浏览器中访问 `http://localhost:8000`

## 支持的 Markdown 语法

- 标题（# ## ### 等）
- 段落和换行
- **粗体** 和 *斜体*
- 代码块和 `行内代码`
- 列表（有序和无序）
- 引用
- 表格
- 数学公式

## 数学公式示例

```markdown
# 数学公式示例

行内公式：质能方程 $E = mc^2$ 是爱因斯坦的著名公式。

块级公式：
$$
\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}
$$

LaTeX 格式：
\[
\sum_{n=1}^{\infty} \frac{1}{n^2} = \frac{\pi^2}{6}
\]
```

## 技术栈

- **前端**: HTML5, CSS3, JavaScript (ES6+)
- **Markdown 解析**: Marked.js
- **数学公式渲染**: KaTeX
- **Word 文档生成**: docx.js
- **文件下载**: FileSaver.js
- **部署**: GitHub Pages

## 贡献

欢迎提交 Issues 和 Pull Requests！

## 许可证

MIT License
