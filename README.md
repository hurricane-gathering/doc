# HTML 简历转 Word 工具

一个简洁高效的 Python 工具，用于将 HTML 格式的简历转换为 Word 文档（.docx）。

## ✨ 特性

- 🎯 **精准转换**: 完美还原 HTML 简历的格式和样式，包括边框、行距等细节
- 📝 **完整支持**: 支持标题、联系信息、教育背景、工作经验等常见简历元素
- ✨ **章节边框**: 自动为章节标题添加底部边框线，与 HTML 完全一致
- 🎨 **样式还原**: 精确还原加粗、斜体、居中、右对齐等所有样式
- 🚀 **高性能**: 快速转换，适用于批量处理
- 💡 **易于使用**: 简洁的 API，支持文件和字符串两种输入方式
- 🔧 **易维护**: 模块化设计，代码清晰易读

## 📦 安装依赖

```bash
pip install python-docx beautifulsoup4 lxml
```

或者使用 uv (推荐):

```bash
uv pip install python-docx beautifulsoup4 lxml
```

## 🚀 快速开始

### 方式 1: 从 HTML 文件转换

```python
from pathlib import Path
from main import convert_html_to_word

# 转换 HTML 文件为 Word
output_file = convert_html_to_word(
    html_path=Path('resume.html'),
    output_path=Path('resume.docx')
)
print(f"转换完成！文件保存在: {output_file}")
```

### 方式 2: 从 HTML 字符串转换

```python
from main import convert_html_to_word

html_content = '''
<!DOCTYPE html>
<html>
<body>
    <div class="name-header">张三</div>
    <div class="contact-info">+86 138 0000 0000 | email@example.com</div>
    ...
</body>
</html>
'''

output_file = convert_html_to_word(
    html_content=html_content,
    output_path=Path('resume.docx')
)
```

### 方式 3: 使用 HTMLToWordConverter 类

```python
from main import HTMLToWordConverter

converter = HTMLToWordConverter()
doc = converter.convert(html_content)
converter.save(Path('resume.docx'))
```

## 📖 运行示例

项目包含一个完整的示例，直接运行即可：

```bash
python main.py
```

这将生成：

- `sample_resume.html` - 示例 HTML 简历
- `resume_output.docx` - 转换后的 Word 文档

## 🎨 支持的 HTML 元素

| HTML 元素  | CSS 类名         | 说明                                  |
| ---------- | ---------------- | ------------------------------------- |
| 姓名标题   | `.name-header`   | 居中显示，25pt 加粗                   |
| 联系信息   | `.contact-info`  | 居中显示，12pt                        |
| 章节标题   | `.section-title` | 13pt 加粗，**带底部黑色实线边框** ✨  |
| 列表内容   | `.ul-section`    | 支持加粗、斜体、圆点，紧凑行距 (1.15) |
| 右对齐文本 | `.right-span`    | 日期等信息右对齐，**继承样式** ✨     |
| 圆点标记   | `.dot`           | 列表项前的圆点符号，自动缩进 0.3 英寸 |

### 🎯 样式还原细节

- ✅ **章节标题边框**: 完美还原 HTML 的 `border-bottom: 1px solid`，黑色实线
- ✅ **紧凑行距**: 1.15 倍行距，接近 HTML 的 `line-height: 1.2`
- ✅ **精确间距**: 段前/段后间距精细控制，视觉效果与 HTML 一致
- ✅ **样式继承**: 右对齐文本自动继承主文本的加粗/斜体样式

## 📋 HTML 模板示例

```html
<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
  </head>
  <body>
    <div class="a4-page">
      <div class="name-header">姓名</div>
      <div class="contact-info">联系方式</div>

      <div class="section-title">教育背景</div>
      <ul class="ul-section">
        <li>
          <b>大学名称<span class="right-span">2020-2024</span></b>
        </li>
        <li><i>专业名称</i></li>
        <li><span class="dot"></span><b>GPA:</b> 3.8/4.0</li>
      </ul>

      <div class="section-title">工作经验</div>
      <ul class="ul-section">
        <li>
          <b>公司名称<span class="right-span">2024-至今</span></b>
        </li>
        <li>
          <i>职位名称<span class="right-span">城市</span></i>
        </li>
        <li><span class="dot"></span>工作描述1</li>
        <li><span class="dot"></span>工作描述2</li>
      </ul>
    </div>
  </body>
</html>
```

## ⚙️ API 参考

### `convert_html_to_word()`

主要转换函数，支持灵活的输入方式。

**参数:**

- `html_path` (Optional[Path]): HTML 文件路径
- `html_content` (Optional[str]): HTML 字符串内容
- `output_path` (Optional[Path]): 输出 Word 文件路径，默认为 `resume.docx`

**返回:**

- `Path`: 生成的 Word 文档路径

**异常:**

- `ValueError`: 参数无效时
- `FileNotFoundError`: HTML 文件不存在时
- `IOError`: 文件读写失败时

### `HTMLToWordConverter` 类

用于更精细的控制和自定义。

**主要方法:**

- `__init__()`: 初始化转换器
- `convert(html_content: str) -> Document`: 转换 HTML 为 Word 文档
- `save(output_path: Path)`: 保存文档到文件

## 🔧 技术栈

- **python-docx**: Word 文档生成
- **BeautifulSoup4**: HTML 解析
- **lxml**: 高性能 XML/HTML 解析器

## 📝 代码规范

本项目遵循以下编码规范：

- ✅ PEP 8 代码风格
- ✅ 完整的类型注解
- ✅ 详细的文档字符串
- ✅ 模块化设计
- ✅ 完善的异常处理

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT License

## 🔗 相关资源

- [python-docx 文档](https://python-docx.readthedocs.io/)
- [BeautifulSoup 文档](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
