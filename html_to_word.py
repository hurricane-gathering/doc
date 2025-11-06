from pathlib import Path
from typing import Optional, Union
import re
import copy
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def html_to_word(
    html_content: str,
    output_path: Optional[Union[str, Path]] = None
) -> Document:
    """将 HTML 转换为 Word 文档

    Args:
        html_content: HTML 字符串
        output_path: 输出路径，为 None 则不保存

    Returns:
        Document 对象

    Raises:
        ValueError: HTML 内容为空
        IOError: 文件保存失败
    """

    doc = Document()
    _setup_styles(doc)

    soup = BeautifulSoup(html_content, 'lxml')

    # 添加姓名和联系信息
    _process_header(doc, soup)

    # 添加间隔
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # 处理所有章节
    _process_sections(doc, soup)

    # 保存文档
    if output_path:
        _save_document(doc, output_path)

    return doc


def _setup_styles(doc: Document):
    """配置文档默认样式"""
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')


def _process_header(doc: Document, soup: BeautifulSoup):
    """处理姓名和联系信息"""
    # 添加姓名
    name_header = soup.find('div', class_='name-header')
    if name_header:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(name_header.get_text(strip=True))
        run.font.size = Pt(25)
        run.bold = True
        para.paragraph_format.space_after = Pt(6)

    # 添加联系信息
    contact_div = soup.find('div', class_='contact-info')
    if contact_div:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(contact_div.get_text(strip=True))
        run.font.size = Pt(12)
        para.paragraph_format.space_after = Pt(12)


def _process_sections(doc: Document, soup: BeautifulSoup):
    """处理所有章节"""
    sections = soup.find_all('div', class_='section-title')

    for idx, section in enumerate(sections):
        # 添加章节标题
        _add_section_title(doc, section.get_text(strip=True))

        # 处理该章节的列表
        next_ul = section.find_next_sibling('ul', class_='ul-section')
        if next_ul:
            _process_list(doc, next_ul)

        # 章节间距（最后一个除外）
        if idx < len(sections) - 1:
            doc.add_paragraph().paragraph_format.space_after = Pt(8)


def _add_section_title(doc: Document, title: str):
    """添加章节标题（加粗 + 底部边框）"""
    para = doc.add_paragraph()
    run = para.add_run(title)
    run.font.size = Pt(13)
    run.bold = True

    # 添加底部边框
    p = para._element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)

    # 设置间距
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.space_before = Pt(6)


def _process_list(doc: Document, ul_element):
    """处理列表元素"""
    for li in ul_element.find_all('li', recursive=False):
        has_dot = li.find('span', class_='dot') is not None

        # 提取右对齐文本
        right_span = li.find('span', class_='right-span')
        right_text = right_span.get_text(strip=True) if right_span else ""
        if right_span:
            right_span.extract()

        # 移除圆点元素（但保留其位置信息）
        dot_span = li.find('span', class_='dot')
        if dot_span:
            dot_span.extract()

        # 添加列表项，支持部分加粗
        _add_list_item_with_formatting(doc, li, has_dot, right_text)


def _add_list_item(doc: Document, text: str, is_bold: bool, is_italic: bool,
                   has_dot: bool, right_text: str):
    """添加单个列表项（保留用于兼容性）"""
    para = doc.add_paragraph()

    # 添加圆点（如果需要）
    if has_dot:
        para.paragraph_format.left_indent = Inches(0.3)
        para.add_run('• ').font.size = Pt(11)

    # 添加主文本
    run = para.add_run(text)
    run.font.size = Pt(11)
    run.bold = is_bold
    run.italic = is_italic

    # 添加右对齐文本（如果有）
    if right_text:
        right_run = para.add_run('\t' + right_text)
        right_run.font.size = Pt(11)
        right_run.bold = is_bold
        right_run.italic = is_italic
        para.paragraph_format.tab_stops.add_tab_stop(
            Inches(6.0), WD_ALIGN_PARAGRAPH.RIGHT)

    # 设置紧凑的行距
    para.paragraph_format.space_after = Pt(1)
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.line_spacing = 1.15


def _add_list_item_with_formatting(doc: Document, li_element, has_dot: bool, right_text: str):
    """添加单个列表项，支持部分加粗和斜体"""
    para = doc.add_paragraph()

    # 添加圆点（如果需要）
    if has_dot:
        para.paragraph_format.left_indent = Inches(0.3)
        para.add_run('• ').font.size = Pt(11)

    # 遍历 li 元素的所有子节点，分别处理加粗和非加粗部分
    for element in li_element.children:
        if isinstance(element, Tag):
            # 处理标签元素
            if element.name == 'b' or element.name == 'strong':
                # 加粗文本，移除换行符但保留空格
                text = element.get_text(separator=' ', strip=False)
                text = re.sub(r'[\n\r]+', ' ', text)  # 将换行符替换为空格
                text = text.strip()
                if text:
                    run = para.add_run(text)
                    run.font.size = Pt(11)
                    run.bold = True
            elif element.name == 'i' or element.name == 'em':
                # 斜体文本，移除换行符但保留空格
                text = element.get_text(separator=' ', strip=False)
                text = re.sub(r'[\n\r]+', ' ', text)
                text = text.strip()
                if text:
                    run = para.add_run(text)
                    run.font.size = Pt(11)
                    run.italic = True
            elif element.name in ['span', 'div', 'p']:
                # 递归处理嵌套元素
                _process_nested_elements(para, element)
            else:
                # 其他标签，提取文本，移除换行符但保留空格
                text = element.get_text(separator=' ', strip=False)
                text = re.sub(r'[\n\r]+', ' ', text)
                text = text.strip()
                if text:
                    run = para.add_run(text)
                    run.font.size = Pt(11)
        elif isinstance(element, NavigableString):
            # 处理纯文本节点，移除换行符但保留空格
            text = str(element)
            # 将换行符替换为空格，然后移除首尾空白
            text = re.sub(r'[\n\r]+', ' ', text)
            text = text.strip()
            if text:
                run = para.add_run(text)
                run.font.size = Pt(11)

    # 添加右对齐文本（如果有）
    if right_text:
        right_run = para.add_run('\t' + right_text)
        right_run.font.size = Pt(11)
        para.paragraph_format.tab_stops.add_tab_stop(
            Inches(6.0), WD_ALIGN_PARAGRAPH.RIGHT)

    # 设置紧凑的行距
    para.paragraph_format.space_after = Pt(1)
    para.paragraph_format.space_before = Pt(1)
    para.paragraph_format.line_spacing = 1.15


def _process_nested_elements(para, element):
    """递归处理嵌套元素，保持格式"""
    for child in element.children:
        if isinstance(child, Tag):
            if child.name == 'b' or child.name == 'strong':
                # 加粗文本，移除换行符但保留空格
                text = child.get_text(separator=' ', strip=False)
                text = re.sub(r'[\n\r]+', ' ', text)
                text = text.strip()
                if text:
                    run = para.add_run(text)
                    run.font.size = Pt(11)
                    run.bold = True
            elif child.name == 'i' or child.name == 'em':
                # 斜体文本，移除换行符但保留空格
                text = child.get_text(separator=' ', strip=False)
                text = re.sub(r'[\n\r]+', ' ', text)
                text = text.strip()
                if text:
                    run = para.add_run(text)
                    run.font.size = Pt(11)
                    run.italic = True
            elif child.name in ['span', 'div', 'p']:
                _process_nested_elements(para, child)
            else:
                # 其他标签，移除换行符但保留空格
                text = child.get_text(separator=' ', strip=False)
                text = re.sub(r'[\n\r]+', ' ', text)
                text = text.strip()
                if text:
                    run = para.add_run(text)
                    run.font.size = Pt(11)
        elif isinstance(child, NavigableString):
            # 处理纯文本节点，移除换行符但保留空格
            text = str(child)
            # 将换行符替换为空格，然后移除首尾空白
            text = re.sub(r'[\n\r]+', ' ', text)
            text = text.strip()
            if text:
                run = para.add_run(text)
                run.font.size = Pt(11)


def _save_document(doc: Document, output_path: Union[str, Path]):
    """保存文档"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        doc.save(str(output_path))
        print(f"✓ Word 文档已保存至: {output_path}")
    except Exception as e:
        raise IOError(f"保存文档失败: {e}")

