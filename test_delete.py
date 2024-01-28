from docx import Document
import re


def has_only_digits_and_symbols(text):
    pattern = r'^[0-9\s!"#$%&\'()*+,\-.\/:;<=>?@\[\\\]^_`{|}…]*$'
    if re.match(pattern, text):
        #print("不是有效字段")
        return True#不包含
    else:
        #print("是有效字段")
        return False#"包含有效内容"

def delete_paragraphs_with_comma(file_path):
    # 打开 Word 文档
    doc = Document(file_path)

    # 遍历文档中的每个段落
    for paragraph in doc.paragraphs:
        #print(paragraph.text)
        str=paragraph.text
        if has_only_digits_and_symbols(paragraph.text):
            p = paragraph._element
            p.getparent().remove(p)
            paragraph._p = paragraph._element = None
        elif len(str)==0:
            p = paragraph._element
            p.getparent().remove(p)
            paragraph._p = paragraph._element = None
    # 保存修改后的文档
    doc.save(file_path)

def rubbish_delete():
    file_path = "best.docx"  # Word 文件路径
    delete_paragraphs_with_comma(file_path)

