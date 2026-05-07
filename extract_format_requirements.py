import fitz
import os
import sys

os.environ['PYTHONIOENCODING'] = 'utf-8'

try:
    import fitz
except ImportError:
    print('请先安装PyMuPDF库: pip install PyMuPDF')
    sys.exit(1)

pdf_path = os.path.join(os.path.dirname(__file__), '上海交通大学MBA学位论文格式的统一要求.pdf')
if not os.path.exists(pdf_path):
    print(f'PDF文件不存在: {pdf_path}')
    sys.exit(1)

doc = fitz.open(pdf_path)

output_lines = []
output_lines.append(f'总页数: {len(doc)}')
output_lines.append('')
output_lines.append('='*80)
output_lines.append('上海交通大学MBA学位论文格式的统一要求 - 详细内容提取')
output_lines.append('='*80)

# 提取所有页面的文本内容
for page_num in range(len(doc)):
    page = doc[page_num]
    output_lines.append('')
    output_lines.append(f'=== 第 {page_num + 1} 页 ===')
    output_lines.append('')
    
    # 获取页面文本
    text = page.get_text()
    if text.strip():
        output_lines.append(text)
    else:
        output_lines.append('[此页可能为图片或扫描内容，无法直接提取文本]')
    
    # 尝试获取详细的文本块信息
    blocks = page.get_text('dict')['blocks']
    for block in blocks:
        if 'lines' in block:
            for line in block['lines']:
                for span in line['spans']:
                    text = span['text'].strip()
                    if text and len(text) > 5:  # 只显示较长的文本段
                        font = span['font']
                        size = span['size']
                        flags = span['flags']
                        bold = '加粗' if flags & 2 else '常规'
                        italic = '斜体' if flags & 1 else ''
                        origin = span['origin']
                        output_lines.append(f'  文本: "{text}"')
                        output_lines.append(f'    字体: {font}, 字号: {size:.1f}, 字重: {bold} {italic}')
                        output_lines.append(f'    位置: y={origin[1]:.1f}, x={origin[0]:.1f}')
                        output_lines.append('')

doc.close()

# 写入文件
with open('format_requirements_detail.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

print('详细格式要求提取完成，结果已保存到 format_requirements_detail.txt')