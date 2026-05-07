import fitz
import os
import sys

os.environ['PYTHONIOENCODING'] = 'utf-8'

try:
    import fitz
except ImportError:
    print('请先安装PyMuPDF库: pip install PyMuPDF')
    sys.exit(1)

pdf_path = os.path.join(os.path.dirname(__file__), '上海交通大学MBA学位论文样本.pdf')
if not os.path.exists(pdf_path):
    print(f'PDF文件不存在: {pdf_path}')
    sys.exit(1)

doc = fitz.open(pdf_path)

output_lines = []
output_lines.append(f'总页数: {len(doc)}')
output_lines.append('')

# 读取前15页的内容和格式信息
for page_num in range(min(15, len(doc))):
    page = doc[page_num]
    output_lines.append(f'=== 第 {page_num + 1} 页 ===')
    
    # 获取文本块信息
    blocks = page.get_text('dict')['blocks']
    for block in blocks:
        if 'lines' in block:
            for line in block['lines']:
                for span in line['spans']:
                    text = span['text'].strip()
                    if text:
                        font = span['font']
                        size = span['size']
                        flags = span['flags']
                        bold = '加粗' if flags & 2 else '常规'
                        italic = '斜体' if flags & 1 else ''
                        color = span['color']
                        origin = span['origin']
                        bbox = span['bbox']
                        output_lines.append(f'  文本: "{text[:60]}"')
                        output_lines.append(f'    字体: {font}, 字号: {size:.1f}, 字重: {bold} {italic}')
                        output_lines.append(f'    位置: y={origin[1]:.1f}, x={origin[0]:.1f}')
                        output_lines.append('')
    output_lines.append('')

doc.close()

# 写入文件
with open('output_utf8.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

print('分析完成，结果已保存到 output_utf8.txt')
