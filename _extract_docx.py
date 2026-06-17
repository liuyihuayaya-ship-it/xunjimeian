import zipfile
import re
import sys

docx_path = r'C:\Users\lenovo\Desktop\常用运算符和数学函数_含示例解释.docx'
try:
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            data = f.read()
            # Find all paragraph nodes
            paragraphs = re.findall(b'<w:p[ >].*?</w:p>', data, re.DOTALL)
            for p in paragraphs:
                texts = re.findall(b'<w:t[^>]*>([^<]*)</w:t>', p)
                line = ''.join(t.decode('utf-8', errors='ignore') for t in texts)
                if line.strip():
                    print(line)
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    sys.exit(1)
