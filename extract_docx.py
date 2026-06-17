#!/usr/bin/env python
import zipfile
import re
import os

docx_path = r'D:\cc\target.docx'
if not os.path.exists(docx_path):
    print(f"File not found: {docx_path}")
    exit(1)

with zipfile.ZipFile(docx_path, 'r') as z:
    with z.open('word/document.xml') as f:
        data = f.read().decode('utf-8', errors='ignore')

        # Find all paragraphs
        paragraphs = re.findall(r'<w:p[ >].*?</w:p>', data, re.DOTALL)
        for p in paragraphs:
            texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)
            line = ''.join(texts)
            if line.strip():
                print(line)
