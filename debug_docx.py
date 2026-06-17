#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn

docx_path = r'C:\Users\lenovo\Desktop\常用运算符和数学函数_含示例解释.docx'
doc = Document(docx_path)

body = doc.element.body

print("=== Searching for Section markers ===")
for i, p in enumerate(body.findall(qn('w:p'))):
    texts = p.findall(qn('w:t'))
    combined = ''.join(t.text or '' for t in texts)
    if '五' in combined or '���' in combined:
        print(f"P{i}: contains '五' or garbled: '{combined[:120]}'")
    if '4.2' in combined:
        print(f"P{i}: contains '4.2': '{combined[:120]}'")

print("\n=== All paragraphs around section 4/5 ===")
for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if text and ('四' in text or '五' in text or '六' in text or '4.' in text):
        print(f"P{i}: '{text[:100]}'")
