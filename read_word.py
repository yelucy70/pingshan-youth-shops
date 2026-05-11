from docx import Document

doc = Document('店铺介绍.docx')

for i, para in enumerate(doc.paragraphs):
    if para.text.strip():
        print(f"段落{i+1}: {para.text}")
        print('---')
