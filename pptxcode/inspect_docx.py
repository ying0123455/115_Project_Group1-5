import sys
from docx import Document

def main():
    doc = Document('TWSE113年簡明財務報告.docx')
    content = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if text:
            content.append(text)
    
    # Write to a UTF-8 file
    with open('doc_content.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(content))
    print(f"Successfully wrote {len(content)} paragraphs to doc_content.txt")

if __name__ == '__main__':
    main()
