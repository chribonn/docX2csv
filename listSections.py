from docx import Document

docName = "Test.docx"

document = Document(docName)
pg = 1
print ('{0:<15} | {1:<30} | {2:>6} | {3:>6}'.format('Style Name', 'Text', 'Section', 'Page'))

for para in document.paragraphs:
    if para.contains_page_break:
        pg += 1
    else:
        for run in para.runs:
            if para.contains_page_break:
                pg += 1


    styleName = (para.style.name)
    if styleName.startswith('Column'):
        print ('{0:<15} | {1:<30} | {2:>6} | {3:>6}'.format(styleName, para.text[0:30], ' ', pg))
        print ('-' * 80)
        

