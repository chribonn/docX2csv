from docx import Document
from docx.enum.style import WD_STYLE_TYPE

docName = "Test.docx"

styles = Document(docName).styles
assert styles[0].type == WD_STYLE_TYPE.PARAGRAPH

print (styles)