import zipfile
import tempfile
import xml.etree.ElementTree as ET


#constants
NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NW_URI_TAG = '{' + NS_URI + '}'
XML_DOC_PATH = 'word/document.xml'

def extract_document_xml(docx_file, tmp_dir):
    with zipfile.ZipFile(docx_file, 'r') as zip_ref:
        zip_ref.extract(XML_DOC_PATH, tmp_dir)
            
    return tmp_dir


def get_styles(docx_file):
    styles = []

    tree = ET.parse(docx_file)
    root = tree.getroot()

    ET.register_namespace("w", NS_URI)
    ns = {"w": NS_URI}

    for x in root.findall('.//w:pStyle', ns):
        style = x.get(NW_URI_TAG + 'val')
        if style not in styles:
            styles.append(style)
            
    return sorted(styles)



            
# if __name__ == "__main__":
#     tmp_dir = tempfile.TemporaryDirectory()
#     docx_file = "Test.docx"
#     print (extract_document_xml(docx_file, tmp_dir.name))
    