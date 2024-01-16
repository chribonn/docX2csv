import xml.etree.ElementTree as ET

# needs to be updated with document name docx
fileName = r'.\Test\word\document.xml'

styles = []

#constants
NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NW_URI_TAG = '{' + NS_URI + '}'


tree = ET.parse(fileName)

root = tree.getroot()

ET.register_namespace("w", NS_URI)
ns = {"w": NS_URI}

for x in root.findall('.//w:pStyle', ns):
    style = x.get(NW_URI_TAG + 'val')
    if style not in styles:
        styles.append(style)
        
print (sorted(styles))
