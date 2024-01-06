# https://docs.python.org/3/library/xml.etree.elementtree.html
import xml.etree.ElementTree as ET

# params
fileName = r'.\Test\word\document.xml'

#constants
NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NW_URI_TAG = '{' + NS_URI + '}'


def proc_pPr(branch):
    # print (branch.tag, branch.attrib)
    name = branch.find(NW_URI_TAG + 'pStyle')
    if name is not None:
        style = name.get(NW_URI_TAG + 'val')
    else:
        style = None
        
    return style


tree = ET.parse(fileName)

root = tree.getroot()


ET.register_namespace("w", NS_URI)
ns = {"w": NS_URI}
# ET.dump(tree)

for x in root.findall('.//w:p', ns):
    # print (x)
    for y in x:
        if y.tag == NW_URI_TAG + 'pPr':
            result = proc_pPr(y) or (None)
            if result is not None:
                print (result)
            

