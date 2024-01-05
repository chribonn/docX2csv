# https://docs.python.org/3/library/xml.etree.elementtree.html
import xml.etree.ElementTree as ET
# tree = ET.parse(r'D:\chrib\PythonProjects\summariseDocx\data.xml')
tree = ET.parse(r'D:\chrib\PythonProjects\summariseDocx\Test\word\document.xml')

root = tree.getroot()

ET.register_namespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
# ET.dump(tree)
for x in root.findall('.//w:p', ns):
    print (x.text, " ", x.tag)
