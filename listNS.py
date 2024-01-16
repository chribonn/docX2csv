import xml.etree.ElementTree as ET

fileName = r'.\Test\word\document.xml'

my_namespaces = dict([
    node for (_, node) in ET.iterparse(fileName, events=['start-ns'])
])

print (my_namespaces["wpc"])