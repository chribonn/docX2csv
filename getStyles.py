import xml.etree.ElementTree as ET
import csv

# params
# needs to be updated with document name docx
fileName = r'.\Test\word\document.xml'

csvFileName = r'.\output.csv'
styles = ('ColumnAStyle', 'ColumnBStyle', 'ColumnCStyle')

#constants
NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NW_URI_TAG = '{' + NS_URI + '}'


def proc_pPr_pStyle(branch, srchStyles):
    # print (branch.tag, branch.attrib)
    name = branch.find(NW_URI_TAG + 'pStyle')
    if name is None:
        style = None
    else:
        style = name.get(NW_URI_TAG + 'val')
        # Process only if this is a style we are interested in
        if style not in srchStyles:
            style = None
        
    return style


def proc_pPr_sectPr(branch):
    return branch.find(NW_URI_TAG + 'sectPr') is not None



def proc_r_t(branch):
    name = branch.find(NW_URI_TAG + 't')
    text = '' if name is None else name.text
    
    # if text is null check for the xml:space="preserve" in which case add a space
    if text == '' or text is None:
        if name.get('{http://www.w3.org/XML/1998/namespace}space') == "preserve":
            text = ' '

    return '' if text is None else text



def proc_r_lastRenderedPageBreak(branch):
    return branch.find(NW_URI_TAG + 'lastRenderedPageBreak') is not None
    


def savecsv(csvFile, csvList):
    csvColumns = ['Style','Text','Section','Page','Line']
    try:
        with open(csvFile, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=csvColumns)
            writer.writeheader()
            for data in csvList:
                writer.writerow(data)
    except IOError:
        print("I/O error")


tree = ET.parse(fileName)

root = tree.getroot()
page = section = line = 1
csvList = []


ET.register_namespace("w", NS_URI)
ns = {"w": NS_URI}
# ET.dump(tree)

for x in root.findall('.//w:p', ns):
    # print (x)
    text = ''
    style = None
    for y in x:
        if y.tag == NW_URI_TAG + 'pPr':
            # The section check need to be on top because the node may not have a style
            if proc_pPr_sectPr(y):
                section += 1
            style = proc_pPr_pStyle(y, styles) or (None)
            if style is None:
                break
        elif y.tag == NW_URI_TAG + 'r':
            # Search for a rendered page breaks
            if proc_r_lastRenderedPageBreak(y):
                page += 1
                line = 1

            text += proc_r_t(y)
        
    line += 1

# https://www.tutorialspoint.com/How-to-save-a-Python-Dictionary-to-CSV-file
print (csvList)
savecsv (csvFileName, csvList)
