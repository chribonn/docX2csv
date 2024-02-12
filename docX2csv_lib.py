import zipfile
import xml.etree.ElementTree as ET


#constants
VERSION = "0.2"
NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NW_URI_TAG = '{' + NS_URI + '}'
XML_DOC_PATH = 'word/document.xml'

def extract_document_xml(docx_file, tmp_dir):
    with zipfile.ZipFile(docx_file, 'r') as zip_ref:
        zip_ref.extract(XML_DOC_PATH, tmp_dir)
            
    return tmp_dir



def get_styles(docx_file):
    """
    Returns a list of all the Styles defined in the document. This is used to populate the drop down list
    Returns the number of styles found. This is used for a progress bar.
    

    Args:
        docx_file (file_path): the file that is being processed

    Returns:
        list: _Styles defined in the document
        int: the number of style tags found
    """
    styles = []
    i = 0

    tree = ET.parse(docx_file)
    root = tree.getroot()

    ET.register_namespace("w", NS_URI)
    ns = {"w": NS_URI}

    for x in root.findall('.//w:pStyle', ns):
        i += 1
        style = x.get(NW_URI_TAG + 'val')
        if style not in styles:
            styles.append(style)
            
    return sorted(styles), i


def proc_pPr_pStyle(branch, srchStyles):
    """
    Returns two parameters:
        * The matched style or None
        * A boolean to indicate whether the style tag was found

    Args:
        branch : The XML branch being analysed)
        srchStyles (list): The styles being looked for

    Returns:
        style: The style that was matched
        styletag_found: Boolean 
    """
    # print (branch.tag, branch.attrib)
    name = branch.find(NW_URI_TAG + 'pStyle')
    if name is None:
        style = None
        styletag_found = False
    else:
        styletag_found = True
        style = name.get(NW_URI_TAG + 'val')
        # Process only if this is a style we are interested in
        if style not in srchStyles:
            style = None
        
    return style, styletag_found


def proc_pPr_sectPr(branch):
    return branch.find(NW_URI_TAG + 'sectPr') is not None
        

def proc_r_t(branch):
    text = ''
    name = None
    for y in branch:
        if y.tag == NW_URI_TAG + 'r':
            name = y.find(NW_URI_TAG + 't')
            text += '' if name is None else name.text
    
    # if text is null check for the xml:space="preserve" in which case add a space
    if name is not None and (text == '' or text is None):
        if name.get('{http://www.w3.org/XML/1998/namespace}space') == "preserve":
            text = ' '

    return '' if text is None else text


def proc_r_lastRenderedPageBreak(branch):
    for y in branch.iter(NW_URI_TAG + 'r'):
        name = y.find(NW_URI_TAG + 'lastRenderedPageBreak')
        if name is not None:
            return True

    return False


def proc_r_br(branch):
    for y in branch.iter(NW_URI_TAG + 'r'):
        name = y.find(NW_URI_TAG + 'br')
        if name is None:
            break
        
        if name.get(NW_URI_TAG + 'type') == "page":
            return True

    return False


def proc_pPr_sectBr(branch):
    """
    This module looks fof a section break on the next page. This is identified by the lack of a
    <w:type w:val="continuous"/>

    Args:
        branch :  xml node being analysed

    Returns:
        Boolean : New page or not
    """
    sect_pPr = branch.find(NW_URI_TAG + 'pPr')
    if sect_pPr is None:
        return False

    sect_sectPr = sect_pPr.find(NW_URI_TAG + 'sectPr')
    if sect_sectPr is None:
        return False
    
    if sect_sectPr.find(NW_URI_TAG + 'type') is None:
        return True

    return False
    

def page_break(branch):
    """
    Returns whether there is a page break (rendered, manual or section) in this node.
        there may be multiple tags all indicating a page break but these need to be counted as one.

    Args:
        branch (xml branch): The <w:p node

    Returns:
        Boolean: if the node was found
    """
    if proc_r_lastRenderedPageBreak(branch) or proc_r_br(branch) or proc_pPr_sectBr(branch):
        return True
    
    return False
