import docx

def list_sections_in_word_document(file_path):
    """
    Reads in a word document and lists all the sections in the document.

    Args:
    file_path (str): The file path of the word document.

    Returns:
    list: A list of section names in the document.
    """
    sections = []
    doc = docx.Document(file_path)
    for paragraph in doc.paragraphs:
        if paragraph.style.name == 'Heading 1':
            sections.append(paragraph.text)
    return sections

# Example usage
file_path = 'D:\chrib\Downloads\Cloud_Guide.docx'
sections_list = list_sections_in_word_document(file_path)
print(sections_list)