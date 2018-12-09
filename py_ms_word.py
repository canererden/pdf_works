"""
author: Caner Erden
"""
import os
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path= os.getcwd()+'\\word_samples'):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    contentToRead = ["header2.xml", "document.xml", "footer2.xml"]
    paragraphs = []

    for xmlfile in contentToRead:
        xml_content = document.read('word/{}'.format(xmlfile))
        tree = XML(xml_content)
        for paragraph in tree.getiterator(PARA):
            texts = [node.text
                     for node in paragraph.getiterator(TEXT)
                     if node.text]
            if texts:
                textData = ''.join(texts)
                if xmlfile == "footer2.xml":
                    extractedTxt = "Footer : " + textData
                elif xmlfile == "header2.xml":
                    extractedTxt = "Header : " + textData
                else:
                    extractedTxt = textData

                paragraphs.append(extractedTxt)
    document.close()
    return '\n\n'.join(paragraphs)


print(get_docx_text())