# Habib, P
# 02/21/2021
# A program that extracts the document xml from a word doc

import zipfile
from xml.dom import minidom
import shutil

def extract_xml(docx_filename):
    """Extract document.xml from the unziped .docx file and pretty print it.

    Keyword arguments:
    docx_filename: str -- the filename of the docx file
    """
    with zipfile.ZipFile(docx_filename) as zip_ref:
        # Unzip the document into a temp folder
        zip_ref.extractall("temp")
        # Use minidom to parse the xml so we can print it later
        xml_content = minidom.parse('temp/word/document.xml')

    # Create a new document in this direrctory
    with open('document.xml', 'w') as f:
        # Pretty print the XML so that the tags are indented an on their own lines
        f.write(xml_content.toprettyxml())

    # Remove the temp directory with shutil (Shell Util)
    shutil.rmtree('temp', ignore_errors=True)

