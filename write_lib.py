try:
    from docx import Document
except ImportError:
    print("DOCX conversion Unavailable")
try:
    from docx.shared import Inches
except ImportError:
    print("DOCX Conversion Unavailable")
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("HTML Conversion Unavailable")
try:
    import markdown
except ImportError:
    print("HTML Conversion Unavailable")
    print("Markdown Conversion Unavailable")
import re
import collections
import os

def strip_html(stringIn):
    stringIn = str(stringIn)
    stripped_item = re.sub('<[^<]+?>', '', stringIn)
    return stripped_item

def html_to_json(htmlString,outFile=None):
    """Convert HTML to pywriter's JSON syntax"""
    htmlString = markdown.markdown(htmlString)
    htmlString = BeautifulSoup(htmlString, 'html.parser')
    finalDict = {}
    finalDict['document'] = collections.OrderedDict()
    for item in htmlString:
        stringItem = str(item)
        if '<h1' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'title'
        if '<h2' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'subtitle'
        if '<h3' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'tagline'
        if '<b' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'bold'
        if '<strong' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'bold'
        if '<i' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'italic'
        if '<em' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'italic'
        if '<li' in stringItem:
            stripped_item = strip_html(item)
            finalDict['document'][stripped_item] = 'bulletpoint'
        if outFile:
            finalDict['outFile'] = outFile
        else:
            folder = os.getcwd()
            if '\\' in folder:
                folder = folder.replace('\\','/')
            finalDict['outFile'] = folder + '/save.docx'
    return finalDict

def generate_docx(dictIn):
    """Takes either a list for unformatted text, or a dict with formatting options"""
    try:
        docStruct = dictIn['document']
    except KeyError:
        docStruct = list(dictIn)
    documentReal = Document()

    # Try and build the document from a dict
    try:
        previousPara = False
        for key, value in docStruct.items():
            if value == 'paragraph':
                if not previousPara:
                    paragraph = documentReal.add_paragraph('')
                paragraph.add_run(key)
                previousPara = True
            elif value == 'title':
                previousPara = False
                documentReal.add_heading(key, level=0)
            elif value == 'subtitle':
                previousPara = False
                documentReal.add_heading(key, level=1)
            elif value == 'tagline':
                previousPara = False
                documentReal.add_heading(key, level=2)
            elif value == 'bold':
                if not previousPara:
                    paragraph = documentReal.add_paragraph('')
                paragraph.add_run(key).bold = True
            elif value == 'italic':
                if not previousPara:
                    paragraph = documentReal.add_paragraph('')
                paragraph.add_run(key).italic = True
            elif value == 'bulletpoint':
                paragraph = documentReal.add_paragraph('')
                prun = paragraph.add_run('* ' + key)
            elif value == 'image':
                documentReal.add_picture(key, width=Inches(1.0))

    # If not a dict, process it as a list
    except AttributeError:
        for item in docStruct:
            documentReal.add_paragraph(item)
    # Save the file from the dict value
    try:
        documentReal.save(dictIn['outFile'])
    except KeyError:
        # Or save into the current working directory.
        folder = os.getcwd()
        if '\\' in folder:
            folder = folder.replace('\\', '/')
        documentReal.save(folder + '/save.docx')

def generate_markdown(dictIn):
    """Takes either a list for unformatted text, or a dict with formatting options"""
    try:
        docStruct = dictIn['document']
    except AttributeError:
        docStruct = list(dictIn)
