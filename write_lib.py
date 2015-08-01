from docx import Document
from docx.shared import Inches

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
        runi = 0
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
            elif value == 'boldnitalic':
                if not previousPara:
                    paragraph = documentReal.add_paragraph('')
                runi += 1
                prun[runi] = paragraph.add_run(key)
                prun[runi].bold = True
                prun[runi].italic = True
            elif value == 'bulletpoint':
                paragraph = documentReal.add_paragraph('')
                prun = paragraph.add_run('* ' + key)
            elif value == 'image':
                document.add_picture(key, width=Inches(1.0))

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
            documentReal.save(folder + '\\save.docx')
        else:
            documentReal.save(folder + '/save.docx')
