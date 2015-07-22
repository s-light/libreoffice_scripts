#!/usr/bin/python3

import os.path
import uno
from com.sun.star.beans import PropertyValue

# https://wiki.openoffice.org/wiki/Python_as_a_macro_language#Error_handling_and_debugging

# https://wiki.openoffice.org/wiki/Framework/Article/Filter
docTypes = {
    'com.sun.star.text.TextDocument' : {
        'lo':{
            'name': 'writer',
            'extension': 'odt',
        },
        'ms97':{
            'extension': 'doc',
            'filter': 'MS Word 97',
        },
        'msXML':{
            'extension': 'docx',
            'filter': 'MS Word 2007 XML',
        },
        'pdf':{
            'extension': 'pdf',
            'filter': 'writer_pdf_Export',
        },

    },
    'com.sun.star.sheet.SpreadsheetDocument' : {
        'lo':{
            'name': 'calc',
            'extension': 'ods',
        },
        'ms97':{
            'extension': 'xls',
            'filter': 'MS Excel 97',
        },
        'msXML':{
            'extension': 'xlsx',
            'filter': 'Calc MS Excel 2007 XML',
        },
        'pdf':{
            'extension': 'pdf',
            'filter': 'calc_pdf_Export',
        },
    },
    'com.sun.star.presentation.PresentationDocument' : {
        'lo':{
            'name': 'impress',
            'extension': 'odp',
        },
        'ms97':{
            'extension': 'ppt',
            'filter': 'MS PowerPoint 97',
        },
        'msXML':{
            'extension': 'pptx',
            'filter': 'Impress MS PowerPoint 2007 XML',
        },
        'pdf':{
            'extension': 'pdf',
            'filter': 'impress_pdf_Export',
        },
    },
}

def getDocType(currentDoc):
    """test document type and return type converting information"""
    resultType = None
    for docType in docTypes:
        # print("docType: {}".format(docType))
        if currentDoc.supportsService(docType):
            resultType = docTypes[docType]
    return resultType


def compose_new_URL(doc_type, dest_type, url_current, url_addition):
    """exchange extension and add url_addition to filename"""
    # if !url_current:
    #     url_current = currentDoc.getLocation()
    # print("url_current: {}".format(url_current))

    if not url_addition:
        url_addition = ''

    url_base, url_ext = os.path.splitext(url_current)
    # print("url_base: {}".format(url_base))
    # print("url_ext: {}".format(url_ext))

    # swicht extension
    url_ext_new = url_ext.replace(
        doc_type['lo']['extension'],
        doc_type[dest_type]['extension']
    )
    # print("url_ext_new: {}".format(url_ext_new))

    url_new = url_base + url_addition + url_ext_new
    # print("url_new: {}".format(url_new))

    return url_new


def convert_dict_to_PropertyValue_List(sourceDict):
    properties_List = []
    for (key, value) in sourceDict.items():
        p = PropertyValue()
        p.Name = key
        p.Value = value
        properties_List.append(p)
    return properties_List


# def save_as(dest_type=None, url_addition=None, additional_properties=None ):
def save_as(\
    dest_type=None,\
    currentDoc=None,\
    doc_type=None,\
    url_current=None,\
    url_addition=None,\
    additional_properties=None\
):
    if dest_type:
        if not currentDoc:
            currentDoc = XSCRIPTCONTEXT.getDocument()
        if not doc_type:
            doc_type = getDocType(currentDoc)
        if not url_current:
            url_current = currentDoc.getLocation()
        # currentDoc = XSCRIPTCONTEXT.getDocument()
        # doc_type = getDocType(currentDoc)
        # url_current = currentDoc.getLocation()

        url_new = compose_new_URL(
            doc_type, # doc_type
            dest_type, # dest_type,
            url_current, # url_current,
            url_addition # url_addition
        )

        properties=[]

        p = PropertyValue()
        p.Name = 'FilterName'
        p.Value = doc_type[dest_type]['filter']
        properties.append(p)

        p = PropertyValue()
        p.Name = 'Overwrite'
        p.Value = True
        properties.append(p)

        p = PropertyValue()
        p.Name = 'InteractionHandler'
        p.Value = ''
        properties.append(p)

        if additional_properties:
            properties.extend(additional_properties)

        currentDoc.storeToURL(url_new, tuple(properties))
    else:
        print("no type given. i cant work...")

##########################################

def save_as_PDF(*args):
    """save current document as PDF file"""
    save_as(dest_type = 'pdf')

def save_as_PDF_HiRes(*args):
    """save current document as HiRes PDF file

        this means all images are used with their orignial resolution.
    """
    # https://wiki.openoffice.org/wiki/API/Tutorials/PDF_export
    filter_data = {
        # General properties
        'UseLosslessCompression' : False,
        'Quality' : 100,
        'ReduceImageResolution' : False,
        'MaxImageResolution' : 1200, # 75, 150, 300, 600, 1200
        'ExportBookmarks' : True,
        'EmbedStandardFonts' : True,
        # Links
        'ExportBookmarksToPDFDestination' : True,
        'ConvertOOoTargetToPDFTarget' : True,
        'ExportLinksRelativeFsys' : True,
        'PDFViewSelection' : 2,
        # Security
        # ...
    }

    additional_properties=[]
    p = PropertyValue()
    p.Name = 'FilterData'
    p.Value = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        tuple(convert_dict_to_PropertyValue_List(filter_data)))
    additional_properties.append(p)

    save_as(
        'pdf',
        url_addition='__hires',
        additional_properties = additional_properties
    )

def save_as_PDF_600dpi(*args):
    """save current document as PDF file

        images will be reduced to 600dpi (lossless compression)
    """
    # https://wiki.openoffice.org/wiki/API/Tutorials/PDF_export
    filter_data = {
        # General properties
        'UseLosslessCompression' : False,
        'Quality' : 100,
        'ReduceImageResolution' : True,
        'MaxImageResolution' : 600, # 75, 150, 300, 600, 1200
        'ExportBookmarks' : True,
        'EmbedStandardFonts' : True,
        # Links
        'ExportBookmarksToPDFDestination' : True,
        'ConvertOOoTargetToPDFTarget' : True,
        'ExportLinksRelativeFsys' : True,
        'PDFViewSelection' : 2,
        # Security
        # ...
    }

    additional_properties=[]
    p = PropertyValue()
    p.Name = 'FilterData'
    p.Value = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        tuple(convert_dict_to_PropertyValue_List(filter_data)))
    additional_properties.append(p)

    save_as(
        'pdf',
        url_addition='__600dpi',
        additional_properties = additional_properties
    )

def save_as_PDF_75dpi(*args):
    """save current document as PDF file

        images will be reduced to 75dpi (JPG compression)
    """
    # https://wiki.openoffice.org/wiki/API/Tutorials/PDF_export
    filter_data = {
        # General properties
        'UseLosslessCompression' : True,
        'Quality' : 100,
        'ReduceImageResolution' : True,
        'MaxImageResolution' : 75, # 75, 150, 300, 600, 1200
        'ExportBookmarks' : True,
        'EmbedStandardFonts' : True,
        # Links
        'ExportBookmarksToPDFDestination' : True,
        'ConvertOOoTargetToPDFTarget' : True,
        'ExportLinksRelativeFsys' : True,
        'PDFViewSelection' : 2,
        # Security
        # ...
    }

    additional_properties=[]
    p = PropertyValue()
    p.Name = 'FilterData'
    p.Value = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        tuple(convert_dict_to_PropertyValue_List(filter_data)))
    additional_properties.append(p)

    save_as(
        'pdf',
        url_addition='__75dpi',
        additional_properties = additional_properties
    )

def save_as_PDF_Default(*args):
    """save current document as PDF file

        all configuration options will be set to default
    """
    # https://wiki.openoffice.org/wiki/API/Tutorials/PDF_export
    filter_data = {
        # General properties
        'PageRange' : '',
        # 'Selection' : any,
        'UseLosslessCompression' : False,
        'Quality' : 90,
        'ReduceImageResolution' : False,
        'MaxImageResolution' : 300, # 75, 150, 300, 600, 1200
        'SelectPdfVersion' : 0,
        'UseTaggedPDF' : False,
        'ExportFormFields' : True,
        'FormsType' : 0,
        'AllowDuplicateFieldNames' : False,
        'ExportBookmarks' : True,
        'ExportNotes' : False,
        'ExportNotesPages' : False,
        'IsSkipEmptyPages' : False,
        'EmbedStandardFonts' : False,
        'IsAddStream' : False,
        'Watermark' : '',
        # Initial view
        # ...
        # User interface
        # ...
        # Links
        'ExportBookmarksToPDFDestination' : False,
        'ConvertOOoTargetToPDFTarget' : False,
        'ExportLinksRelativeFsys' : False,
        'PDFViewSelection' : 0,
        # Security
        # ...
    }

    additional_properties=[]
    p = PropertyValue()
    p.Name = 'FilterData'
    p.Value = uno.Any(
        "[]com.sun.star.beans.PropertyValue",
        tuple(convert_dict_to_PropertyValue_List(filter_data)))
    additional_properties.append(p)

    save_as(
        'pdf',
        url_addition='__default',
        additional_properties = additional_properties
    )

##########################################

def save_as_ms97(*args):
    """save current document as Microsoft 97 file"""
    save_as(dest_type = 'ms97')

def save_as_msXML(*args):
    """save current document as Microsoft XML file"""
    save_as(dest_type = 'msXML')

def save_as_ms(*args):
    """save current document as Microsoft 97 & XML file"""
    save_as('ms97')
    save_as('msXML')

##########################################

def save_as_All(*args):
    """save current document as All available file-variants

        this includes:
            Microsoft 97
            Microsoft XML
            PDF HiRes
            PDF 600dpi
            PDF 75dpi
    """
    save_as('ms97')
    save_as('msXML')
    save_as_PDF_HiRes()
    save_as_PDF_600dpi()
    save_as_PDF_75dpi()

def save_as_Multi(*args):
    """save current document as multiple file-variants

        this includes:
            Microsoft 97
            Microsoft XML
            PDF (last used configuration)
    """
    save_as('ms97')
    save_as('msXML')
    save_as('pdf')



##########################################

def test(*args):
    """test different things from this tests"""
    print(42*'~')
    print("args: {}".format(args))
    print(42*'~')
    # print("test the current open document type:")
    #get the doc from the scripting context which is made available to all scripts
    # currentDoc = XSCRIPTCONTEXT.getDocument()
    # myCurrentDocType = getDocType(currentDoc)

    print("save_as_ms97:")
    save_as('ms97')
    # save_as(dest_type = 'ms97')
    print("save_as_msXML:")
    save_as('msXML')
    # save_as_msXML()
    # save_as(dest_type = 'msXML')
    print("save_as_pdf:")
    save_as(dest_type = 'pdf')

    print(42*'~')


# lists the scripts, that shall be visible inside OOo. Can be omitted, if
# all functions shall be visible, however here getNewString shall be suppressed
g_exportedScripts = test,\
    save_as_ms97,\
    save_as_msXML,\
    save_as_ms,\
    save_as_PDF,\
    save_as_PDF_HiRes,\
    save_as_PDF_600dpi,\
    save_as_PDF_75dpi,\
    save_as_Multi,\
    save_as_All,
