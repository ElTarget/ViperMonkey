#!/usr/bin/env pypy3

"""@package vipermonkey.export_doc_text Export the document
text/tables of a Word document with LibreOffice.  This is Python 3.

"""

import subprocess
import time
import argparse
import json
import os
import signal

from core.utils import safe_str_convert

###################################################################################################
def is_word_file(fname):
    """Check to see if the given file is a Word file.

    @param fname (str) The path of the file to check.

    @return (bool) True if the file is a Word file, False if not.

    """

    # Easy case. The file commands tells us this is a Word file.
    typ = subprocess.check_output(["file", fname])
    if ((b"Microsoft Office Word" in typ) or
        (b"Word 2007+" in typ) or
        (b"Microsoft OOXML" in typ)):
        return True

    # Harder case. This could be a Word file saved as XML.
    if (b"XML 1.0 document" not in typ):
        return False
    contents = None
    try:
        f = open(fname, "r")
        contents = f.read()
        f.close()
    except IOError:
        return False
    return ('<?mso-application progid="Word.Document"?>' in contents)

###################################################################################################
def get_tables(document):
    """Get the text tables embedded in the Word doc.

    @param document (Writer) LibreOffice component containing the
    document.

    @return (list) List of 2D arrays containing text content of all
    cells in all text tables of the document

    """

    data_array_list = []

    text_tables = document.getTextTables()
    table_count = 0
    while table_count < text_tables.getCount():
        data_array_list.append(text_tables.getByIndex(table_count).getDataArray())
        table_count += 1

    return data_array_list

###########################################################################
def get_text(fname):
    """Get the text of a Word document using LibreOffice.

    @param fname (str) The name of the Word file.

    @return (str) The document text on success, None on failure.

    """

    # Run headless LibreOffice to save the Word document as text.
    cmd = ["libreoffice", "--invisible",
           "--nofirststartwizard", "--headless",
           "--norestore", "--convert-to", "txt", fname]
    out = subprocess.check_call(cmd)

    # Get the contexts of the text file.
    out_fname = fname + ".txt"
    r = None
    try:

        # Read text.
        f = open(out_fname, "rb")
        r = safe_str_convert(f.read())
        f.close()

        # Clean up text file.
        os.remove(out_fname)
        
    except IOError:
        pass

    # Done.
    return r

###########################################################################
## Main Program
###########################################################################
if __name__ == '__main__':
    arg_parser = argparse.ArgumentParser(description="export text from various properties in a Word "
                                         "document via the LibreOffice API")
    arg_parser.add_argument("--tables", action="store_true",
                            help="export a list of 2D lists containing the cell contents"
                            "of each text table in the document")
    arg_parser.add_argument("--text", action="store_true",
                            help="export a string containing the document text")
    arg_parser.add_argument("-f", "--file", action="store", required=True,
                            help="path to the word doc")
    args = arg_parser.parse_args()

    # Make sure this is a word file.
    if (not is_word_file(args.file)):
    
        # Not Word, so no text.
        exit()

    # Get doc text or tables.
    if args.text:
        print((get_text(args.file)))
    elif args.tables:
        print((json.dumps(get_tables(document))))
