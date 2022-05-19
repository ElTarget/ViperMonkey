LibreOffice Macros
==================

ViperMonkey uses LibreOffice macros to dump Excel sheet contents and
Word tables using LibreOffice run from the command line. This
directory contains the LibreOffice macros for doing this.

For ViperMonkey to use these macros in LibreOffice the macro file will
need to be copied to the following directory:

```
~/.config/libreoffice/4/user/basic/Standard/
```

Note: If you already have standard macros defined in LibreOffice you
will need to go into the LibreOffice macro editor GUI and paste the
ViperMonkey macros in with your existing Standard macros, otherwise
doing the ViperMonkey file copy will wipe out your existing macros.

Dumping Excel Sheets With Macro
-------------------------------

To dump all sheets from an Excel file from the command line run the
following:

```
libreoffice --invisible --nofirststartwizard --headless --norestore 'macro:///Module1.ExportSheetsFromFile("EXCEL_FILE_NAME_GOES_HERE")'
```

Dumping Word Document Text
--------------------------

To dump Word document text from the command line run the following:

```
libreoffice --headless --convert-to txt WORD_FILE_NAME_GOES_HERE
```
