"""
Use idYes to save the document, or use idNo to close the document without saving
If you use idYes, you'll need to provide a reference to a file to save to in the second
parameter (SavingIn).If the file has not been saved, save it to a specific file path.
"""

import win32com.client

app = win32com.client.Dispatch('InDesign.Application')

idYes = 2036691744
if app.Documents.Count > 0:
    myDocument = app.Documents.Item(1)
    if not myDocument.Saved:
        myFile = r'TestDocument.indd'
        myDocument.Close(idYes, myFile)
    else:
        myDocument.Close(idYes)
