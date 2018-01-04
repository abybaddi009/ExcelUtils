# Welcome to ExcelUtils!
Here you will find a curated list of Excel Functions.

Images
-------------

insertPicturesAsLinks():

This sub procedure will insert a picture from `picpath` to your cell specified by `cellAddress` and hyperlink each picture with the links given in the `fileArray`

Syntax:

    insertPicturesAsLinks(picpath As String, cellAddress As String, fileArray As Variant)

Example:

    fileArray = Array("Folder\File1.pdf", "Folder\File2.docx", "Folder\File3.pdf")
    insertPicturesAsLinks Application.ActiveWorkbook.Path & "\link.png", "A1", fileArray
