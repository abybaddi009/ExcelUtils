Sub insertPicture(picpath As String, cellAddress As String, fileArray As Variant)
    '----------------------------------------------------------------------------
    ' "THE BURGER-WARE LICENSE" (Revision 42):
    ' <abybaddi009 gmail.com> wrote this code. As long as you retain this notice you
    ' can do whatever you want with this stuff. If we meet some day, and you think
    ' this stuff is worth it, you can buy me a burger in return. ;-) -Abhishek Baddi
    '----------------------------------------------------------------------------
    ' Usage: 
    ' insertPicture Application.ActiveWorkbook.Path & "\link.png", "A1", fileArray
    '----------------------------------------------------------------------------
    ' example:
    ' Sub PutLinksInACell()
    '    Dim rangeAddress As String
    '    Dim fileArray
    '    fileArray = Array("Folder\File1.pdf", "Folder\File2.docx", "Folder\File3.pdf")
    '    'rangeAddress = Selection.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    '    insertPicture Application.ActiveWorkbook.Path & "\link.png", "B9", fileArray
    ' End Sub

    Dim spacing As Long, size As Long
    
    size = Range(cellAddress).Font.size
    spacing = size * 0.2
    
    x_coor = Range(cellAddress).Cells(1, 1).Left
    y_coor = Range(cellAddress).Cells(1, 1).Top
    
    For i = 1 To 3
        ActiveSheet.Pictures.Insert(picpath).Select
        With Selection
            With .ShapeRange
                .LockAspectRatio = msoTrue
                .Height = size
            End With
            .Left = x_coor + 5
            .Top = y_coor + size * (i - 1) + spacing * i
            .Placement = 1
            .PrintObject = True
        End With
        
        ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), Address:= _
            fileArray(i - 1)
        
        Range(cellAddress).Select
    Next
    Range(cellAddress).HorizontalAlignment = xlLeft
    Range(cellAddress).VerticalAlignment = xlTop
End Sub
