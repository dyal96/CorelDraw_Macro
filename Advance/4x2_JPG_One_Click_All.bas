Attribute VB_Name = "AutoExports"
Sub 4x2_JPG_One_Click_All()

    Dim expFilter As ExportFilter
    Dim filePath As String
    Dim docName As String
    Dim pageName As String
    Dim fullFileName As String
    Dim exportFolder As String
    Dim pg As Page
    Dim sr As ShapeRange
    Dim i As Integer
    Dim widthStr As String, heightStr As String

    ' Export folder (ensure this path exists)
    exportFolder = "C:\Users\Dell\Desktop\JPG_Exports"

    If Dir(exportFolder, vbDirectory) = "" Then
        MsgBox "Export folder not found: " & exportFolder, vbCritical
        Exit Sub
    End If

    ' Ensure document exists
    If ActiveDocument Is Nothing Then
        MsgBox "No document open!", vbExclamation
        Exit Sub
    End If

    ' Get document name without extension
    docName = ActiveDocument.Name
    If InStrRev(docName, ".") > 0 Then
        docName = Left(docName, InStrRev(docName, ".") - 1)
    End If

    ' Set reference point for alignment
    ActiveDocument.ReferencePoint = cdrCenter

    ' Define export size
    widthStr = "48"
    heightStr = "24"

    ' Loop through all pages
    For i = 1 To ActiveDocument.Pages.Count
        Set pg = ActiveDocument.Pages(i)
        pg.Activate

        ' Select all shapes on the current page
        pg.Shapes.All.CreateSelection
        If pg.Shapes.Count = 0 Then
            ' Skip empty pages
            GoTo SkipPage
        End If

        Set sr = ActiveSelectionRange

        ' Resize and center align
        sr.SetSize CSng(widthStr), CSng(heightStr)
        sr.AlignAndDistribute 3, 3, 2, 0, False, 2

        ' Build export file name with size
        pageName = "Page" & pg.Index
        fullFileName = exportFolder & "\" & widthStr & "x" & heightStr & "_" & docName & "_" & pageName & ".jpg"

        ' Export selected shapes as JPG
        Set expFilter = ActiveDocument.ExportBitmap(fullFileName, cdrJPEG, cdrSelection, cdrRGBColorImage)
        With expFilter
            .Compression = 0         ' 100% quality
            .Smoothing = 0
            .Optimized = True
            .Progressive = False
            .Finish
        End With

SkipPage:
    Next i

    MsgBox "All pages exported as JPG with size in filename!", vbInformation

End Sub

