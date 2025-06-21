Attribute VB_Name = "AutoExports"
Sub JPG_Orginal_Export_All_Pages_in_One_Click()

    Dim expFilter As ExportFilter
    Dim docName As String
    Dim pageName As String
    Dim fullFileName As String
    Dim exportFolder As String
    Dim pg As Page
    Dim sr As ShapeRange
    Dim i As Integer
    Dim widthIn As Double, heightIn As Double
    Dim sizeStr As String

    ' Use CorelDRAW native folder picker
    exportFolder = CorelScriptTools.GetFolder("Select Export Folder")
    If exportFolder = "" Then
        MsgBox "Export cancelled.", vbExclamation
        Exit Sub
    End If

    ' Ensure a document is open
    If ActiveDocument Is Nothing Then
        MsgBox "No document open!", vbExclamation
        Exit Sub
    End If

    docName = ActiveDocument.Name
    If InStrRev(docName, ".") > 0 Then
        docName = Left(docName, InStrRev(docName, ".") - 1)
    End If

    ActiveDocument.ReferencePoint = cdrCenter

    For i = 1 To ActiveDocument.Pages.Count
        Set pg = ActiveDocument.Pages(i)
        pg.Activate

        ' Select all shapes on the current page
        pg.Shapes.All.CreateSelection
        If pg.Shapes.Count = 0 Then GoTo SkipPage

        Set sr = ActiveSelectionRange

        ' Get width and height in inches (CorelDRAW units are inches by default)
        widthIn = Round(sr.SizeWidth, 2)
        heightIn = Round(sr.SizeHeight, 2)
        sizeStr = widthIn & "x" & heightIn

        ' Align to center
        sr.AlignAndDistribute 3, 3, 2, 0, False, 2

        ' Filename
        pageName = "Page" & pg.Index
        fullFileName = exportFolder & "\" & sizeStr & "_" & docName & "_" & pageName & ".jpg"

        ' Export to JPG
        Set expFilter = ActiveDocument.ExportBitmap(fullFileName, cdrJPEG, cdrSelection, cdrRGBColorImage)
        With expFilter
            .Compression = 0
            .Smoothing = 0
            .Optimized = True
            .Progressive = False
            .Finish
        End With

SkipPage:
    Next i

    MsgBox "All pages exported with size in inches in filename.", vbInformation

End Sub

