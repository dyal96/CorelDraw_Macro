Attribute VB_Name = "AutoExports"
Sub TIFF_Orginal_Export_All_Pages_in_One_Click()

    Dim expTIF As ExportFilter
    Dim docName As String
    Dim pageName As String
    Dim tifFile As String
    Dim exportFolder As String
    Dim pg As Page
    Dim sr As ShapeRange
    Dim widthIn As Double, heightIn As Double
    Dim sizeStr As String
    Dim i As Integer

    ' Prompt user to pick folder
    exportFolder = CorelScriptTools.GetFolder("Select Export Folder")
    If exportFolder = "" Then
        MsgBox "Export cancelled.", vbExclamation
        Exit Sub
    End If

    ' Check document
    If ActiveDocument Is Nothing Then
        MsgBox "No document open!", vbExclamation
        Exit Sub
    End If

    ' Clean document name
    docName = ActiveDocument.Name
    If InStrRev(docName, ".") > 0 Then
        docName = Left(docName, InStrRev(docName, ".") - 1)
    End If

    ' Set alignment reference
    ActiveDocument.ReferencePoint = cdrCenter

    ' Loop through all pages
    For i = 1 To ActiveDocument.Pages.Count
        Set pg = ActiveDocument.Pages(i)
        pg.Activate

        ' Select all objects on the page
        pg.Shapes.All.CreateSelection
        If pg.Shapes.Count = 0 Then GoTo SkipPage

        Set sr = ActiveSelectionRange

        ' Get size in inches (CorelDRAW units are inches)
        widthIn = Round(sr.SizeWidth, 2)
        heightIn = Round(sr.SizeHeight, 2)
        sizeStr = widthIn & "x" & heightIn

        ' Center align
        sr.AlignAndDistribute 3, 3, 2, 0, False, 2

        ' Build file name
        pageName = "Page" & pg.Index
        tifFile = exportFolder & "\" & sizeStr & "_" & docName & "_" & pageName & ".tif"

        ' Export to TIFF
        Set expTIF = ActiveDocument.ExportBitmap(tifFile, cdrTIFF, cdrSelection, cdrRGBColorImage)
        expTIF.Finish

SkipPage:
    Next i

    MsgBox "All pages exported as TIFF with size in inches.", vbInformation

End Sub

