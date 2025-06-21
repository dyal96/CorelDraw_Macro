Attribute VB_Name = "QuickExport"
Sub QuickTIFF()
    Dim expTIF As ExportFilter
    Dim docName As String
    Dim pageName As String
    Dim tifFile As String
    Dim exportFolder As String
    Dim shellApp As Object

    ' Check if a document is open
    If ActiveDocument Is Nothing Then
        MsgBox "No document open!", vbExclamation
        Exit Sub
    End If

    ' Check if any objects are selected
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "No objects selected!", vbExclamation
        Exit Sub
    End If

    ' Export folder
    exportFolder = "C:\Users\Dell\Desktop\TIFF_Exports"

    ' Get clean document name
    docName = ActiveDocument.Name
    If InStrRev(docName, ".") > 0 Then
        docName = Left(docName, InStrRev(docName, ".") - 1)
    End If

    ' Build file name: DocumentName_PageX.tif
    pageName = "Page" & ActivePage.Index
    tifFile = exportFolder & "\" & docName & "_" & pageName & ".tif"

    ' Export manually selected objects to TIFF (no compression property set)
    Set expTIF = ActiveDocument.ExportBitmap(tifFile, cdrTIFF, cdrSelection, cdrRGBColorImage)
    expTIF.Finish

    MsgBox "TIFF exported to: " & tifFile, vbInformation
End Sub


Sub QuickJPG()
    Dim expFilter As ExportFilter
    Dim filePath As String
    Dim docName As String
    Dim pageName As String
    Dim fullFileName As String
    Dim exportFolder As String

    ' Export folder
    exportFolder = "C:\Users\Dell\Desktop\JPG_Exports"

    ' Check if document and selection exist
    If ActiveDocument Is Nothing Then
        MsgBox "No document open!", vbExclamation
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "No objects selected!", vbExclamation
        Exit Sub
    End If

    ' Ensure folder exists
    If Dir(exportFolder, vbDirectory) = "" Then
        MsgBox "Export folder not found: " & exportFolder, vbCritical
        Exit Sub
    End If

    ' Get document name (remove extension if present)
    docName = ActiveDocument.Name
    If InStrRev(docName, ".") > 0 Then
        docName = Left(docName, InStrRev(docName, ".") - 1)
    End If

    pageName = "Page" & ActivePage.Index
    fullFileName = exportFolder & "\" & docName & "_" & pageName & ".jpg"

    ' Export selected shapes as JPG using original size
    Set expFilter = ActiveDocument.ExportBitmap(fullFileName, cdrJPEG, cdrSelection, cdrRGBColorImage)

    With expFilter
        .Compression = 80         ' 80% quality
        .Smoothing = 0
        .Optimized = True
        .Progressive = False
        .Finish
    End With

    MsgBox "Exported to: " & fullFileName, vbInformation
End Sub

