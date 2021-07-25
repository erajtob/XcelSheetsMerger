Attribute VB_Name = "XcelMerger"
Sub Combine()
'ErajExcelMerger
    Dim i As Integer
    Dim xTCount As Variant
    Dim xWs As Worksheet
    Dim cWs As Worksheet
    Dim zWs As Variant
    Dim NewName As String
    'Dim Exclude() As String
    On Error Resume Next
LInput:
    xTCount = Application.InputBox("The number of title/header rows", "", "1")
    If TypeName(xTCount) = "Boolean" Then Exit Sub
    If Not IsNumeric(xTCount) Then
        MsgBox "Only can enter number", , "Merger for Excel"
        GoTo LInput
    End If
    
    'Add extra Sheet to workbook for the Merged dump
    Set cWs = ActiveWorkbook.Worksheets.Add(Sheets(1))

    'Input for Combined Sheet
    NewName = InputBox("What Do you Want to Name the Combined Sheet ?")
    cWs.Name = NewName

    'Copy Title and Paste on A1 of Merged Sheet
    Worksheets(4).Range("A1").EntireRow.Copy Destination:=cWs.Range("A1")
    
    'Exclude = Split("Sheet1,Product", ",")

    'Switch Row - 1 to + 1 for 1st entry in Line 23
    For Each xWs In ThisWorkbook.Sheets
        If InStr(1, " " & cWs.Name & " Sheet1 Product ", " " & xWs.Name & " ", vbTextCompare) = 0 Then
            zWs = xWs.Name
            xWs.Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy
                   cWs.Cells(cWs.UsedRange.Cells(cWs.UsedRange.Count).Row + 1, 1).PasteSpecial Paste:=xlPasteValues
        End If
    Next xWs
    
End Sub


