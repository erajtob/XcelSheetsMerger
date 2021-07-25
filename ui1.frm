VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ui1 
   Caption         =   "XcelSheetMerger"
   ClientHeight    =   5412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6972
   OleObjectBlob   =   "ui1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ui1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mergeButton_Click()
'ErajExcelMerger
    Dim i As Integer
    Dim xTCount As Variant
    Dim xWs As Worksheet
    Dim cWs As Worksheet
    Dim NewName As String
    Dim Selected_Sheets As String
    Dim listLoop As Integer
    Dim Exclude() As String
    Dim xClude As String
    Dim Delim As String
    On Error Resume Next
LInput:
    xTCount = xTCountBox.Value
    If TypeName(xTCount) = "Boolean" Then Exit Sub
    If Not IsNumeric(xTCount) Then
        MsgBox "Only can enter number", , "Merger for Excel"
        GoTo LInput
    End If
    
    'Add extra Sheet to workbook for the Merged dump
    Set cWs = ActiveWorkbook.Worksheets.Add(Sheets(1))

    'Input for Combined Sheet
    NewName = cWsBox.Value
    cWs.Name = NewName

    'Copy Title and Paste on A1 of Merged Sheet
    Worksheets(sWsBox.Value).Range("A1").EntireRow.Copy Destination:=cWs.Range("A1")
    
    For listLoop = 1 To Me.ListBox1.ListCount
        If Me.ListBox1.Selected(listLoop - 1) Then
            Selected_Sheets = Selected_Sheets & "," & Me.ListBox1.List(listLoop - 1)
        End If
    Next
    Selected_Sheets = Mid(Selected_Sheets, 2)
    
    Delim = ","
    Exclude = Split(Selected_Sheets, ",")
    xClude = Join(Exclude, Delim)
    xClude = Delim & cWs.Name & Delim & xClude & Delim
    Me.debuglab.Caption = xClude
    'Switch Row - 1 to + 1 for 1st entry in Line 23
    For Each xWs In ThisWorkbook.Sheets
        If InStr(1, xClude, Delim & xWs.Name & Delim, vbTextCompare) = 0 Then
            xWs.Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy
                   cWs.Cells(cWs.UsedRange.Cells(cWs.UsedRange.Count).Row + 1, 1).PasteSpecial Paste:=xlPasteValues
        End If
    Next xWs
End Sub



Private Sub UserForm_Initialize()
    Dim J As Long
    Dim K As Worksheet
        Me.sWsBox.Clear
        For J = 1 To Sheets.Count
            Me.sWsBox.AddItem Sheets(J).Name
        Next
        
    For Each K In Worksheets
        Me.ListBox1.AddItem K.Name
        Next K
End Sub

