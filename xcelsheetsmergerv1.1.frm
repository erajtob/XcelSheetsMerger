VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} xcelmerger 
   Caption         =   "XcelSheetsMerger"
   ClientHeight    =   5412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6972
   OleObjectBlob   =   "xcelsheetsmergerv1.1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "xcelmerger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
    offSetRowsBox.Enabled = True
    offSetRowsBox.BackColor = RGB(255, 255, 255)
    Else
    offSetRowsBox.Enabled = False
    offSetRowsBox.BackColor = RGB(232, 232, 232)
    End If
End Sub

Private Sub Label6_Click()
    ActiveWorkbook.FollowHyperlink _
    Address:="https://github.com/erajtob/XcelSheetsMerger"
End Sub


Private Sub mergeButton_Click()
'ErajExcelMerger
    Dim i As Integer
    Dim xRows As Integer
    Dim yCol As Integer
    Dim offSetL As Integer
    Dim sCount As Integer
    Dim xTCount As Variant
    Dim xWs As Worksheet
    Dim cWs As Worksheet
    Dim NewName As String
    Dim Selected_Sheets As String
    Dim listLoop As Integer
    Dim Exclude() As String
    Dim xClude As String
    Dim Delim As String
    Dim chckBox As Boolean
    Dim offSetRows As Integer
    On Error Resume Next
LInput:
    xTCount = xTCountBox.Value
    If TypeName(xTCount) = "Boolean" Then Exit Sub
    If Not IsNumeric(xTCount) Then
        MsgBox "Only can enter number", , "Merger for Excel"
        GoTo LInput
    End If
    Application.ScreenUpdating = False
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
    sCount = Sheets.Count - (UBound(Exclude) - LBound(Exclude) + 2)
    chckBox = Me.CheckBox1.Value
    'offSetRow count
    offSetRows = Me.offSetRowsBox.Value - 1
    'Outer Loop to keep sheet count to determine 1st paste incase of offset
    For offSetL = sCount To 0 Step -1
        'Inner Loop to iterate through worksheets
        For Each xWs In ThisWorkbook.Sheets
            'InStr to exclude sheets from selected sheets to exclude
            If InStr(1, xClude, Delim & xWs.Name & Delim, vbTextCompare) = 0 Then
                'Offset requires first paste to be positive offset thus splitting with IF statement
                If chckBox = True And offSetL = sCount Then
                    xWs.Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy
                           cWs.Cells(cWs.UsedRange.Cells(cWs.UsedRange.Count).Row + 1, 1).PasteSpecial Paste:=xlPasteValues
                    offSetL = offSetL - 1
                ElseIf chckBox = True And offSetL < sCount Then
                    xWs.Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy
                           cWs.Cells(cWs.UsedRange.Cells(cWs.UsedRange.Count).Row - offSetRows, 1).PasteSpecial Paste:=xlPasteValues
                    offSetL = offSetL - 1
                'No Offset Code Run
                ElseIf chckBox = False Then
                    xWs.Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy
                           cWs.Cells(cWs.UsedRange.Cells(cWs.UsedRange.Count).Row + 1, 1).PasteSpecial Paste:=xlPasteValues
                    offSetL = offSetL - 1
                End If
            End If
        Next xWs
    Next
    'Offset row for the last paste
    If chckBox = True Then
        xRows = cWs.UsedRange.Rows.Count
        yRows = cWs.UsedRange.Columns.Count
        cWs.Rows(xRows & ":" & xRows - offSetRows).EntireRow.Delete
    End If
    Application.ScreenUpdating = True
    MsgBox "All Sheets Merged.", vbOKOnly, "Done"
End Sub

Private Sub UserForm_Initialize()
    Dim J As Long
    Dim K As Worksheet
    xcelmerger.Caption = "XcelSheetsMerger " & versionNo.Caption
    offSetRowsBox.BackColor = RGB(232, 232, 232)
        Me.sWsBox.Clear
        For J = 1 To Sheets.Count
            Me.sWsBox.AddItem Sheets(J).Name
        Next
        
    For Each K In Worksheets
        Me.ListBox1.AddItem K.Name
        Next K
End Sub



