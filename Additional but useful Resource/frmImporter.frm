VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImporter 
   Caption         =   "Text"
   ClientHeight    =   7130
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9070.001
   OleObjectBlob   =   "frmImporter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fileList As Collection
Private selectedSheets As Collection

'---------------- Initialization ----------------
Private Sub UserForm_Initialize()
    Set fileList = New Collection
    Set selectedSheets = New Collection

    cboAction.Clear
    cboAction.AddItem "COPY"
    cboAction.AddItem "MOVE"
    cboAction.ListIndex = 0

    cboFormat.Clear
    cboFormat.AddItem "XLSX"
    cboFormat.AddItem "CSV"
    cboFormat.AddItem "TXT"
    cboFormat.ListIndex = 0

    chkConvert.Value = False
    txtSavePath.Text = ""
End Sub

'---------------- File Management ----------------
Private Sub btnAdd_Click()
    Dim fd As FileDialog, v
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = True
        .Title = "Select Excel or text files"
        .Filters.Clear
        .Filters.Add "Excel/Text Files", "*.xlsx;*.xls;*.csv;*.txt"
        If .Show <> -1 Then Exit Sub
        For Each v In .SelectedItems
            If Not ItemInCollection(fileList, v) Then fileList.Add v
        Next v
    End With
    RefreshFileList
    LoadSheetList
End Sub

Private Sub btnRemove_Click()
    Dim i As Long
    For i = lstFiles.ListCount - 1 To 0 Step -1
        If lstFiles.Selected(i) Then fileList.Remove i + 1
    Next i
    RefreshFileList
    LoadSheetList
End Sub

Private Sub btnClear_Click()
    Set fileList = New Collection
    RefreshFileList
    lstSheets.Clear
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Long
    Dim allSelected As Boolean

    allSelected = True
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) = False Then
            allSelected = False
            Exit For
        End If
    Next i

    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = Not allSelected
    Next i
End Sub

'---------------- Browse Save As ----------------
Private Sub btnBrowseSave_Click()
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Result.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx," & _
                    "CSV File (*.csv), *.csv," & _
                    "Text File (*.txt), *.txt", _
        Title:="Select Save Location and File Name")
    If savePath <> False Then txtSavePath.Text = savePath
End Sub

'---------------- Import Logic ----------------
Private Sub btnImport_Click()
    If lstSheets.ListCount = 0 Then
        MsgBox "Please load files and select sheets to import.", vbExclamation
        Exit Sub
    End If

    If txtSavePath.Text = "" Then
        MsgBox "Please select a save path using the Browse button.", vbExclamation
        Exit Sub
    End If

    Dim i As Long
    Dim targetWB As Workbook
    Dim sheetInfo As Variant, filePath As String, sheetName As String
    Dim wbSrc As Workbook, ws As Worksheet
    Dim tgtPath As String
    Dim action As String, formatType As String

    tgtPath = txtSavePath.Text
    action = UCase(cboAction.Value)
    formatType = UCase(cboFormat.Value)

    ' Open existing workbook or create new
    On Error Resume Next
    Set targetWB = Workbooks.Open(tgtPath)
    If targetWB Is Nothing Then
        Set targetWB = Workbooks.Add
        Application.DisplayAlerts = False
        Do While targetWB.Sheets.Count > 0
            targetWB.Sheets(1).Delete
        Loop
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0

    ' Collect selected sheets
    Set selectedSheets = New Collection
    For i = 0 To lstSheets.ListCount - 1
        If lstSheets.Selected(i) Then selectedSheets.Add lstSheets.List(i)
    Next i

    If selectedSheets.Count = 0 Then
        MsgBox "No sheets selected.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Import sheets
    For Each sheetInfo In selectedSheets
        filePath = Split(sheetInfo, "|")(0)
        sheetName = Split(sheetInfo, "|")(1)

        On Error Resume Next
        Set wbSrc = Workbooks.Open(filePath, ReadOnly:=True)
        If Err.Number <> 0 Or wbSrc Is Nothing Then
            MsgBox "Cannot open: " & filePath, vbExclamation
            Err.Clear
            GoTo NextSheet
        End If
        On Error GoTo 0

        Set ws = wbSrc.Sheets(sheetName)
        If Not ws Is Nothing Then
            If action = "COPY" Then
                ws.Copy After:=targetWB.Sheets(targetWB.Sheets.Count)
            ElseIf action = "MOVE" Then
                ws.Move After:=targetWB.Sheets(targetWB.Sheets.Count)
            End If
        End If

        wbSrc.Close SaveChanges:=False
        Set ws = Nothing
        Set wbSrc = Nothing
NextSheet:
    Next sheetInfo

    ' Save workbook in selected format
    Select Case formatType
        Case "XLSX"
            targetWB.SaveAs fileName:=tgtPath, FileFormat:=xlOpenXMLWorkbook
        Case "CSV"
            Dim csvPath As String, wsLoop As Worksheet
            For Each wsLoop In targetWB.Sheets
                csvPath = Replace(tgtPath, ".csv", "_" & wsLoop.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".csv")
                wsLoop.Copy
                ActiveWorkbook.SaveAs fileName:=csvPath, FileFormat:=xlCSV
                ActiveWorkbook.Close False
            Next wsLoop
        Case "TXT"
            Dim txtPath As String
            For Each wsLoop In targetWB.Sheets
                txtPath = Replace(tgtPath, ".txt", "_" & wsLoop.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")
                wsLoop.Copy
                ActiveWorkbook.SaveAs fileName:=txtPath, FileFormat:=xlTextWindows
                ActiveWorkbook.Close False
            Next wsLoop
    End Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Sheets successfully added to: " & tgtPath, vbInformation

    targetWB.Close SaveChanges:=True
    Unload Me
End Sub

'---------------- Helpers ----------------
Private Sub RefreshFileList()
    Dim i As Long
    lstFiles.Clear
    For i = 1 To fileList.Count
        lstFiles.AddItem fileList(i)
    Next i
End Sub

Private Function ItemInCollection(col As Collection, key As Variant) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = key Then ItemInCollection = True: Exit Function
    Next i
    ItemInCollection = False
End Function

Private Sub LoadSheetList()
    Dim f As Variant, wbTemp As Workbook, ws As Worksheet
    lstSheets.Clear
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each f In fileList
        On Error Resume Next
        Set wbTemp = Workbooks.Open(fileName:=f, ReadOnly:=True)
        On Error GoTo 0
        If Not wbTemp Is Nothing Then
            For Each ws In wbTemp.Sheets
                lstSheets.AddItem f & "|" & ws.Name
            Next ws
            wbTemp.Close SaveChanges:=False
            Set wbTemp = Nothing
        End If
    Next f
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


