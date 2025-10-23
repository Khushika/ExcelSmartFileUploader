VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSheetSelector 
   Caption         =   "UserForm1"
   ClientHeight    =   4900
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5780
   OleObjectBlob   =   "frmSheetSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSheetSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Public fileList As Collection

Public Sub LoadSheets(files As Collection)
    Set fileList = files
    Dim f As Variant, wb As Workbook, ws As Worksheet
    
    lstSheets.Clear
    lstSheets.ColumnCount = 2
    lstSheets.ColumnWidths = "150 pt; 150 pt"
    
    Application.ScreenUpdating = False
    
    For Each f In fileList
        Set wb = Workbooks.Open(f, ReadOnly:=True)
        For Each ws In wb.Sheets
            lstSheets.AddItem Dir(f)
            lstSheets.List(lstSheets.ListCount - 1, 1) = ws.Name
        Next ws
        wb.Close SaveChanges:=False
    Next f
    
    Application.ScreenUpdating = True
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To lstSheets.ListCount - 1
        lstSheets.Selected(i) = True
    Next i
End Sub

Private Sub btnImport_Click()
    On Error Resume Next
    frmImporter.Hide
    Me.Hide
    
    Dim targetWB As Workbook
    Set targetWB = ActiveWorkbook
    ImportFilesFromList fileList, targetWB, frmImporter.cboAction.Value, _
                        IIf(frmImporter.chkConvert.Value, frmImporter.cboFormat.Value, ""), _
                        Trim(frmImporter.txtSavePath.Text)
    Me.Hide
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


