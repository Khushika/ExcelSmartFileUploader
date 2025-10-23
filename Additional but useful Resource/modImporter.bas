Attribute VB_Name = "modImporter"
Option Explicit

Public Sub ImportFilesFromList(fileList As Collection, targetWB As Workbook, action As String, tgtFormat As String, savePath As String)
    On Error GoTo errHandler
    
    Dim srcWB As Workbook, ws As Worksheet, sh As String
    Dim filePath As Variant
    Dim selectedSheets As Collection
    Set selectedSheets = New Collection
    
    ' Collect selected sheets from global SheetSelector form
    Dim i As Long
    For i = 1 To frmSheetSelector.lstSheets.ListCount
        If frmSheetSelector.lstSheets.Selected(i - 1) Then
            selectedSheets.Add frmSheetSelector.lstSheets.List(i - 1, 0) & "|" & frmSheetSelector.lstSheets.List(i - 1, 1)
        End If
    Next i
    
    If selectedSheets.Count = 0 Then
        MsgBox "No sheets selected for import.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each filePath In fileList
        Set srcWB = Workbooks.Open(filePath, ReadOnly:=True)
        
        For i = 1 To selectedSheets.Count
            Dim fileName As String, sheetName As String
            fileName = Split(selectedSheets(i), "|")(0)
            sheetName = Split(selectedSheets(i), "|")(1)
            
            If LCase(fileName) = LCase(Dir(filePath)) Then
                On Error Resume Next
                Set ws = srcWB.Sheets(sheetName)
                On Error GoTo 0
                
                If Not ws Is Nothing Then
                    ws.Copy After:=targetWB.Sheets(targetWB.Sheets.Count)
                    If action = "MOVE" Then ws.Delete
                End If
            End If
        Next i
        
        srcWB.Close SaveChanges:=False
    Next filePath

    If tgtFormat <> "" And savePath <> "" Then
        Select Case UCase(tgtFormat)
            Case "CSV": targetWB.SaveAs fileName:=savePath, FileFormat:=xlCSV
            Case "TXT": targetWB.SaveAs fileName:=savePath, FileFormat:=xlText
            Case Else:  targetWB.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
        End Select
        MsgBox "Import complete and saved as " & tgtFormat & "!", vbInformation
    Else
        MsgBox "Import complete!", vbInformation
    End If
    
cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
errHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume cleanup
End Sub

