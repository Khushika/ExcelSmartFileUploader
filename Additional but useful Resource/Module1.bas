Attribute VB_Name = "Module1"
Option Explicit

' Entry point to open the importer form
Public Sub OpenImporter()
    frmImporter.Show
End Sub

' Utility: Get file extension
Public Function GetFileExt(fpath As String) As String
    Dim p As Long
    p = InStrRev(fpath, ".")
    If p > 0 Then GetFileExt = LCase(Mid(fpath, p + 1)) Else GetFileExt = ""
End Function

' Ensure valid sheet name
Public Function SafeSheetName(s As String) As String
    Dim invalidChars As Variant, ch As Variant
    invalidChars = Array("\", "/", "*", "[", "]", ":", "?", "")
    For Each ch In invalidChars
        s = Replace(s, ch, "-")
    Next ch
    If Len(s) = 0 Then s = "Sheet"
    SafeSheetName = Left(s, 31)
End Function

' ---- MAIN IMPORT FUNCTION ----
Public Sub ImportSelectedSheets(sheetList As Collection, targetWB As Workbook, _
                                action As String, targetFormat As String, savePath As String)
    Dim item As Variant
    Dim wbSrc As Workbook, ws As Worksheet
    Dim srcFile As String, wsName As String
    Dim splitPos As Long
    Dim tempWB As Workbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each item In sheetList
        On Error GoTo ImportErr
        splitPos = InStr(item, "|")
        srcFile = Left(item, splitPos - 1)
        wsName = Mid(item, splitPos + 1)

        Set wbSrc = Workbooks.Open(fileName:=srcFile, ReadOnly:=False, AddToMru:=False)
        wbSrc.Sheets(wsName).Copy After:=targetWB.Sheets(targetWB.Sheets.Count)

        If UCase(action) = "MOVE" Then
            Application.DisplayAlerts = False
            wbSrc.Sheets(wsName).Delete
            Application.DisplayAlerts = True
        End If

        wbSrc.Close SaveChanges:=False
        On Error GoTo 0
    Next item

    ' Save if user selected convert option
    If Len(Trim(savePath)) > 0 Then
        Select Case UCase(targetFormat)
            Case "XLSX"
                ' Save if user selected convert option
If Len(Trim(savePath)) > 0 Then
    On Error Resume Next
    ' Delete existing file if already present to prevent conflict
    If Dir(savePath) <> "" Then Kill savePath
    On Error GoTo 0

    Select Case UCase(targetFormat)
        Case "XLSX"
            ' Ensure workbook is saved safely even if open
            targetWB.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

        Case "CSV", "TXT"
            ' Copy first sheet and save separately
            targetWB.Sheets(1).Copy
            Set tempWB = ActiveWorkbook

            If UCase(targetFormat) = "CSV" Then
                tempWB.SaveAs fileName:=savePath, FileFormat:=xlCSV, CreateBackup:=False
            Else
                tempWB.SaveAs fileName:=savePath, FileFormat:=xlText, CreateBackup:=False
            End If

            tempWB.Close SaveChanges:=False
    End Select
End If
            Case "CSV", "TXT"
                targetWB.Sheets(1).Copy
                Set tempWB = ActiveWorkbook
                If UCase(targetFormat) = "CSV" Then
                    tempWB.SaveAs fileName:=savePath, FileFormat:=xlCSV, CreateBackup:=False
                Else
                    tempWB.SaveAs fileName:=savePath, FileFormat:=xlText, CreateBackup:=False
                End If
                tempWB.Close SaveChanges:=False
        End Select
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Import completed successfully!", vbInformation
    Exit Sub

ImportErr:
    MsgBox "Error importing sheet '" & wsName & "' from file '" & srcFile & "': " & Err.Description, vbExclamation
    On Error Resume Next
    wbSrc.Close SaveChanges:=False
    On Error GoTo 0
NextItem:
    Resume Next
End Sub


