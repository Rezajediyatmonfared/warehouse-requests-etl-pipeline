Attribute VB_Name = "Module_MergeToSheets"
'==============================================================================
' Merge ~500 Excel files into ONE workbook - one sheet per file
'------------------------------------------------------------------------------
' Usage:
'   1) Change the FOLDER_PATH constant below to your source folder.
'   2) Alt+F11 -> Insert this module into a new workbook (macro workbook).
'   3) Run MergeFilesToSeparateSheets.
' Result:
'   - Each file's visible worksheet is copied into this workbook as a new sheet,
'     content unchanged (values, formats, merged cells, RTL, column widths...).
'   - Sheets named "XLR_NoRangeSheet" are skipped.
'   - Hidden / veryHidden sheets are skipped (only the first visible sheet of
'     each file is used).
'   - Temporary lock files (~$...) are ignored.
'   - Sheet names are made unique automatically.
'==============================================================================
Option Explicit

' >>> CHANGE THIS FOLDER PATH TO YOUR OWN FOLDER <<<
Public Const FOLDER_PATH As String = "C:\MyExcelFiles\"

Public Const SKIP_SHEET_NAME As String = "XLR_NoRangeSheet"

Public Sub MergeFilesToSeparateSheets()
    Dim fso As Object          ' Scripting.FileSystemObject - handles Persian (Unicode) paths
    Dim root As Object         ' Source folder
    Dim fil As Object          ' Current file
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim newName As String
    Dim count As Long

    On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(FOLDER_PATH) Then
        MsgBox "Folder not found:" & vbCrLf & FOLDER_PATH, vbExclamation, "Merge Files"
        Exit Sub
    End If

    Set root = fso.GetFolder(FOLDER_PATH)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each fil In root.Files
        ' Skip temporary Excel lock files and non-Excel files
        If Left$(fil.Name, 2) <> "~$" Then
            If LCase$(fso.GetExtensionName(fil.Name)) = "xlsx" Or _
               LCase$(fso.GetExtensionName(fil.Name)) = "xlsm" Or _
               LCase$(fso.GetExtensionName(fil.Name)) = "xls" Then

                Set srcWb = Workbooks.Open(Filename:=fil.Path, ReadOnly:=True, UpdateLinks:=0)

                ' Use only the FIRST VISIBLE sheet of each file
                Set srcWs = FirstVisibleSheet(srcWb)

                If Not srcWs Is Nothing Then
                    If StrComp(Trim$(srcWs.Name), SKIP_SHEET_NAME, vbTextCompare) <> 0 Then
                        newName = UniqueSheetName(fil.Name & "_" & srcWs.Name)
                        srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                        Set dstWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                        dstWs.Name = newName
                        count = count + 1
                    End If
                End If

                srcWb.Close SaveChanges:=False
                DoEvents
            End If
        End If
    Next fil

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox count & " sheet(s) merged successfully.", vbInformation, "Merge Files"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Merge Files"
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Resume CleanUp
End Sub

' Returns the first visible (not hidden / veryHidden) worksheet, or Nothing.
Private Function FirstVisibleSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If StrComp(Trim$(ws.Name), SKIP_SHEET_NAME, vbTextCompare) <> 0 Then
                Set FirstVisibleSheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

' Builds a unique sheet name (Excel limit: 31 chars, no : \ / ? * [ ])
Private Function UniqueSheetName(ByVal baseName As String) As String
    Dim nm As String
    Dim badChars As Variant
    Dim i As Long
    Dim suffix As Long

    ' Remove illegal characters
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(badChars) To UBound(badChars)
        baseName = Replace(baseName, badChars(i), "_")
    Next i
    ' Trim extension if present
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    If Len(baseName) = 0 Then baseName = "Sheet"

    nm = Left$(baseName, 31)
    suffix = 1
    Do While SheetNameExists(nm)
        suffix = suffix + 1
        nm = Left$(baseName, 31 - Len("_" & suffix)) & "_" & suffix
    Loop

    UniqueSheetName = nm
End Function

Private Function SheetNameExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetNameExists = Not ws Is Nothing
End Function
