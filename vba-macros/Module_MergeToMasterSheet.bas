Attribute VB_Name = "Module_MergeToMasterSheet"
'==============================================================================
' Merge ALL Excel files into ONE sheet named "MasterData"
'------------------------------------------------------------------------------
' Usage:
'   1) Change the FOLDER_PATH constant below to your source folder.
'   2) Alt+F11 -> Insert this module into a new workbook (macro workbook).
'   3) Run MergeFilesToMasterSheet.
' Result:
'   - Each file's FIRST VISIBLE worksheet is copied completely (including its
'     header rows) and stacked below the previous block.
'   - One blank row is left between blocks.
'   - "MasterData" sheet is created if it does not exist.
'   - Temporary lock files (~$...) are ignored.
'   - Hidden / veryHidden sheets and "XLR_NoRangeSheet" are skipped.
'   - Right-To-Left display is preserved on MasterData.
'==============================================================================
Option Explicit

' >>> CHANGE THIS FOLDER PATH TO YOUR OWN FOLDER <<<
Public Const FOLDER_PATH As String = "C:\MyExcelFiles\"

Public Const SKIP_SHEET_NAME As String = "XLR_NoRangeSheet"
Public Const MASTER_SHEET_NAME As String = "MasterData"

Public Sub MergeFilesToMasterSheet()
    Dim fso As Object          ' Scripting.FileSystemObject - handles Persian (Unicode) paths
    Dim root As Object
    Dim fil As Object
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim srcRange As Range
    Dim nextRow As Long
    Dim count As Long
    Dim firstBlock As Boolean

    On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(FOLDER_PATH) Then
        MsgBox "Folder not found:" & vbCrLf & FOLDER_PATH, vbExclamation, "Merge To Master"
        Exit Sub
    End If

    ' Get or create the MasterData sheet
    On Error Resume Next
    Set dstWs = ThisWorkbook.Sheets(MASTER_SHEET_NAME)
    On Error GoTo ErrHandler
    If dstWs Is Nothing Then
        Set dstWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        dstWs.Name = MASTER_SHEET_NAME
    Else
        dstWs.Cells.Clear
    End If
    dstWs.DisplayRightToLeft = True   ' Preserve RTL display

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set root = fso.GetFolder(FOLDER_PATH)
    firstBlock = True
    nextRow = 1

    For Each fil In root.Files
        ' Skip temporary lock files and non-Excel files
        If Left$(fil.Name, 2) <> "~$" Then
            If LCase$(fso.GetExtensionName(fil.Name)) = "xlsx" Or _
               LCase$(fso.GetExtensionName(fil.Name)) = "xlsm" Or _
               LCase$(fso.GetExtensionName(fil.Name)) = "xls" Then

                Set srcWb = Workbooks.Open(Filename:=fil.Path, ReadOnly:=True, UpdateLinks:=0)

                Set srcWs = Nothing
                Dim ws As Worksheet
                For Each ws In srcWb.Worksheets
                    If ws.Visible = xlSheetVisible Then
                        If StrComp(Trim$(ws.Name), SKIP_SHEET_NAME, vbTextCompare) <> 0 Then
                            Set srcWs = ws
                            Exit For
                        End If
                    End If
                Next ws

                If Not srcWs Is Nothing Then
                    Set srcRange = srcWs.UsedRange

                    If Not firstBlock Then
                        nextRow = nextRow + 1   ' one blank row between blocks
                    End If

                    srcRange.Copy
                    dstWs.Cells(nextRow, 1).PasteSpecial Paste:=xlPasteAll
                    Application.CutCopyMode = False

                    nextRow = nextRow + srcRange.Rows.count
                    firstBlock = False
                    count = count + 1
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

    MsgBox count & " file(s) merged into '" & MASTER_SHEET_NAME & "' successfully.", _
           vbInformation, "Merge To Master"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Merge To Master"
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Resume CleanUp
End Sub
