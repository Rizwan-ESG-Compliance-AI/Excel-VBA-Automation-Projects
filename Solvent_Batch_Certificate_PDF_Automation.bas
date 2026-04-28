Option Explicit

'============ SAFE CONSTANTS (no external refs needed) ============
Private Const CDP_TYPE_STRING As Long = 4        ' msoPropertyTypeString
Private Const FD_FOLDER_PICKER As Long = 4       ' msoFileDialogFolderPicker
Private Const SHAPE_ROUNDED_RECT As Long = 5     ' msoShapeRoundedRectangle

'=========================== PUBLIC MACROS =========================

' Save ActiveSheet as PDF into the per-workbook folder (asks checklist)
Public Sub SaveAsPDF_PerWorkbookFolder()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim reason As String
    Dim folderPath As String, fileName As String, fullPath As String
    Dim solvent As String, batchDateRaw As String, batchDate As String, batchNo As String
    Dim printRange As Range

    ' Checklist
    If Not ConfirmPreSave() Then
        MsgBox "Save cancelled. Please complete the checklist and try again.", vbInformation
        Exit Sub
    End If

    ' Validation
    reason = ValidateSheet(ws)
    If reason <> "" Then
        MsgBox "PDF not saved:" & vbCrLf & reason, vbExclamation
        Exit Sub
    End If

    ' Per-workbook folder (remembered)
    folderPath = GetWorkbookPdfFolder(wb)
    If Len(folderPath) = 0 Then
        MsgBox "No folder selected. Operation cancelled.", vbExclamation
        Exit Sub
    End If
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Filename from meta
    ExtractMeta ws, solvent, batchDateRaw, batchNo
    If IsDate(batchDateRaw) Then batchDate = Format(CDate(batchDateRaw), "yyyymmdd") Else batchDate = Format(Date, "yyyymmdd")
    If Len(solvent) = 0 Then solvent = "UnknownProduct"
    If Len(batchNo) = 0 Then batchNo = "BatchX"
    fileName = SanitizeFileName(solvent & "_" & batchDate & "_" & batchNo & ".pdf")
    fullPath = folderPath & fileName

    ' Page setup
    Set printRange = ws.UsedRange
    ws.PageSetup.PrintArea = printRange.Address
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        .LeftMargin = Application.InchesToPoints(0.9)
        .RightMargin = Application.InchesToPoints(0.9)
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .Orientation = IIf(printRange.Width > printRange.Height, xlLandscape, xlPortrait)
    End With

    On Error GoTo SaveErr
    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=fullPath, Quality:=xlQualityStandard
    On Error GoTo 0

    MsgBox "PDF saved successfully:" & vbCrLf & fullPath, vbInformation
    Exit Sub
SaveErr:
    MsgBox "Error saving PDF: " & Err.Description & vbCrLf & "Path: " & fullPath, vbCritical
End Sub

' One-click: ask checklist → pick/change folder → save PDF now
Public Sub ChangeFolderAndSavePDF()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim wb As Workbook: Set wb = ws.Parent
    Dim newFolder As String
    Dim reason As String, ok As Boolean

    ' Checklist
    If Not ConfirmPreSave() Then
        MsgBox "Operation cancelled. No changes made.", vbInformation
        Exit Sub
    End If

    ' Validation
    reason = ValidateSheet(ws)
    If reason <> "" Then
        MsgBox "PDF not saved:" & vbCrLf & reason, vbExclamation
        Exit Sub
    End If

    ' Pick folder
    newFolder = PickFolder("Select a folder to save PDFs for this workbook")
    If Len(newFolder) = 0 Then
        MsgBox "Operation cancelled. Folder unchanged.", vbInformation
        Exit Sub
    End If

    ' Ensure folder exists (create if needed)
    If Not FolderExists(newFolder) Then
        If MsgBox("Folder doesn't exist:" & vbCrLf & newFolder & vbCrLf & _
                  "Create it now?", vbQuestion + vbYesNo) = vbYes Then
            ok = CreateFolderIfMissing(newFolder)
            If Not ok Then
                MsgBox "Couldn't create folder. Nothing changed.", vbCritical
                Exit Sub
            End If
        Else
            MsgBox "Folder unchanged.", vbInformation
            Exit Sub
        End If
    End If

    ' Remember on this workbook
    WriteCdpString wb, "PdfFolder", newFolder

    ' Save now to that folder
    SaveCurrentSheetPdfToFolder ws, newFolder
End Sub

' Print ActiveSheet (same checklist + validation)
Public Sub PrintCurrentSheetWithChecklist(Optional ByVal choosePrinter As Boolean = False, _
                                          Optional ByVal copies As Long = 1)
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim printRange As Range
    Dim reason As String

    If Not ConfirmPreSave() Then
        MsgBox "Print cancelled. Please complete the checklist and try again.", vbInformation
        Exit Sub
    End If

    reason = ValidateSheet(ws)
    If reason <> "" Then
        MsgBox "Print cancelled:" & vbCrLf & reason, vbExclamation
        Exit Sub
    End If

    If choosePrinter Then
        On Error Resume Next
        Application.Dialogs(xlDialogPrinterSetup).Show
        If Err.Number <> 0 Then
            Err.Clear
            MsgBox "Printer selection cancelled. Nothing printed.", vbInformation
            Exit Sub
        End If
        On Error GoTo 0
    End If

    Set printRange = ws.UsedRange
    ws.PageSetup.PrintArea = printRange.Address
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        .LeftMargin = Application.InchesToPoints(0.9)
        .RightMargin = Application.InchesToPoints(0.9)
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .Orientation = IIf(printRange.Width > printRange.Height, xlLandscape, xlPortrait)
    End With

    If copies < 1 Then copies = 1
    On Error GoTo PrintErr
    ws.PrintOut Copies:=copies, Collate:=True
    On Error GoTo 0

    MsgBox "Sheet sent to printer" & IIf(copies > 1, " (" & copies & " copies)", "") & ".", vbInformation
    Exit Sub
PrintErr:
    MsgBox "Error while printing: " & Err.Description, vbCritical
End Sub

'========================= BUTTON CREATORS =========================

' Drop a "Save as PDF" button on the active sheet
Public Sub AddSavePDFButton()
    Dim shp As Shape
    Dim ws As Worksheet: Set ws = ActiveSheet

    On Error Resume Next
    ws.Shapes("btnSavePDF").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECT, 60, 40, 220, 40)
    With shp
        .Name = "btnSavePDF"
        .TextFrame.Characters.Text = "Save as PDF (with checklist)"
        .OnAction = "SaveAsPDF_PerWorkbookFolder"
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Line.Visible = False
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Color = vbWhite
            .Bold = True
            .Size = 12
        End With
    End With

    MsgBox "Button added: Save as PDF (with checklist).", vbInformation
End Sub

' Drop a "Print" button on the active sheet
Public Sub AddPrintButton()
    Dim shp As Shape
    Dim ws As Worksheet: Set ws = ActiveSheet

    On Error Resume Next
    ws.Shapes("btnPrintSheet").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECT, 60, 100, 220, 40)
    With shp
        .Name = "btnPrintSheet"
        .TextFrame.Characters.Text = "Print Sheet (with checklist)"
        .OnAction = "PrintCurrentSheetWithChecklist"
        .Fill.ForeColor.RGB = RGB(0, 158, 73)
        .Line.Visible = False
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Color = vbWhite
            .Bold = True
            .Size = 12
        End With
    End With

    MsgBox "Button added: Print Sheet (with checklist).", vbInformation
End Sub

' Drop a "Change Folder & Save PDF" button on the active sheet
Public Sub AddChangeFolderButton()
    Dim shp As Shape
    Dim ws As Worksheet: Set ws = ActiveSheet

    On Error Resume Next
    ws.Shapes("btnChangeFolder").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECT, 60, 160, 220, 40)
    With shp
        .Name = "btnChangeFolder"
        .TextFrame.Characters.Text = "Change Folder & Save PDF"
        .OnAction = "ChangeFolderAndSavePDF"
        .Fill.ForeColor.RGB = RGB(192, 112, 0)
        .Line.Visible = False
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        With .TextFrame.Characters.Font
            .Color = vbWhite
            .Bold = True
            .Size = 12
        End With
    End With

    MsgBox "Button added: Change Folder & Save PDF.", vbInformation
End Sub

'=========================== HELPER ROUTINES =======================

' Save ActiveSheet to a specific folder
Private Sub SaveCurrentSheetPdfToFolder(ByVal ws As Worksheet, ByVal targetFolder As String)
    Dim solvent As String, batchDateRaw As String, batchDate As String, batchNo As String
    Dim fileName As String, fullPath As String
    Dim printRange As Range

    ExtractMeta ws, solvent, batchDateRaw, batchNo
    If IsDate(batchDateRaw) Then batchDate = Format(CDate(batchDateRaw), "yyyymmdd") Else batchDate = Format(Date, "yyyymmdd")
    If Len(solvent) = 0 Then solvent = "UnknownProduct"
    If Len(batchNo) = 0 Then batchNo = "BatchX"

    fileName = SanitizeFileName(solvent & "_" & batchDate & "_" & batchNo & ".pdf")
    If Right$(targetFolder, 1) <> "\" Then targetFolder = targetFolder & "\"
    fullPath = targetFolder & fileName

    Set printRange = ws.UsedRange
    ws.PageSetup.PrintArea = printRange.Address
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = True
        .LeftMargin = Application.InchesToPoints(0.9)
        .RightMargin = Application.InchesToPoints(0.9)
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .Orientation = IIf(printRange.Width > printRange.Height, xlLandscape, xlPortrait)
    End With

    On Error GoTo SaveErr
    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=fullPath, Quality:=xlQualityStandard
    On Error GoTo 0

    MsgBox "PDF saved to:" & vbCrLf & fullPath, vbInformation
    Exit Sub
SaveErr:
    MsgBox "Error saving PDF: " & Err.Description & vbCrLf & "Path: " & fullPath, vbCritical
End Sub

' Per-workbook folder memory (uses CustomDocumentProperties)
Private Function GetWorkbookPdfFolder(ByVal wb As Workbook) As String
    Dim p As String
    p = ReadCdpString(wb, "PdfFolder")
    If FolderExists(p) Then
        GetWorkbookPdfFolder = p
        Exit Function
    End If
    p = PickFolder("Select folder to save PDFs for this workbook")
    If Len(p) > 0 Then
        WriteCdpString wb, "PdfFolder", p
        GetWorkbookPdfFolder = p
    Else
        GetWorkbookPdfFolder = ""
    End If
End Function

' --- Custom Document Properties helpers ---
Private Function ReadCdpString(ByVal wb As Workbook, ByVal propName As String) As String
    On Error Resume Next
    ReadCdpString = CStr(wb.CustomDocumentProperties(propName).Value)
    On Error GoTo 0
End Function

Private Sub WriteCdpString(ByVal wb As Workbook, ByVal propName As String, ByVal propValue As String)
    On Error Resume Next
    With wb.CustomDocumentProperties
        .Item(propName).Value = CStr(propValue)
        If Err.Number <> 0 Then
            Err.Clear
            .Add Name:=propName, LinkToContent:=False, Type:=CDP_TYPE_STRING, Value:=CStr(propValue)
        End If
    End With
    On Error GoTo 0
End Sub

' --- File system & dialogs ---
Private Function PickFolder(ByVal titleText As String) As String
    On Error Resume Next
    With Application.FileDialog(FD_FOLDER_PICKER)
        .Title = titleText
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1) Else PickFolder = ""
    End With
    On Error GoTo 0
End Function

Private Function FolderExists(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(p) > 0) And CreateObject("Scripting.FileSystemObject").FolderExists(p)
    On Error GoTo 0
End Function

Private Function CreateFolderIfMissing(ByVal p As String) As Boolean
    On Error GoTo Fail
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    CreateFolderIfMissing = True
    Exit Function
Fail:
    CreateFolderIfMissing = False
End Function

' --- Validation & parsing ---
Private Function ValidateSheet(ByVal ws As Worksheet) As String
    Dim c As Range
    Dim redFound As Boolean, greenEmptyFound As Boolean
    Dim firstRed As String, firstGreenEmpty As String
    For Each c In ws.UsedRange
        If CellIsRed(c) Then
            redFound = True
            If firstRed = "" Then firstRed = c.Address(False, False)
        End If
        If CellIsGreen(c) Then
            If IsCellEmptyValue(c) Then
                greenEmptyFound = True
                If firstGreenEmpty = "" Then firstGreenEmpty = c.Address(False, False)
            End If
        End If
    Next c
    If redFound Or greenEmptyFound Then
        Dim msg As String
        If redFound Then msg = msg & "- A red cell was found at " & firstRed & "." & vbCrLf
        If greenEmptyFound Then msg = msg & "- A green cell is empty at " & firstGreenEmpty & "." & vbCrLf
        ValidateSheet = msg & "Please fix the above before saving."
    Else
        ValidateSheet = ""
    End If
End Function

Private Function CellIsRed(ByVal c As Range) As Boolean
    On Error Resume Next
    Dim d As Long, disp As Long, idx As Variant
    d = c.Interior.Color
    disp = c.DisplayFormat.Interior.Color
    idx = c.Interior.ColorIndex
    On Error GoTo 0
    CellIsRed = (d = vbRed) Or (disp = vbRed) Or (idx = 3) Or (disp = RGB(255, 0, 0)) Or IsStrongRed(disp) Or IsStrongRed(d)
End Function

Private Function CellIsGreen(ByVal c As Range) As Boolean
    On Error Resume Next
    Dim d As Long, disp As Long, idx As Variant
    d = c.Interior.Color
    disp = c.DisplayFormat.Interior.Color
    idx = c.Interior.ColorIndex
    On Error GoTo 0
    CellIsGreen = (d = vbGreen) Or (disp = vbGreen) Or (idx = 4) Or (disp = RGB(0, 255, 0)) Or IsStrongGreen(disp) Or IsStrongGreen(d)
End Function

Private Function IsCellEmptyValue(ByVal c As Range) As Boolean
    Dim tgt As Range
    If c.MergeCells Then
        Set tgt = c.MergeArea.Cells(1, 1)
    Else
        Set tgt = c
    End If
    Dim v As Variant, v2 As Variant, t As String
    v = tgt.Value: v2 = tgt.Value2: t = tgt.Text
    If IsEmpty(v) Then IsCellEmptyValue = True: Exit Function
    If Len(Trim$(CStr(v))) = 0 Then IsCellEmptyValue = True: Exit Function
    If Len(Trim$(CStr(v2))) = 0 Then IsCellEmptyValue = True: Exit Function
    If Len(Trim$(t)) = 0 Then IsCellEmptyValue = True: Exit Function
    If CStr(v) = "'" Or CStr(v2) = "'" Then IsCellEmptyValue = True: Exit Function
    IsCellEmptyValue = False
End Function

Private Function IsStrongGreen(ByVal clr As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = clr Mod 256: g = (clr \ 256) Mod 256: b = (clr \ 65536) Mod 256
    IsStrongGreen = (g >= 200 And r <= 80 And b <= 80)
End Function

Private Function IsStrongRed(ByVal clr As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = clr Mod 256: g = (clr \ 256) Mod 256: b = (clr \ 65536) Mod 256
    IsStrongRed = (r >= 200 And g <= 80 And b <= 80)
End Function

Private Sub ExtractMeta(ByVal ws As Worksheet, _
                        ByRef solvent As String, _
                        ByRef batchDateRaw As String, _
                        ByRef batchNo As String)
    Dim cell As Range, j As Integer, label As String
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            label = LCase(Trim(CStr(cell.Value)))
            Select Case label
                Case "product:", "product"
                    For j = 1 To 3
                        If Trim(CStr(cell.Offset(0, j).Value)) <> "" Then
                            solvent = Trim(CStr(cell.Offset(0, j).Value))
                            Exit For
                        End If
                    Next j
                Case "cert date:", "cert. date:", "cert date", "date:", "date"
                    For j = 1 To 3
                        If Trim(CStr(cell.Offset(0, j).Value)) <> "" Then
                            batchDateRaw = Trim(CStr(cell.Offset(0, j).Value))
                            Exit For
                        End If
                    Next j
                Case "batch no:", "batch no.", "batch no"
                    For j = 1 To 3
                        If Trim(CStr(cell.Offset(0, j).Value)) <> "" Then
                            batchNo = Trim(CStr(cell.Offset(0, j).Value))
                            Exit For
                        End If
                    Next j
            End Select
        End If
    Next cell
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace(s, ch, "_")
    Next ch
    SanitizeFileName = s
End Function

' --- Unified 3-point checklist ---
Private Function ConfirmPreSave() As Boolean
    Dim msg As String, resp As VbMsgBoxResult
    msg = "Before proceeding, please confirm you have:" & vbCrLf & vbCrLf & _
          "• Checked the water content" & vbCrLf & _
          "• Updated the batch number" & vbCrLf & _
          "• Updated the certificate date" & vbCrLf & vbCrLf & _
          "Continue?"
    resp = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Pre-Save Checklist")
    ConfirmPreSave = (resp = vbYes)
End Function
