Attribute VB_Name = "C_Create_Format112_384"
Public isExit As Boolean
Public Property Get rrFilePath() As String
    rrFilePath = "X:\Resulting\Open Array\Wound\" & MonthName(Month(Now)) & " " & Year(Date) & " - Wound Rerun Sheet.xlsx" 'rrFilePath = "C:\Users\jacob\OneDrive\Desktop\UTI Iterim\UTI Rerun Sheet.xlsx"
End Property
Public Property Get rrFileName() As String
    rrFileName = MonthName(Month(Now)) & " " & Year(Date) & " - Wound Rerun Sheet.xlsx"
End Property

Sub Convert_To_384_AccuFill_Import()

Call OptimizeCode_Begin

Dim AccufillImport As Worksheet, Save384File As Variant, FileName384 As String
Dim HelperColumnPosition As Range, accessionNumber As Range
Dim HelperColumnDLastRow As Long, UTITargetRange As Range
Dim OpenArray1and2Array As Variant, OpenArray1and2Match, AccufillPosition1and2Match
Dim AccufillImportPositionLastRow As Long, AccufillPosition As Range, Patient384Location As Range
Dim FileName384csvPath As String, FileName384csvFolderLocation As String
Dim FirstSampleOfOpenArray As Range, fiName As Variant, PatientRange As Range
Dim WoundRerunSheet As Worksheet, WoundRerunSheetLastRow As Long, WoundRerunSheetSearchRange As Range, WoundRerunMatch As Range, PatientID As Range, FindRerunMatch As Variant
Dim Currentyear As String, CurrentMonthNoYear As String
Dim RNB6 As String, SplB5 As String

Set AccufillImport = ThisWorkbook.Sheets("Accufill Import 384-File")


Currentyear = Year(Date)
CurrentMonthNoYear = MonthName(Month(Now))
FileName384csvPath = "X:\Resulting\Open Array\Wound\"   'D:\384File    '384 file location
FileName384csvFolderLocation = "Resulting Macro\384 Files"
'-----------------------------------------------------------------
SplB5 = Format(Left(importInfoWS.Range("B5").Value, 10), "YYYYMMDD")
RNB6 = importInfoWS.Range("B6").Value

If IsEmpty(SplB5) = False Then
    FileName384 = (Format(Now, "YYYYMMDD")) & "_Wound_" & SplB5 & "_" & RNB6
Else
    FileName384 = (Format(Now, "YYYYMMDD")) & "_Wound_*RackDate*_*RackNumber*"
End If
'-----------------------------------------------------------------
For Each FirstSampleOfOpenArray In importInfoWS.Range("FirstSampleOfAllOpenArrays").Cells
     If IsEmpty(FirstSampleOfOpenArray) = False Then
          If IsEmpty(importInfoWS.Range("B12")) = False Then
               If (IsEmpty(importInfoWS.Range("B7")) = True Or IsEmpty(importInfoWS.Range("B8")) = True) Then
                    Application.ScreenUpdating = True
                    With importInfoWS.Range("B7:B8")
                         .Borders.Color = RGB(255, 0, 0)
                         .Borders.Weight = xlThick
                              MsgBox ("Enter missing information before continuing")
                         .Borders.Color = RGB(0, 0, 0)
                         .Borders.Weight = xlMedium
                         Exit Sub
                    End With
               End If
          ElseIf IsEmpty(importInfoWS.Range("B36")) = False Then
               If (IsEmpty(importInfoWS.Range("B9")) = True Or IsEmpty(importInfoWS.Range("B10")) = True) Then
                    Application.ScreenUpdating = True
                    With importInfoWS.Range("B9:B10")
                         .Borders.Color = RGB(255, 0, 0)
                         .Borders.Weight = xlThick
                              MsgBox ("Enter missing information before continuing")
                         .Borders.Color = RGB(0, 0, 0)
                         .Borders.Weight = xlMedium
                         Exit Sub
                    End With
               End If
          End If
    ElseIf (IsEmpty(FirstSampleOfOpenArray) = True And FirstSampleOfOpenArray.Address = "$B$12") Then
        With importInfoWS.Range("B7:B8")
            .Value = "N/A"
        End With
    ElseIf (IsEmpty(FirstSampleOfOpenArray) = True And FirstSampleOfOpenArray.Address = "$B$36") Then
        With importInfoWS.Range("B9:B10")
            .Value = "N/A"
        End With
    End If
Next FirstSampleOfOpenArray

'    Set WoundRerunSheet = Workbooks.Open(rrFilePath).Sheets("Sheet1")
'        With WoundRerunSheet
'            WoundRerunSheetLastRow = .Cells(.Rows.count, "A").End(xlUp).Row
'        End With
'
'    Set WoundRerunSheetSearchRange = WoundRerunSheet.Range("A1:A" & WoundRerunSheetLastRow).Cells
'    Set PatientRange = importInfoWS.Range("B12:B59")
'
'        PatientRange.Borders.Color = RGB(0, 0, 0)
'        PatientRange.Borders.Weight = xlThin
'
'        Set WoundRerunMatch = Nothing
'    For Each PatientID In PatientRange.Cells
'        FindRerunMatch = Application.match(PatientID, WoundRerunSheetSearchRange, 0)      'find patient id on rerun sheet
'        If Not IsError(FindRerunMatch) Then
'            If WoundRerunMatch Is Nothing Then
'                Set WoundRerunMatch = PatientID   'if it's found then add it to WoundRerunMatch if nothing is in WoundRerunMatch
'            Else
'                Set WoundRerunMatch = Application.Union(PatientID, WoundRerunMatch)
'            End If
'        End If
'    Next PatientID
'
'    If Not WoundRerunMatch Is Nothing Then
'        With WoundRerunMatch.Borders
'            .Color = RGB(230, 0, 0)       '<----DARK RED
'            .Weight = xlThick
'        End With
'    End If
'
'WoundRerunSheet.Parent.Close False

     With importInfoWS
          HelperColumnDLastRow = .Cells(.Rows.count, "D").End(xlUp).Row
     End With
     With AccufillImport
          AccufillImportPositionLastRow = .Cells(.Rows.count, "B").End(xlUp).Row
          With .Range("C2:C" & AccufillImportPositionLastRow)
               .Clear
          End With
     End With

Set AccufillPosition = AccufillImport.Range("B1:B" & AccufillImportPositionLastRow).Cells

OpenArray1and2Array = Array("A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", _
                            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", _
                            "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", _
                            "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", _
                            "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A23", "A24", _
                            "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21", "B22", "B23", "B24", _
                            "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C20", "C21", "C22", "C23", "C24", _
                            "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D20", "D21", "D22", "D23", "D24", _
                            "E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", _
                            "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                            "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G10", "G11", "G12", _
                            "H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9", "H10", "H11", "H12", _
                            "E13", "E14", "E15", "E16", "E17", "E18", "E19", "E20", "E21", "E22", "E23", "E24", _
                            "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", "F24", _
                            "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G20", "G21", "G22", "G23", "G24", _
                            "H13", "H14", "H15", "H16", "H17", "H18", "H19", "H20", "H21", "H22", "H23", "H24")

For Each HelperColumnPosition In importInfoWS.Range("D12:G" & HelperColumnDLastRow).Cells
     If HelperColumnPosition.Column = 4 Then
          Set accessionNumber = HelperColumnPosition.Offset(0, -2)
     ElseIf HelperColumnPosition.Column = 5 Then
          Set accessionNumber = HelperColumnPosition.Offset(0, -3)
     ElseIf HelperColumnPosition.Column = 6 Then
          Set accessionNumber = HelperColumnPosition.Offset(0, -4)
     ElseIf HelperColumnPosition.Column = 7 Then
          Set accessionNumber = HelperColumnPosition.Offset(0, -5)
     End If
    OpenArray1and2Match = Application.match(HelperColumnPosition.Value, OpenArray1and2Array, 0) 'searching macro helper columns for a match inside array
    
        If Not IsError(OpenArray1and2Match) Then
            AccufillPosition1and2Match = Application.match(HelperColumnPosition.Value, AccufillPosition, 0)
            If Not IsError(AccufillPosition1and2Match) Then
            Set Patient384Location = AccufillImport.Cells(AccufillPosition1and2Match, 3)  '<---needs changed to 3 after correct positioning is confirmed
                If Not IsEmpty(accessionNumber) Then
                    On Error Resume Next
                    Patient384Location.Value = Split(accessionNumber.Value, Chr(10))(1)
                    If IsEmpty(Patient384Location.Value) Then   'Patient384Location.Value = Split(AccessionNumber.Value, Chr(10))(1)
                        Patient384Location.Value = accessionNumber.Value
                    End If
                Else
                    GoTo NextIteration
                End If
            End If
        End If

NextIteration: Next HelperColumnPosition

     With importInfoWS.Range("B12:B59")
          .Borders.Weight = xlThin
     End With


'    'save accufill 384 file
'AccufillImport.Copy
''ChDrive "D"
'With ActiveWorkbook
'    Save384File = Application.GetSaveAsFilename(InitialFileName:=FileName384csvPath & FileName384 & "_384_File" & ".csv", FileFilter:="AccuFill-384File (*.csv),*.csv", Title:="Save As")
'    MsgBox Save384File
'    If Save384File = False Then
'        Exit Sub
'    Else
'        .SaveAs fileName:=Save384File, FileFormat:=xlCSV
'        .Close True
'        MsgBox ("File has been saved to " & FileName384csvPath)
'    End If
'End With
'    'save macro file
' ChDrive "X"
'    SaveResultingMacro = "X:\Resulting\Open Array\Wound\Analyzed Wound Excel Files\" & Currentyear & "\" & CurrentMonthNoYear & "\"
'    ChDir SaveResultingMacro
'fiName = Application.GetSaveAsFilename(InitialFileName:=SaveResultingMacro & FileName384 & ".xlsb", FileFilter:="Excel Macro-Enabled Workbook Binary (*.xlsb), *.xlsb", Title:="Save As")
'If fiName = False Then Exit Sub
'ActiveWorkbook.SaveAs fileName:=fiName

Call OptimizeCode_End

End Sub
