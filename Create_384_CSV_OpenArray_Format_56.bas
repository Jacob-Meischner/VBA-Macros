Attribute VB_Name = "B_Create_384_CSV"
'Public isExit = boolean
Public Property Get rrFilePath() As String
    rrFilePath = "X:\Jacob\Macros\Open Array\UTM Open Array\UTI\Resulting Macro\UTI Rerun Sheet.xlsx"  '"X:\Resulting\Open Array\UTI\" & MonthName(Month(Now)) & " " & Year(Date) & " - UTI Rerun Sheet.xlsx"
End Property
Public Property Get rrFileName() As String
    rrFileName = "UTI Rerun Sheet.xlsx" 'MonthName(Month(Now)) & " " & Year(Date) & " - UTI Rerun Sheet.xlsx"
End Property

Sub Convert_To_384_AccuFill_Import()

Call OptimizeCode_Begin

Dim ImportPatientInformation As Worksheet, AccufillImport As Worksheet, Save384File As Variant, FileName384 As String
Dim HelperColumnPosition As Range
Dim HelperColumnDLastRow As Long, UTITargetRange As Range
Dim OpenArray1and2Array As Variant, OpenArray1and2Match, AccufillPosition1and2Match
Dim AccufillImportPositionLastRow As Long, AccufillPosition As Range, Patient384Location As Range
Dim FileName384csvPath As String, FileName384csvFolderLocation As String
Dim AccessionNumber As Range, RNB6 As String, RNC6 As String, SplB5 As String, SplC5 As String
Dim DateRange As Range, fiName As Variant, PatientRange As Range
Dim UTIRerunSheet As Worksheet, UTIRerunSheetLastRow As Long, UTIRerunSheetSearchRange As Range, UTIRerunMatch As Range, PatientID As Range, FindRerunMatch As Variant

Dim UTIRerunSheetTarget As String, InitialTargetRerun As Range, TargetListLastRow As Long

Set ImportPatientInformation = ThisWorkbook.Sheets("Import Patient Information")
Set AccufillImport = ThisWorkbook.Sheets("Accufill Import 384-File")

FileName384csvPath = "C:\Users\jacob\OneDrive\Documents\Excel\UTM Open Array\384-Example\384-Files\"   'D:\384File    '384 file location
FileName384csvFolderLocation = "Resulting Macro\384 Files"

SplB5 = Format(Left(ImportPatientInformation.Range("B5").Value, 10), "YYYYMMDD") 'grabs/splits date from Import Patient Information.Range("B5")
SplC5 = Format(Left(ImportPatientInformation.Range("C5").Value, 10), "YYYYMMDD") 'grabs/splits date from Import Patient Information.Range("C5")
RNB6 = ImportPatientInformation.Range("B6").Value
RNC6 = ImportPatientInformation.Range("C6").Value

If SplB5 = SplC5 Then
    FileName384 = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_" & SplB5 & "_RackID_" & RNB6 & "," & RNC6 & "_384-File"
ElseIf SplC5 = "" Then
    FileName384 = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_" & SplB5 & "_RackID_" & RNB6 & "_384-File"
ElseIf SplB5 <> SplC5 And SplC5 <> "" Then
    FileName384 = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_" & SplB5 & "_" & SplC5 & "_RackID_" & RNB6 & "," & RNC6 & "_384-File"
Else
    FileName384 = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_YYYYMMDD_RackID_X,X_384-File"
End If


For Each DateRange In ImportPatientInformation.Range("B5:C5").Cells
    If IsEmpty(DateRange.Value) Then
        With ImportPatientInformation.Range(DateRange, DateRange.Offset(4, 0))
            .Value = "N/A"
        End With
    End If
    If Not IsEmpty(DateRange.Value) Then
        If IsEmpty(DateRange.Offset(1, 0).Value) Or IsEmpty(DateRange.Offset(2, 0)) Then
            Application.ScreenUpdating = True
                With ImportPatientInformation.Range(DateRange.Offset(1, 0), DateRange.Offset(2, 0))
                    .Borders.Color = RGB(255, 0, 0)
                    .Borders.Weight = xlThick
                        MsgBox ("Enter missing information before continuing")
                    .Borders.Color = RGB(0, 0, 0)
                    .Borders.Weight = xlMedium
                    Exit Sub
                End With
        End If
    End If
Next DateRange


        Set UTIRerunSheet = Workbooks.Open(rrFilePath).Sheets("Sheet1")
            With UTIRerunSheet
                UTIRerunSheetLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            End With
            With ImportPatientInformation
                TargetListLastRow = .Cells(.Rows.Count, "N").End(xlUp).Row
            End With
    
        Set UTIRerunSheetSearchRange = UTIRerunSheet.Range("A1:A" & UTIRerunSheetLastRow).Cells
        Set InitialTargetRerun = ImportPatientInformation.Range("N" & TargetListLastRow).Offset(1, 0)
        Set PatientRange = ImportPatientInformation.Range("B10:C103")
            
            PatientRange.Borders.Color = RGB(0, 0, 0)
            PatientRange.Borders.Weight = xlThin
    
            Set UTIRerunMatch = Nothing
        For Each PatientID In PatientRange.Cells
            FindRerunMatch = Application.Match(PatientID, UTIRerunSheetSearchRange, 0)      'find patient id on rerun sheet
            If Not IsError(FindRerunMatch) Then
                If UTIRerunMatch Is Nothing Then
                    Set UTIRerunMatch = PatientID   'if it's found then add it to UTIRerunMatch if nothing is in UTIRerunMatch
                Else
                    Set UTIRerunMatch = Application.Union(PatientID, UTIRerunMatch)
                End If
            End If
        Next PatientID

        If Not UTIRerunMatch Is Nothing Then
            With UTIRerunMatch.Borders
                .Color = RGB(230, 0, 0)       '<----DARK RED
                .Weight = xlThick
            End With
        End If

UTIRerunSheet.Parent.Close False

With ImportPatientInformation
    HelperColumnDLastRow = .Cells(.Rows.Count, "D").End(xlUp).Row
End With
With AccufillImport
    AccufillImportPositionLastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
    .Range("SampleInfoPositions").ClearContents
End With

Set AccufillPosition = AccufillImport.Range("B1:B" & AccufillImportPositionLastRow).Cells

OpenArray1and2Array = Array("A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", _
                            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", _
                            "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", _
                            "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10", "D11", "D12", _
                            "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", "A21", "A22", "A23", "A24", _
                            "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21", "B22", "B23", "B24", _
                            "C13", "C14", "C15", "C16", "C17", "C18", "C19", "C20", "C21", "C22", "C23", "C24", _
                            "D13", "D14", "D15", "D16", "D17", "D18", "D19", "D20", "D21", "D22", "D23", "D24", _
                            "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10", "E11", "E12", _
                            "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                            "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G10", "G11", "G12", _
                            "H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9", "H10", "H11", "H12", _
                            "E13", "E14", "E15", "E16", "E17", "E18", "E19", "E20", "E21", "E22", "E23", "E24", _
                            "F13", "F14", "F15", "F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", "F24", _
                            "G13", "G14", "G15", "G16", "G17", "G18", "G19", "G20", "G21", "G22", "G23", "G24", _
                            "H13", "H14", "H15", "H16", "H17", "H18", "H19", "H20", "H21", "H22", "H23", "H24")

For Each HelperColumnPosition In ImportPatientInformation.Range("D10:E" & HelperColumnDLastRow).Cells
        Set AccessionNumber = HelperColumnPosition.Offset(0, -2)
        
    OpenArray1and2Match = Application.Match(HelperColumnPosition.Value, OpenArray1and2Array, 0) 'searching UTI macro helper columns for a match inside array
    
        If Not IsError(OpenArray1and2Match) Then
            AccufillPosition1and2Match = Application.Match(HelperColumnPosition.Value, AccufillPosition, 0)
            If Not IsError(AccufillPosition1and2Match) Then
            Set Patient384Location = AccufillImport.Cells(AccufillPosition1and2Match, 3)  '<---needs changed to 3 after correct positioning is confirmed
                If Not IsEmpty(AccessionNumber) Then
                    Patient384Location.Value = AccessionNumber.Value            'Patient384Location.Value = Split(AccessionNumber.Value, Chr(10))(1)
                Else
                    GoTo NextIteration
                End If
            End If
        End If

NextIteration: Next HelperColumnPosition

            'SAVE 384 FILE TO QS DESTOP 384well FOLDER
'AccufillImport.Copy
'With ActiveWorkbook
'    Save384File = Application.GetSaveAsFilename(InitialFileName:=FileName384csvPath & FileName384, Filefilter:="AccuFill-384File (*.csv),*.csv")
'    If Save384File = False Then
'        Exit Sub
'    Else
'        .SaveAs fileName:=Save384File, FileFormat:=xlCSV
'        .Close False
'        MsgBox ("File has been saved to " & FileName384csvPath & ".")
'    End If
'End With

                                                         '"X:\Resulting\Open Array\UTI\Analyzed UTI Excel Files\" & Currentyear & "\" & CurrentMonthNoYear & "\" & "UTI Macro - "
'SAVE ENTIRE MACRO FILE TO PATH TO THE RIGHT ----->

'fiName = Application.GetSaveAsFilename(InitialFileName:="C:\Users\jacob\OneDrive\Desktop\UTI Iterim\" & "UTI Macro - RackID:", FileFilter:="Excel Macro-Enabled Workbook Binary (*.XLSB), *.XLSB", Title:="Save As")
'If fiName = False Then Exit Sub
'ActiveWorkbook.SaveAs fileName:=fiName

Call OptimizeCode_End

End Sub
