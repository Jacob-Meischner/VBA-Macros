Attribute VB_Name = "F_Auto_File_Create"
Sub UTI_Auto_Create_Files()
    Call OptimizeCode_Begin

Dim CurrentMonthRerunFile As Worksheet
Dim CurrentMonthName As String, PreviousMonth As String, Currentyear As String
Dim NewRerunFileName As String, NewRerunFileNameExt As String, PlainRerunFileName As String
Dim SourceFileName As String, DestinationFileName As String
Dim FSO As Object, OldRerunFileExt As String, OldRerunFileName As String
Dim OldRerunFileWorksheet As Worksheet, OldRerunFileWorkbook As Workbook, ALastRow As Long, PatientName As Range, DestinationOnNewWorkbook As Range, OldRerunName As String
Dim FindPatientRackInformation As Range
Dim SaveEachFileCurrentYear As String, SaveEachFileCurrentMonth As String
Dim CurrentMonthNoYear As String
Dim TestNameAbbrev As String
    TestNameAbbrev = "UTI"
Dim NewandPreviousMonthRerunSheetLocation As String
    NewandPreviousMonthRerunSheetLocation = "X:\Resulting\Open Array\UTI\"
Dim AnalyzedMacroFilesLocation As String
    AnalyzedMacroFilesLocation = "X:\Resulting\Open Array\UTI\Analyzed UTI Excel Files\"
Dim ArchivedRerunSheetsLocation As String
    ArchivedRerunSheetsLocation = "X:\Resulting\Open Array\UTI\Archived UTI Rerun Sheets\"
    
    
    Currentyear = Year(Date)
    CurrentMonthNoYear = MonthName(Month(Now))
    CurrentMonthName = MonthName(Month(Now)) & " " & Currentyear
    PreviousMonth = Format(DateAdd("m", -1, Date), "mmmm") & " " & Currentyear
    
    NewRerunFileName = NewandPreviousMonthRerunSheetLocation & CurrentMonthName & " - " & TestNameAbbrev & " Rerun Sheet"
    NewRerunFileNameExt = NewandPreviousMonthRerunSheetLocation & CurrentMonthName & " - " & TestNameAbbrev & " Rerun Sheet.xlsx"
    SaveEachFileCurrentYear = AnalyzedMacroFilesLocation & Currentyear
    SaveEachFileCurrentMonth = AnalyzedMacroFilesLocation & Currentyear & "\" & CurrentMonthNoYear
    
        If Dir(SaveEachFileCurrentYear, vbDirectory) = "" Then
            MkDir SaveEachFileCurrentYear                   'if directory "X:\Resulting\Analyzed RRM Excel Files\" & CurrentYear does not exist - make it
        Else

        End If
        
        If Dir(SaveEachFileCurrentMonth, vbDirectory) = "" Then
            MkDir SaveEachFileCurrentMonth                  'if directory "X:\Resulting\Analyzed RRM Excel Files\" & CurrentYear & "\" & CurrentMonthNoYear does not exist, then make it
        Else
        
        End If
            'create new rerun file for current month if one doesn't already exist
        If Dir(NewRerunFileNameExt) = "" Then                      'checking if rerun file for current month does not exist in file directory
            Workbooks.Add.SaveAs fileName:=NewRerunFileName        'if file does not exist - then create, name as current month & year & RRM Rerun Sheet, and save as new workbook
            MsgBox "A new rerun file has been created in folder: " & vbNewLine & NewRerunFileName, vbOKOnly
            
        Else
                                                                
        End If

            'Rerun File names & extensions - will print "Current Month & Current Year - *TEST ABBREV* Rerun Sheet
        OldRerunFileExt = NewandPreviousMonthRerunSheetLocation & PreviousMonth & " - " & TestNameAbbrev & " Rerun Sheet.xlsx"
        OldRerunFileName = NewandPreviousMonthRerunSheetLocation & PreviousMonth & " - " & TestNameAbbrev & " Rerun Sheet"
        PlainRerunFileName = CurrentMonthName & " - " & TestNameAbbrev & " Rerun Sheet"
        OldRerunName = PreviousMonth & " - " & TestNameAbbrev & " Rerun Sheet"

        If Dir(OldRerunFileExt) <> "" Then
           
           Workbooks.Open (OldRerunFileExt)
            Set OldRerunFileWorkbook = Workbooks(OldRerunName)
            Set OldRerunFileWorksheet = OldRerunFileWorkbook.Sheets(1)
            Set CurrentMonthRerunFile = Workbooks.Open(NewRerunFileNameExt).Sheets(1)

            With OldRerunFileWorksheet
                ALastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            End With
                
            Set DestinationOnNewWorkbook = CurrentMonthRerunFile.Cells(Rows.Count, "A").End(xlUp).Offset(1, 0)
                
                
            For Each PatientName In OldRerunFileWorksheet.Range("A1:A" & ALastRow).Cells
                If IsEmpty(PatientName.Value) = False And IsEmpty(PatientName.Offset(0, 2).Value) = True And PatientName.Interior.Color <> RGB(0, 0, 0) Then '<-------------FLUVID MACRO DOESN'T HAVE BLACK BACKGROUND COLOR WITH RACK NAMES
                    
                    With OldRerunFileWorksheet.Range(PatientName, PatientName.End(xlUp))
                        Application.FindFormat.Clear
                        Application.FindFormat.Interior.Color = RGB(0, 0, 0)
                        Set FindPatientRackInformation = .Find("*", SearchDirection:=xlPrevious, searchformat:=True)
                    End With
                    With PatientName.Offset(0, 1)
                        .Value = ("Automatically transferred to: " & PlainRerunFileName & ".")
                        .Font.Color = RGB(0, 0, 0)
                    End With
                    
                    DestinationOnNewWorkbook.Value = FindPatientRackInformation.Value   'rack text for incomplete patient to new wb
                    DestinationOnNewWorkbook.Interior.Color = FindPatientRackInformation.Interior.Color 'rack color to new wb
                    DestinationOnNewWorkbook.Font.Color = FindPatientRackInformation.Font.Color 'rack font color to new wb
                    Set DestinationOnNewWorkbook = DestinationOnNewWorkbook.Offset(1, 0)

                    With DestinationOnNewWorkbook
                        .Value = PatientName.Value
                        .Interior.Color = PatientName.Interior.Color
                        .Font.Color = PatientName.Font.Color
                    End With
                    With DestinationOnNewWorkbook.Offset(0, 1)
                        .Value = PatientName.Offset(0, 1).Value
                        .Interior.Color = PatientName.Offset(0, 1).Interior.Color
                        .Font.Color = PatientName.Font.Color
                    End With
                    Set DestinationOnNewWorkbook = DestinationOnNewWorkbook.Offset(1, 0)
                End If
            Next PatientName
            
            Application.FindFormat.Clear
            
            With CurrentMonthRerunFile.Range("A:A")
                .ColumnWidth = 25
                .Rows.AutoFit
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
            
            OldRerunFileWorksheet.Parent.Close True
            CurrentMonthRerunFile.Parent.Close True
            
            Dim CurrentYearDir As String
            CurrentYearDir = ArchivedRerunSheetsLocation & Currentyear & "\"
           
                If Dir(CurrentYearDir) = "" Then       'if folder named as current year does not exist in X:\Resulting\Archived RRM Rerun Sheets
                    MkDir CurrentYearDir
                    Else
                End If
            
            Set FSO = CreateObject("Scripting.Filesystemobject")        'scripting.filesystemobject is excels Object for moving/creating/copying files to other locations
            SourceFileName = OldRerunFileExt            'file path & file name & file extension to be moved
            DestinationFileName = CurrentYearDir & PreviousMonth & " - " & TestNameAbbrev & " Rerun Sheet.xlsx"
            
            FSO.movefile Source:=SourceFileName, Destination:=DestinationFileName
                MsgBox ("Last months rerun sheet has been moved to: " & DestinationFileName & "." & vbNewLine & vbNewLine & "All reruns without a 2nd rerun have been transferred to: " & PlainRerunFileName & ".")
        End If
        
        Call OptimizeCode_End
End Sub
