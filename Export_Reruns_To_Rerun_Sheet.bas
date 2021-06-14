Attribute VB_Name = "E_UTI_Export_Reruns"
Sub UTI_Export_Normal_Reruns()

    Call OptimizeCode_Begin

Dim RerunsToPull As Worksheet, UTIRRSheet As Worksheet, ExportCol As Variant, col As Variant
Dim RRSheetTargetLastRow As Long, RRSheetSearchRng As Range, RRSheetDest As Range, ThinBorders As Range
Dim RackNumber As String, RackDate As String, Rerun As Range, RRSheetALastRow As Long

    Set RerunsToPull = ThisWorkbook.Sheets("Reruns To Pull")
    
    Set UTIRRSheet = Workbooks.Open(rrFilePath).Sheets(1)
        With UTIRRSheet
            RRSheetTargetLastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
            RRSheetALastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With

    Set RRSheetSearchRng = UTIRRSheet.Range("B1:B" & RRSheetTargetLastRow).Cells

    ExportCol = Array("A", "D")
    
    Set RRSheetDest = UTIRRSheet.Range("B" & RRSheetTargetLastRow).Offset(1, -1) 'First value will goto rr sheet Column A - setting destination range off of column B last row
    
    For Each col In ExportCol
    
    RackNumber = RerunsToPull.Cells(6, col)
    RackDate = Left(RerunsToPull.Cells(2, col).Value, (InStr(RerunsToPull.Cells(2, col).Value, "  ")) - 1) 'Trims DATE  TIME value in A2 & D2
        
        If IsEmpty(RerunsToPull.Cells(7, col)) = False Then
            With RRSheetDest
                .Value = CStr(RackDate & " " & RackNumber)
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
            End With
        Else
            GoTo NextColumn
        End If
            Set RRSheetDest = RRSheetDest.Offset(1, 0)
            
        For Each Rerun In RerunsToPull.Range(RerunsToPull.Cells(7, col), RerunsToPull.Cells(Rows.Count, col).End(xlUp)).Cells
            With RRSheetDest
                .Value = Rerun.Value
            End With
            With RRSheetDest.Offset(0, 1)
                .Value = Rerun.Offset(0, 2).Value
            End With
            Set RRSheetDest = RRSheetDest.Offset(1, 0)
       Next Rerun
    
NextColumn:    Next col
    
    'With UTIRRSheet.Range("A:A", "D:D")     'NEED TO DO FORMATTING FOR RERUN SHEET
        
    'End With
    
    Call OptimizeCode_End
End Sub

Sub UTI_Bordered_Reruns()

    Call OptimizeCode_Begin

Dim RR2Pull As Worksheet, UTIRerunWS As Worksheet, BorderedExportCol As Variant, BorderedRRCol As Variant, BorderedPatient As Range, BorderedTarget As Range
Dim UTIRerunWSPatientRngLastRow As Long, UTIRRSheetSearchRng As Range, RackDate2 As String, RackNumber2 As String, FinalTargetDestination As Range
Dim PatientNameMatch As Variant, TargetIDRNG As Range, CountNumberOfRRSheetDuplicates As Integer


Set RR2Pull = ThisWorkbook.Sheets("Reruns To Pull")
Set UTIRerunWS = Workbooks.Open(rrFilePath).Worksheets("Sheet1")

    With UTIRerunWS
        UTIRerunWSPatientRngLastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    End With

    Set UTIRRSheetSearchRng = UTIRerunWS.Range("A1:A" & UTIRerunWSPatientRngLastRow)
    
    BorderedExportCol = Array("A", "D")
       
    For Each BorderedRRCol In BorderedExportCol
        
        RackNumber2 = RR2Pull.Cells(6, BorderedRRCol)
        RackDate2 = Left(RR2Pull.Cells(2, BorderedRRCol).Value, (InStr(RR2Pull.Cells(2, BorderedRRCol).Value, "  ")) - 1)
        
        For Each BorderedPatient In RR2Pull.Range(RR2Pull.Cells(7, BorderedRRCol), RR2Pull.Cells(Rows.Count, BorderedRRCol).End(xlUp)).Cells
            If BorderedPatient.Borders.Color = RGB(230, 0, 0) Then
                Set BorderedTarget = BorderedPatient.Offset(0, 2)
                PatientNameMatch = Application.Match(BorderedPatient.Value, UTIRRSheetSearchRng, 0) 'find patientID match in Column A - The ID's will always be placed in Column A
                If Not IsError(PatientNameMatch) Then
                    CountNumberOfRRSheetDuplicates = 0
                    CountNumberOfRRSheetDuplicates = Application.WorksheetFunction.CountIf(UTIRRSheetSearchRng, BorderedPatient)    'count the number of targets that particular patient has
                Set TargetIDRNG = UTIRerunWS.Range(UTIRerunWS.Cells(PatientNameMatch, 2), UTIRerunWS.Cells(((PatientNameMatch + CountNumberOfRRSheetDuplicates) - 1), 2)).Find(BorderedTarget.Value, LookIn:=xlValues) 'create a searchable range based on number of duplicates each patientID has
                    If Not (TargetIDRNG Is Nothing) Then
                        Set FinalTargetDestination = UTIRerunWS.Cells(TargetIDRNG.Row, Columns.Count).End(xlToLeft).Offset(0, 1)
                        With FinalTargetDestination
                            .Value = BorderedTarget.Value
                            .Interior.Color = BorderedTarget.Interior.Color
                            .Borders.Weight = BorderedTarget.Borders.Weight
                            .Borders.Color = BorderedTarget.Borders.Color
                            .Offset(0, 1).Value = RackDate2 & " " & RackNumber2
                            .Offset(0, 1).Interior.Color = RGB(0, 0, 0)
                            .Offset(0, 1).Font.Color = RGB(255, 255, 255)
                        End With
                        Set FinalTargetDestination = FinalTargetDestination.Offset(0, 1)
                    Else:
                        MsgBox "Could not find PatiendID for " & BorderedPatient.Value & "."
                        Exit Sub
                    End If
                End If
            End If
        Next BorderedPatient
    Next BorderedRRCol
    
                With UTIRerunWS.Range("A:E")        'UTIRERUNWS IS SET TO WORKBOOKS.OPEN  ADD IN Open Workbook Module to Open/Close rerun sheet
                    .Columns.Width = 25
                    .Font.Size = 12
                    .WrapText = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlVAlignCenter
                End With
    Call OptimizeCode_End

End Sub
