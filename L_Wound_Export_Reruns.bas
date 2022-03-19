Attribute VB_Name = "L_Wound_Export_Reruns"
Sub Wound_Export_Normal_Reruns()

    Call OptimizeCode_Begin

Dim RerunsToPull As Worksheet, FungalRRSheet As Worksheet, ExportCol As Variant, col As Variant
Dim RRSheetTargetLastRow As Long, RRSheetSearchRng As Range, RRSheetDest As Range, ThinBorders As Range
Dim RackNumber As String, RackDate As String, Rerun As Range, RRSheetALastRow As Long

ChDrive "X"

    Set RerunsToPull = ThisWorkbook.Sheets("Reruns To Pull")
    
    Set FungalRRSheet = Workbooks(rrFileName).Sheets(1)
        With FungalRRSheet
            RRSheetTargetLastRow = .Cells(.Rows.count, "B").End(xlUp).Row
            RRSheetALastRow = .Cells(.Rows.count, "A").End(xlUp).Row
        End With

    ExportCol = Array("A")
    
    Set RRSheetDest = FungalRRSheet.Range("B" & RRSheetTargetLastRow).Offset(1, -1) 'First value will goto rr sheet Column A - setting destination range off of column B last row
    
    For Each col In ExportCol
    
    RackNumber = RerunsToPull.Cells(8, col)
    RackDate = Left(RerunsToPull.Cells(2, col).Value, (InStr(RerunsToPull.Cells(2, col).Value, "  ")) - 1) 'Trims DATE  TIME value in A2
        
        If IsEmpty(RerunsToPull.Cells(9, col)) = False Then
            With RRSheetDest
                .Value = CStr(RackDate & " " & RackNumber)
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
            End With
        Else
            GoTo NextColumn
        End If
            Set RRSheetDest = RRSheetDest.Offset(1, 0)
            
        For Each Rerun In RerunsToPull.Range(RerunsToPull.Cells(9, col), RerunsToPull.Cells(Rows.count, col).End(xlUp)).Cells
            With RRSheetDest
                .Value = Rerun.Value
            End With
            With RRSheetDest.Offset(0, 1)
                .Value = Rerun.Offset(0, 2).Value
            End With
            Set RRSheetDest = RRSheetDest.Offset(1, 0)
       Next Rerun
    
NextColumn:    Next col
    
    With FungalRRSheet.Range("A:M")
        .ColumnWidth = 25
        .Font.Size = 12
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    Call OptimizeCode_End
End Sub

Sub Wound_Bordered_Reruns()

    Call OptimizeCode_Begin

Dim RR2Pull As Worksheet, UTIRerunWS As Worksheet, BorderedExportCol As Variant, BorderedRRCol As Variant, BorderedPatient As Range, BorderedTarget As Range
Dim UTIRerunWSPatientRngLastRow As Long, FungalRRSheetSearchRng As Range, RackDate2 As String, RackNumber2 As String, FinalTargetDestination As Range
Dim PatientNameMatch As Variant, TargetIDRNG As Range, CountNumberOfRRSheetDuplicates As Integer


Set RR2Pull = ThisWorkbook.Sheets("Reruns To Pull")
Set UTIRerunWS = Workbooks(rrFileName).Worksheets("Sheet1")

    With UTIRerunWS
        UTIRerunWSPatientRngLastRow = .Cells(Rows.count, "A").End(xlUp).Row
    End With

    Set FungalRRSheetSearchRng = UTIRerunWS.Range("A1:A" & UTIRerunWSPatientRngLastRow)
    
    BorderedExportCol = Array("A")
       
    For Each BorderedRRCol In BorderedExportCol
        
        RackNumber2 = RR2Pull.Cells(8, BorderedRRCol)
        RackDate2 = Left(RR2Pull.Cells(2, BorderedRRCol).Value, (InStr(RR2Pull.Cells(2, BorderedRRCol).Value, "  ")) - 1)
        
        For Each BorderedPatient In RR2Pull.Range(RR2Pull.Cells(9, BorderedRRCol), RR2Pull.Cells(Rows.count, BorderedRRCol).End(xlUp)).Cells
            If BorderedPatient.Borders.Color = RGB(230, 0, 0) Then
                Set BorderedTarget = BorderedPatient.Offset(0, 2)
                PatientNameMatch = Application.match(BorderedPatient.Value, FungalRRSheetSearchRng, 0) 'find patientID match in Column A - The ID's will always be placed in Column A
                If Not IsError(PatientNameMatch) Then
                    CountNumberOfRRSheetDuplicates = 0
                    CountNumberOfRRSheetDuplicates = Application.WorksheetFunction.CountIf(FungalRRSheetSearchRng, BorderedPatient)    'count the number of targets that particular patient has
                Set TargetIDRNG = UTIRerunWS.Range(UTIRerunWS.Cells(PatientNameMatch, 2), UTIRerunWS.Cells(((PatientNameMatch + CountNumberOfRRSheetDuplicates) - 1), 2)).Find(BorderedTarget.Value, LookIn:=xlValues) 'create a searchable range based on number of duplicates each patientID has
                    If Not (TargetIDRNG Is Nothing) Then
                        Set FinalTargetDestination = UTIRerunWS.Cells(TargetIDRNG.Row, Columns.count).End(xlToLeft).Offset(0, 1)
                        With FinalTargetDestination
                            .Value = BorderedTarget.Value
                            .Interior.Color = BorderedTarget.Interior.Color
                            .Borders.Weight = BorderedTarget.Borders.Weight
                            .Borders.Color = BorderedTarget.Borders.Color
                            If BorderedTarget.Interior.Color = RGB(255, 255, 0) Then
                                .Offset(0, 1).Value = "Inconclusive"
                            ElseIf BorderedTarget.Interior.Color = RGB(0, 255, 0) Then
                                .Offset(0, 1).Value = "Detected"
                            ElseIf BorderedTarget.Interior.Color = xlNone Then
                                .Offset(0, 1).Value = "Not Detected"
                            End If
                            .Offset(0, 2).Value = RackDate2 & " " & RackNumber2
                            .Offset(0, 2).Interior.Color = RGB(0, 0, 0)
                            .Offset(0, 2).Font.Color = RGB(255, 255, 255)
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
    
                With UTIRerunWS.Range("A:M")        'UTIRERUNWS IS SET TO WORKBOOKS.OPEN  ADD IN Open Workbook Module to Open/Close rerun sheet
                    .ColumnWidth = 23
                    .Font.Size = 12
                    .WrapText = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlVAlignCenter
                End With
    Call OptimizeCode_End

End Sub
