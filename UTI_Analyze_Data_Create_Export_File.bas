Attribute VB_Name = "C_Compile_Raw_Data_Runs"
Sub Compile_Raw_Data()

Call OptimizeCode_Begin

Dim FileToOpen As Variant, ResultFile As Variant, xRet As Boolean, Name As String, SDInterpretation As Integer, CrtThresholdCutoff As Integer
Dim QSResultFileWB As Workbook, QSResultFileWS As Worksheet, FormattingWS As Worksheet, ImptPtInfo As Worksheet, PullReruns As Worksheet
Dim SampleName As Range, QSTarget As Range, sampleArrayIK As Variant, sampleArrayDE As Variant, i As Integer, FormattingWBCrtLastRow As Long, FlaggedSpecimensLastRow As Long
Dim CrtAverage As Range, FlaggedSpecimens As Range, FinalResult As Range, FirstTarget As Range, SecondTarget As Range, AccessionNumber As Range
Dim FirstTargetCrtValue As Range, SecondTargetCrtValue As Range, CrtSDValue As Range, FinalCrt As Range, twenty As Integer, CoAGStaph As Integer
Dim NameChange As Range, NameChangeMatch As Variant, ElastRow As Long, UlastRow As Long, PreConvertedNames As Range, SampleNameLastRow As Long
Dim DEColumnRng As Range, IKColumnRng As Range, DlastRow As Long, QSNameMatch As Variant, AccessionNumberPosition As Range
Dim Counter As Long, SampleNameStart As Range, TotalRows As Integer, QCProblems As Range, QCProblemsLastRow As Long
Dim NTCArr As Variant, PTC123Arr As Variant, PTC4Arr As Variant, PTC5Arr As Variant, PEC1Arr As Variant, NEC1Arr As Variant
Dim PostImportDLastRow As Long, findOASN As Range, ReverseLoop As String, EndPlateQSMatch As Variant, FindFullNameAndPosition As Range
Dim myPath As String, PositiveExtractionControl As String, NegativeExtractionControl As String, NegativeTemplateControl As String, PositiveTemplateControl1 As String
Dim PositiveTemplateControl2 As String, PositiveTemplateControl3 As String, PositiveTemplateControl4 As String, PositiveTemplateControl5 As String
Dim ColumnBReruns As Range, ColumnBRerunsLastRow As Long, ColumnCReruns As Range, ColumnCRerunsLastRow As Long, RerunPatient As Range, RedBorderSearchRng As Range

    ChDrive "X"
    myPath = "X:\Jacob\Macros\Open Array\UTM Open Array\UTI\Resulting Macro"      '"X:\Jacob\Macros\Open Array\UTM Open Array\UTI\UTI Validation Files"
    ChDir myPath
    
    PositiveExtractionControl = "PEC"
    NegativeExtractionControl = "NEC"
    NegativeTemplateControl = "NTC"
    PositiveTemplateControl1 = "PTC 1"
    PositiveTemplateControl2 = "PTC 2"
    PositiveTemplateControl3 = "PTC 3"
    PositiveTemplateControl4 = "PTC 4"
    PositiveTemplateControl5 = "PTC 5"
    
        PTC5Arr = Array("P. vulgaris", "A. baumannii", "K. pneumoniae", "P. aeruginosa", "K. oxytoca", "E. cloacae", "E. faecium", "M. morganii", "P. stuartii", _
                        "IMP", "C. freundii", "ureR", "S. agalactiae", "TEM/SHV/VEB", "OXA/GES/PER", "E. faecalis", "E. coli", "ESBL", "DHA", "ampC/FOX/ACC", _
                        "E. aerogenes", "Coagulase-neg. Staphylococcus", "C. albicans", "Xeno", "qnrA/qnrS", "OXA", "VIM/KPC", "Vancomycin")
                        
        PTC4Arr = Array("OXA/GES/PER", "ESBL", "Coagulase-neg. Staphylococcus", "Xeno", "OXA")
        
        PTC123Arr = Array("IMP", "TEM/SHV/VEB", "OXA/GES/PER", "ESBL", "DHA", "ampC/FOX/ACC", "Coagulase-neg. Staphylococcus", "Xeno", "qnrA/qnrS", "OXA", "VIM/KPC", "Vancomycin")
        
        PEC1Arr = Array("E. faecalis", "Xeno", "Vancomycin")
        
        NTCArr = Array("P. vulgaris", "A. baumannii", "K. pneumoniae", "P. aeruginosa", "K. oxytoca", "E. cloacae", "E. faecium", "M. morganii", "P. stuartii", _
                        "IMP", "C. freundii", "ureR", "S. agalactiae", "TEM/SHV/VEB", "OXA/GES/PER", "E. faecalis", "E. coli", "ESBL", "DHA", "ampC/FOX/ACC", _
                        "E. aerogenes", "Coagulase-neg. Staphylococcus", "C. albicans", "Xeno", "qnrA/qnrS", "OXA", "VIM/KPC", "Vancomycin")
        
        NEC1Arr = Array("Xeno")
        
CrtThresholdCutoff = 30
SDInterpretation = 2        'if both target crt values are numbers then the SD must be <= 2.00 in order to be called detected
twenty = 20
CoAGStaph = 27
                        Set FormattingWS = ThisWorkbook.Sheets("OpenArray Raw Data")
                        
 FileToOpen = Application.GetOpenFilename(FileFilter:="Excel Files (*.XLSX), *.XLSX", Title:="Select all files needing analyzed", MultiSelect:=True)       'if file types change to csv or something else, this needs changed
        If Not IsArray(FileToOpen) Then Exit Sub
    
        'Name = CStr(FileToOpen)
        'xRet = IsWorkBookOpenNow(Name)

        With FormattingWS
            .Range("D10").Value = "Sample Name"
            .Range("E10").Value = "Target Name"
            .Range("F10").Value = "Crt"
            .Range("H10").Value = "Crt SD"
            .Range("G10").Value = "Crt Average"
            .Range("M10").Value = "Final Result"
            .Range("N10").Value = "Final Crt"
        End With
                'select all result files at once
        For Each ResultFile In FileToOpen   '---------------------------------Import Result Files (Start)----------------------------
            
            Set QSResultFileWB = Workbooks.Open(ResultFile)
            Set QSResultFileWS = QSResultFileWB.Sheets("Results")
            
            TotalRows = 0
            Counter = 0
            
            With QSResultFileWS
                Set SampleName = .Range("A1:Q50").Find("Sample Name")
                SampleNameLastRow = .Cells(.Rows.Count, SampleName.Column).End(xlUp).Row
                Set SampleNameStart = .Range("D" & SampleName.Row).Offset(1, 0)
                
                With QSResultFileWS.Range("D" & SampleNameStart.Row, "D" & SampleNameLastRow)
                    For Each r In .Rows
                        If Application.CountA(r) <> 0 Then
                            Counter = Counter + 1
                        End If
                    Next r
                    TotalRows = Counter
                End With
                sampleArrayDE = .Range("D21:E" & SampleNameLastRow).Value   'sets arrayDE = sample name & target name column values
                sampleArrayIK = .Range("I21:K" & SampleNameLastRow).Value   'sets arrayIK = crt, crt average, sd column values
            End With
            
            With FormattingWS
                DlastRow = FormattingWS.Cells(Rows.Count, "D").End(xlUp).Row
                Set DEColumnRng = FormattingWS.Range("D" & DlastRow).Offset(1, 0)
                Set IKColumnRng = FormattingWS.Range("F" & DlastRow).Offset(1, 0)
                FormattingWS.Range(DEColumnRng, "E" & (DEColumnRng.Row + TotalRows) - 1).Value = sampleArrayDE '(DEColumnRng.Row - 1))).Value = sampleArrayDE '(SampleNameLastRow - 10) inserts sample name and target name column to macro **need to subtract (SampleNameLastRow - 10) because data starts populating to macro on row 11**
                FormattingWS.Range(IKColumnRng, "H" & (DEColumnRng.Row + TotalRows) - 1).Value = sampleArrayIK 'inserts CRT, crt Average, SD to macro **need to subtract (SampleNameLastRow - 10) because data starts populating to macro on row 11**
            End With
                Erase sampleArrayDE
                Erase sampleArrayIK
            QSResultFileWS.Parent.Close False
        Next ResultFile                     '---------------------------------Import Result Files (End)-------------------------------


        Set ImptPtInfo = ThisWorkbook.Worksheets("Import Patient Information")
        Set PullReruns = ThisWorkbook.Worksheets("Reruns To Pull")
        
        With PullReruns             '------------------------------Set Location on Reruns To Pull for Inconclusives + Reruns (Start)----------------------------------------
            .Range("A7:F3000").ClearContents
                ColumnBRerunsLastRow = PullReruns.Cells(Rows.Count, "C").End(xlUp).Row
                ColumnCRerunsLastRow = PullReruns.Cells(Rows.Count, "F").End(xlUp).Row
            Set ColumnBReruns = PullReruns.Range("C" & ColumnBRerunsLastRow).Offset(1, -2)
            Set ColumnCReruns = PullReruns.Range("F" & ColumnCRerunsLastRow).Offset(1, -2)
        End With                    '------------------------------Set Location on Reruns To Pull for Inconclusives + Reruns (End)----------------------------------------
            
        With FormattingWS
                FormattingWBCrtLastRow = .Cells(.Rows.Count, "F").End(xlUp).Row
                ElastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
                UlastRow = .Cells(.Rows.Count, "U").End(xlUp).Row
            Set PreConvertedNames = .Range("U1:U" & UlastRow)
                QCProblemsLastRow = .Cells(Rows.Count, "C").End(xlUp).Row       'change this to Columns A B and C
            Set QCProblems = .Range("C" & QCProblemsLastRow).Offset(1, 0)
                PostImportDLastRow = .Cells(Rows.Count, "D").End(xlUp).Row
             Set RedBorderSearchRng = .Range("D1:D" & PostImportDLastRow).Cells
        End With
           
        For Each NameChange In FormattingWS.Range("E11:E" & ElastRow).Cells     '-----------------------------Translate Target Names (Start)---------------------------------------
            NameChangeMatch = Application.Match(NameChange.Value, PreConvertedNames, 0) 'use helper columns in columns U and V on destination workbook to match/change names of everything in column e
                If Not IsError(NameChangeMatch) Then
                    With NameChange
                        .Value = FormattingWS.Cells(NameChangeMatch, 22)
                    End With
                End If
        Next NameChange                                                         '-----------------------------Translate Target Names (End)-----------------------------------------
        
        For Each CrtAverage In FormattingWS.Range("G11:G" & FormattingWBCrtLastRow).Cells       '----------------------Result Interpretation Conditions(Start)--------------------------------------
           
           Set FirstTarget = CrtAverage.Offset(0, -2)
           Set SecondTarget = CrtAverage.Offset(1, -2)
           Set FirstTargetCrtValue = CrtAverage.Offset(0, -1)
           Set SecondTargetCrtValue = CrtAverage.Offset(1, -1)
           Set FinalResult = CrtAverage.Offset(0, 6)
           Set CrtSDValue = CrtAverage.Offset(0, 1)
           Set FinalCrt = CrtAverage.Offset(0, 7)
           Set AccessionNumber = CrtAverage.Offset(0, -3)
        
            If FirstTarget.Value = SecondTarget.Value Then            'check 2 columns to the left, if this target and the target directly below are the same then
                    If FirstTarget.Value = "Coagulase-neg. Staphylococcus" Then
                        CrtThresholdCutoff = 27
                    Else
                        CrtThresholdCutoff = 30
                    End If
                If FirstTargetCrtValue.Value = "Undetermined" And SecondTargetCrtValue.Value = "Undetermined" Then
                    With FinalResult
                        .Value = "Not Detected"
                    End With
                    With FinalCrt
                        .Value = "0"
                    End With
                ElseIf IsNumeric(FirstTargetCrtValue.Value) = True And IsNumeric(SecondTargetCrtValue.Value) = True Then
                    If (FirstTargetCrtValue.Value <= CrtThresholdCutoff) And (SecondTargetCrtValue.Value <= CrtThresholdCutoff) Then
                        With FinalCrt
                            .Value = (CrtAverage.Value + CrtSDValue.Value)
                        End With
                        
                        If FinalCrt.Value <= CrtThresholdCutoff Then
                            With FinalResult
                                .Value = "Detected"
                                .Interior.Color = RGB(0, 255, 0)
                            End With
                        ElseIf FinalCrt.Value > CrtThresholdCutoff Then
                            With FinalResult
                                .Value = "Not Detected"
                            End With
                            With FinalCrt
                                .Value = "0"
                            End With
                        End If
                    ElseIf (FirstTargetCrtValue.Value <= CrtThresholdCutoff And SecondTargetCrtValue.Value > CrtThresholdCutoff) Or (FirstTargetCrtValue.Value > CrtThresholdCutoff And SecondTargetCrtValue.Value <= CrtThresholdCutoff) Then
                        With FinalCrt
                            .Value = (CrtAverage.Value - CrtSDValue.Value)
                        End With
                        
                        If (FinalCrt.Value <= CrtThresholdCutoff And CrtSDValue.Value <= SDInterpretation) Then
                            With FinalResult
                                .Value = "Detected"
                                .Interior.Color = RGB(0, 255, 0)
                            End With
                        ElseIf (FinalCrt.Value > CrtThresholdCutoff And CrtSDValue.Value <= SDInterpretation) Then
                            With FinalResult
                                .Value = "Not Detected"
                            End With
                            With FinalCrt
                                .Value = "0"
                            End With
                        ElseIf (FinalCrt.Value <= CrtThresholdCutoff And CrtSDValue.Value > SDInterpretation) Then
                            With FinalResult
                                .Value = "Not Detected"
                            End With
                            With FinalCrt
                                .Value = "0"
                            End With
                        End If
                    ElseIf (FirstTargetCrtValue.Value > CrtThresholdCutoff And SecondTargetCrtValue.Value > CrtThresholdCutoff) Then    '----
                        With FinalCrt '----
                            .Value = "0" '----
                        End With '----
                        With FinalResult '----
                            .Value = "Not Detected" '----
                        End With '----
                    End If
                ElseIf IsNumeric(FirstTargetCrtValue.Value) = True And IsNumeric(SecondTargetCrtValue.Value) = False Then
                    With CrtAverage
                        .Value = Application.Average(FirstTargetCrtValue.Value, 0) 'find average of firsttargetcrtvalue & 0 (Undetermined)
                    End With
                    With CrtAverage.Offset(1, 0)
                        .Value = vbNullString                                       'getting rid of the 2nd crtAverage value that's imported from the QS file.
                    End With
                    With CrtSDValue
                        .Value = Application.WorksheetFunction.StDev(FirstTargetCrtValue.Value, 0)  '0 being the "Undetermined"
                    End With
                    
                    If FirstTargetCrtValue.Value <= twenty And (AccessionNumber.Value <> PositiveExtractionControl) And (AccessionNumber.Value <> NegativeExtractionControl) _
                    And (AccessionNumber.Value <> NegativeTemplateControl) And (AccessionNumber.Value <> PositiveTemplateControl1) And (AccessionNumber.Value <> PositiveTemplateControl2) _
                    And (AccessionNumber.Value <> PositiveTemplateControl3) And (AccessionNumber.Value <> PositiveTemplateControl4) And (AccessionNumber.Value <> PositiveTemplateControl5) Then
                        With FinalCrt
                            .Value = "500"                                             'place 0 so specimen does not look like a positive in Ligo
                        End With
                        With FinalResult
                            .Value = "Inconclusive"
                            .Interior.Color = RGB(255, 255, 0)
                        End With
                        Set FindFullNameAndPosition = ImptPtInfo.Range("B10:C103").Find(AccessionNumber.Value)
                        If Not (FindFullNameAndPosition Is Nothing) Then
                            If (FindFullNameAndPosition.Borders.Weight = xlThin) Then   'THIS MAKES SURE ONLY SPECIMENS THAT ARE NOT RERUNS ARE PUT ON Reruns To Pull WS
                                If FindFullNameAndPosition.Column = 2 Then
                                    With ColumnBReruns
                                        .Value = FindFullNameAndPosition.Value
                                    End With
                                    With ColumnBReruns.Offset(0, 1)
                                        .Value = FindFullNameAndPosition.Offset(0, -1).Value
                                    End With
                                    With ColumnBReruns.Offset(0, 2)
                                        .Value = FirstTarget.Value
                                    End With
                                    Set ColumnBReruns = ColumnBReruns.Offset(1, 0)
                                ElseIf FindFullNameAndPosition.Column = 3 Then
                                    With ColumnCReruns
                                        .Value = FindFullNameAndPosition.Value
                                    End With
                                    With ColumnCReruns.Offset(0, 1)
                                        .Value = FindFullNameAndPosition.Offset(0, -2).Value
                                    End With
                                    With ColumnCReruns.Offset(0, 2)
                                        .Value = FirstTarget.Value
                                    End With
                                End If
                                    Set ColumnCReruns = ColumnCReruns.Offset(1, 0)
                            Else: GoTo NextIteration
                            End If
                        Else: MsgBox ("Could not find Accession Number " & AccessionNumber.Value & "." & vbNewLine & "It's possible the wrong result file from the QuantStudio was selected.")
                        End If
                    ElseIf FirstTargetCrtValue.Value > twenty Then
                            With FinalResult
                                .Value = "Not Detected"
                            End With
                            With FinalCrt
                                .Value = "0"
                            End With
                    End If
                ElseIf IsNumeric(FirstTargetCrtValue.Value) = False And IsNumeric(SecondTargetCrtValue.Value) = True Then
                    With CrtAverage
                        .Value = Application.Average(SecondTargetCrtValue.Value, 0)
                    End With
                    With CrtAverage.Offset(1, 0)
                        .Value = vbNullString
                    End With
                    With CrtSDValue
                        .Value = Application.WorksheetFunction.StDev(SecondTargetCrtValue.Value, 0)
                    End With
                    If SecondTargetCrtValue.Value <= twenty Then
                        With FinalCrt
                            .Value = "500"
                        End With
                        With FinalResult
                            .Value = "Inconclusive"
                            .Interior.Color = RGB(255, 255, 0)
                        End With
                        Set FindFullNameAndPosition = ImptPtInfo.Range("B10:C103").Find(AccessionNumber.Value)
                        If Not (FindFullNameAndPosition Is Nothing) Then
                            If (FindFullNameAndPosition.Borders.Weight = xlThin) Then   'THIS MAKES SURE ONLY SPECIMENS THAT ARE NOT RERUNS ARE PUT ON Reruns To Pull WS
                                If FindFullNameAndPosition.Column = 2 Then
                                    With ColumnBReruns
                                        .Value = FindFullNameAndPosition.Value
                                    End With
                                    With ColumnBReruns.Offset(0, 1)
                                        .Value = FindFullNameAndPosition.Offset(0, -1).Value
                                    End With
                                    With ColumnBReruns.Offset(0, 2)
                                        .Value = FirstTarget.Value
                                    End With
                                    Set ColumnBReruns = ColumnBReruns.Offset(1, 0)
                                ElseIf FindFullNameAndPosition.Column = 3 Then
                                    With ColumnCReruns
                                        .Value = FindFullNameAndPosition.Value
                                    End With
                                    With ColumnCReruns.Offset(0, 1)
                                        .Value = FindFullNameAndPosition.Offset(0, -2).Value
                                    End With
                                    With ColumnCReruns.Offset(0, 2)
                                        .Value = FirstTarget.Value
                                    End With
                                End If
                                    Set ColumnCReruns = ColumnCReruns.Offset(1, 0)
                            Else: GoTo NextIteration
                            End If
                        Else: MsgBox ("Could not find Accession Number " & AccessionNumber.Value & "." & vbNewLine & "It's possible the wrong result file from the QuantStudio was selected.")
                        End If
                    ElseIf SecondTargetCrtValue.Value > twenty Then
                            With FinalResult
                                .Value = "Not Detected"
                            End With
                            With FinalCrt
                                .Value = "0"
                            End With
                    End If
                End If                                                                          '----------------------Result Interpretation Conditions(End)--------------------------------------
                    If AccessionNumber.Value = PositiveExtractionControl Or AccessionNumber.Value = NegativeExtractionControl Then  'or accessionnumber.value = "NTC" or accessionnumber.value = "PTC1" or accessionnumber.value = "PTC2" or accessionnumber.value = "PTC3" or accessionnumber.value = "PTC4" or accessionnumber.value = "PTC5" then
                        If AccessionNumber.Value = PositiveExtractionControl Then
                            QSNameMatch = Application.Match(FirstTarget.Value, PEC1Arr, 0)
                        ElseIf AccessionNumber.Value = NegativeExtractionControl Then
                            QSNameMatch = Application.Match(FirstTarget.Value, NEC1Arr, 0)
                        End If
                            If Not IsError(QSNameMatch) Then
                                If FinalResult.Value <> "Detected" Then
                                    With QCProblems
                                        .Value = FirstTarget.Value
                                    End With
                                    With QCProblems.Offset(0, -1)
                                        .Value = AccessionNumber.Value
                                    End With
                                        For Each AccessionNumberPosition In FormattingWS.Range(FirstTarget.Offset(0, -1), "D" & PostImportDLastRow).Cells
                                            If AccessionNumberPosition.Value <> PositiveExtractionControl And AccessionNumberPosition.Value <> NegativeExtractionControl Then
                                                Set findOASN = ImptPtInfo.Range("B10:C103").Find(AccessionNumberPosition.Value)
                                                    If Not (findOASN Is Nothing) Then
                                                        If findOASN.Interior.Color = RGB(198, 224, 180) Then
                                                            With QCProblems.Offset(0, -2)
                                                                .Value = ImptPtInfo.Range("B7").Value
                                                            End With
                                                            Set QCProblems = QCProblems.Offset(1, 0)    'asdas
                                                            GoTo NextIteration
                                                        ElseIf findOASN.Interior.Color = RGB(255, 230, 153) Then
                                                            With QCProblems.Offset(0, -2)
                                                                .Value = ImptPtInfo.Range("B8").Value
                                                            End With
                                                            Set QCProblems = QCProblems.Offset(1, 0)    'asdasd
                                                            GoTo NextIteration
                                                        ElseIf findOASN.Interior.Color = RGB(174, 170, 170) Then
                                                            With QCProblems.Offset(0, -2)
                                                                .Value = ImptPtInfo.Range("C7").Value
                                                            End With
                                                            Set QCProblems = QCProblems.Offset(1, 0)    'asdasd
                                                            GoTo NextIteration
                                                        ElseIf findOASN.Interior.Color = RGB(180, 198, 231) Then
                                                            With QCProblems.Offset(0, -2)
                                                                .Value = ImptPtInfo.Range("C8").Value
                                                            End With
                                                            Set QCProblems = QCProblems.Offset(1, 0)    'asdasd
                                                            GoTo NextIteration
                                                        End If
                                                    Else: MsgBox ("Could not find OpenArray Serial Number for QC " & AccessionNumber.Value & ". " & vbNewLine & vbNewLine & _
                                                            "Writing down this row number: " & AccessionNumber.Row & " may help find which OpenArray this control belongs to.")
                                                            GoTo NextIteration
                                                    End If
                                            End If
                                        Next AccessionNumberPosition
                                Else: GoTo NextIteration
                                End If
                            End If
                    End If
                    If AccessionNumber.Value = NegativeTemplateControl Or AccessionNumber.Value = PositiveTemplateControl1 Or _
                    AccessionNumber.Value = PositiveTemplateControl2 Or AccessionNumber.Value = PositiveTemplateControl3 Or _
                    AccessionNumber.Value = PositiveTemplateControl4 Or AccessionNumber.Value = PositiveTemplateControl5 Then       '---------------------NTC-PTC1-PTC5 Conditions(Start)-------------------------
                        If AccessionNumber.Value = NegativeTemplateControl Then
                            EndPlateQSMatch = Application.Match(FirstTarget.Value, NTCArr, 0)
                        ElseIf (AccessionNumber.Value = PositiveTemplateControl1 Or AccessionNumber.Value = PositiveTemplateControl2 Or AccessionNumber.Value = PositiveTemplateControl3) Then
                            EndPlateQSMatch = Application.Match(FirstTarget.Value, PTC123Arr, 0)
                        ElseIf AccessionNumber.Value = PositiveTemplateControl4 Then
                            EndPlateQSMatch = Application.Match(FirstTarget.Value, PTC4Arr, 0)
                        ElseIf AccessionNumber.Value = PositiveTemplateControl5 Then
                            EndPlateQSMatch = Application.Match(FirstTarget.Value, PTC5Arr, 0)
                        End If
                            If Not IsError(EndPlateQSMatch) Then
                                If AccessionNumber.Value = NegativeTemplateControl Then
                                    If FinalResult.Value <> "Not Detected" Then     'What if inconclusive comes back?
                                        With QCProblems
                                            .Value = FirstTarget.Value
                                        End With
                                        With QCProblems.Offset(0, -1)
                                            .Value = AccessionNumber.Value
                                        End With
                                        GoTo FindSerialNumber
                                    Else:
                                        GoTo NextIteration
                                    End If
                                ElseIf AccessionNumber.Value <> NegativeTemplateControl Then              'PositiveTemplateControl1 Or AccessionNumber.Value = PositiveTemplateControl2 Or AccessionNumber.Value = PositiveTemplateControl3 Or AccessionNumber.Value = PositiveTemplateControl4 Or AccessionNumber.Value = PositiveTemplateControl5 Then
                                    If FinalResult.Value <> "Detected" Then
                                        With QCProblems
                                            .Value = FirstTarget.Value
                                        End With
                                        With QCProblems.Offset(0, -1)
                                            .Value = AccessionNumber.Value
                                        End With
FindSerialNumber:                               For i = AccessionNumber.Row To 11 Step -1
                                                ReverseLoop = FormattingWS.Cells(i, 4).Value
                                                If ReverseLoop <> PositiveTemplateControl1 And ReverseLoop <> PositiveTemplateControl2 And ReverseLoop <> PositiveTemplateControl3 And ReverseLoop <> PositiveTemplateControl4 And ReverseLoop <> PositiveTemplateControl5 And ReverseLoop <> NegativeTemplateControl Then
                                                    Set findOASN = ImptPtInfo.Range("B10:C103").Find(ReverseLoop)
                                                        If Not (findOASN Is Nothing) Then
                                                            If findOASN.Interior.Color = RGB(198, 224, 180) Then
                                                                With QCProblems.Offset(0, -2)
                                                                    .Value = ImptPtInfo.Range("B7").Value
                                                                End With
                                                                Set QCProblems = QCProblems.Offset(1, 0)    'asdas
                                                                GoTo NextIteration
                                                            ElseIf findOASN.Interior.Color = RGB(255, 230, 153) Then
                                                                With QCProblems.Offset(0, -2)
                                                                    .Value = ImptPtInfo.Range("B8").Value
                                                                End With
                                                                Set QCProblems = QCProblems.Offset(1, 0)    'asdasd
                                                                GoTo NextIteration
                                                            ElseIf findOASN.Interior.Color = RGB(174, 170, 170) Then
                                                                With QCProblems.Offset(0, -2)
                                                                    .Value = ImptPtInfo.Range("C7").Value
                                                                End With
                                                                Set QCProblems = QCProblems.Offset(1, 0)    'asdasd
                                                                GoTo NextIteration
                                                            ElseIf findOASN.Interior.Color = RGB(180, 198, 231) Then
                                                                With QCProblems.Offset(0, -2)
                                                                    .Value = ImptPtInfo.Range("C8").Value
                                                                End With
                                                                Set QCProblems = QCProblems.Offset(1, 0)    'asdasd
                                                                GoTo NextIteration
                                                            End If
                                                    Else: MsgBox ("Could not find OpenArray Serial Number for QC " & AccessionNumber.Value & ". " & vbNewLine & vbNewLine & _
                                                            "Writing down this row number: " & AccessionNumber.Row & " may help find which OpenArray this control belongs to.")
                                                            GoTo NextIteration
                                                        End If
                                                End If
                                            Next i
                                    Else
                                        GoTo NextIteration
                                    End If
                                End If
                            End If
                    End If                                                                      '---------------------NTC-PTC1-PTC5 Conditions(End)-------------------------
            Else: GoTo NextIteration                            'if the target directly below does not match the target above, then go to the next cell in Crt Average column
            End If
NextIteration:         Next CrtAverage
       
                                                                                                                
        Dim RedBorder As Range, SplitAccessionNumber As String, RedBorderMatch As Variant, ColumnLocation As Integer
                'RedBorder = Name + Accession Number + SpecID
        For Each RedBorder In ImptPtInfo.Range("B10:C103").Cells      '-------------------------Apply Border to Rerun Targets - IF ANY(Start)-------------------------
            If RedBorder.Borders.Color = RGB(230, 0, 0) Then        'If red borders are applied to ImptPTInfo then find how many targets are being rerun on rerun sheet - place name, position + targets on Reruns To Pull
                ColumnLocation = 0
                ColumnLocation = RedBorder.Column   'Column Location - will tell me where the final Name + Targets need to go on Reruns To Pull
                SplitAccessionNumber = Split(RedBorder.Value, Chr(10))(1)
                RedBorderMatch = Application.Match(SplitAccessionNumber, RedBorderSearchRng, 0) 'sample rng on OARawData
                
                If Not IsError(RedBorderMatch) Then 'RedBorderMatch = Row # of where the patient is found on FormattingWS - Starting Row for that patient, first target will always be P.Vulgaris +55 rows will always be Vancomycin
                        
                Dim UTIRRSheet As Worksheet, UTIRRSheetALastRow As Long, UTISearchRangeA As Range, FindNumberOfTargetsForPatient As Integer, PatientTarget As Range
                Dim PatientRowOnRRSheet As Variant, FindTargetonFormattingWS As Variant, PossibleTarget As Range, RerunFinalResult As Range

                Set UTIRRSheet = Workbooks.Open(rrFilePath).Sheets("Sheet1")
                    With UTIRRSheet
                        UTIRRSheetALastRow = .Cells(Rows.Count, "A").End(xlUp).Row
                        Set UTISearchRangeA = UTIRRSheet.Range("A1:A" & UTIRRSheetALastRow).Cells
                    End With
                    PatientRowOnRRSheet = Application.Match(RedBorder.Value, UTISearchRangeA, 0)  'returns the row # where patient rerun data starts on the rerun sheet
                    If Not IsError(PatientRowOnRRSheet) Then
                        FindNumberOfTargetsForPatient = 0
                        FindNumberOfTargetsForPatient = UTIRRSheet.Application.WorksheetFunction.CountIf(UTISearchRangeA, RedBorder.Value)  'returns # of total targets for that patient
                        For Each PatientTarget In UTIRRSheet.Range(UTIRRSheet.Cells(PatientRowOnRRSheet, 2), UTIRRSheet.Cells(((PatientRowOnRRSheet + FindNumberOfTargetsForPatient) - 1), 2)).Cells
                            For Each PossibleTarget In FormattingWS.Range(FormattingWS.Cells(RedBorderMatch, 5), FormattingWS.Cells((RedBorderMatch + 55), 5)).Cells  'for each target on formattingWS.range(Start of patient data:End of patient data) total of 56 rows per patient
                                    With PossibleTarget.Offset(0, -1)   'apply red border to each SampleID column because we want to ignore all targets besides the bordered ones
                                        .Borders.Color = RGB(230, 0, 0) 'we're only pulling the targets that have the red border in Column E with the interpreted result = fill color on Reruns To Pull
                                        .Borders.Weight = xlThick
                                    End With
                                Set RerunFinalResult = PossibleTarget.Offset(0, 8)
                                If PossibleTarget.Value = PossibleTarget.Offset(1, 0).Value Then    'only apply border to one of the duplicate targets
                                  If PossibleTarget.Value = PatientTarget.Value Then
                                  With PossibleTarget
                                      .Borders.Color = RGB(230, 0, 0)
                                      .Borders.Weight = xlThick
                                  End With
                                    If ColumnLocation = "2" Then
                                        With ColumnBReruns
                                            .Value = RedBorder.Value
                                            .Borders.Color = RedBorder.Borders.Color
                                            .Borders.Weight = RedBorder.Borders.Weight
                                        End With
                                        With ColumnBReruns.Offset(0, 1)
                                            .Value = RedBorder.Offset(0, -1).Value
                                        End With
                                        With ColumnBReruns.Offset(0, 2)
                                            .Value = PossibleTarget.Value
                                            .Borders.Color = PossibleTarget.Borders.Color
                                            .Borders.Weight = PossibleTarget.Borders.Weight
                                            .Interior.Color = RerunFinalResult.Interior.Color
                                        End With
                                        Set ColumnBReruns = ColumnBReruns.Offset(1, 0)
                                    ElseIf ColumnLocation = "3" Then
                                        With ColumnCReruns
                                            .Value = RedBorder.Value
                                            .Borders.Color = RedBorder.Borders.Color
                                            .Borders.Weight = RedBorder.Borders.Weight
                                        End With
                                        With ColumnCReruns.Offset(0, 1)
                                            .Value = RedBorder.Offset(0, -1).Value
                                        End With
                                        With ColumnBReruns.Offset(0, 2)
                                            .Value = PossibleTarget.Value
                                            .Borders.Color = PossibleTarget.Borders.Color
                                            .Borders.Weight = PossibleTarget.Borders.Weight
                                            .Interior.Color = RerunFinalResult.Interior.Color
                                        End With
                                        Set ColumnBReruns = ColumnBReruns.Offset(1, 0)
                                    End If
                                  Else
                                      GoTo NxtPossibleTarg
                                  End If
                                End If
NxtPossibleTarg:            Next PossibleTarget
                        Next PatientTarget
                    End If
                End If
            End If
        Next RedBorder                                                       '-------------------------Apply Border to Rerun Targets - IF ANY(End)-------------------------
        
        With FormattingWS.Range("D10:H10", "M10:N10")
            .HorizontalAlignment = xlHAlignCenter
            .Font.Size = 14
            .Font.Bold = True
        End With
        With FormattingWS.Range("A2:C50")
            .HorizontalAlignment = xlHAlignCenter
            .Columns.AutoFit
        End With
        With FormattingWS.Range("D10:N" & FormattingWBCrtLastRow)
            .HorizontalAlignment = xlHAlignCenter
            .Columns.AutoFit
            With FormattingWS.Range("F11:N" & FormattingWBCrtLastRow)
                .NumberFormat = "0.00"
            End With
        End With
            'If xRet <> True Then
            '    QSResultFileWB.Close False
            'End If
            'Else: Exit Sub
        'End If
        
Call OptimizeCode_End

End Sub

Sub Create_Ligo_Exports_File()

Call OptimizeCode_Begin

Dim ws As Worksheet, LigoExportsWS As Worksheet, ImpPtInfo As Worksheet, SaveDuplicateFile As String, UTIFSO As Object, DupSourceLoc As String, DestinationLoc As String
Dim c As Range, RangeValue As Range, i As Integer, RangeValue2 As Range
Dim LigoRanges As Range, LigoRanges2 As Range
Dim LigoExportsDest As Range, SameRowResult As Range, WSLastRow As Long, LigoExportsLastRow As Long
Dim LigoExportsDest2 As Range, AnalysisResultsPath As String, DuplicateFileLocation As String, CreateBackUp As Boolean, FileNameConstant As String
Dim SplitB5 As String, SplitC5 As String, RackNumberB6 As String, RackNumberC6 As String

AnalysisResultsPath = "C:\Users\jacob\OneDrive\Documents\Excel\UTM Open Array\"                         'Analysis Results file location -------------------------
DuplicateFileLocation = "C:\Users\jacob\OneDrive\Documents\Excel\UTM Open Array\Test Duplicate Folder\" 'Duplicate file for us to refer to -------------------------

Set ImpPtInfo = ThisWorkbook.Sheets("Import Patient Information")
Set ws = ThisWorkbook.Sheets("OpenArray Raw Data")
Set LigoExportsWS = ThisWorkbook.Sheets("Ligo Exports")
SplitB5 = Format(Left(ImpPtInfo.Range("B5").Value, 10), "YYYYMMDD") 'grabs/splits date from Import Patient Information.Range("B5")
SplitC5 = Format(Left(ImpPtInfo.Range("C5").Value, 10), "YYYYMMDD") 'grabs/splits date from Import Patient Information.Range("C5")
RackNumberB6 = ImpPtInfo.Range("B6").Value
RackNumberC6 = ImpPtInfo.Range("C6").Value

If SplitB5 = SplitC5 Then
    FileNameConstant = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_" & SplitB5 & "_RackID_" & RackNumberB6 & "," & RackNumberC6 & "_Results"
ElseIf SplitC5 = "" Then
    FileNameConstant = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_" & SplitB5 & "_RackID_" & RackNumberB6 & "_Results"
ElseIf SplitB5 <> SplitC5 And SplitC5 <> "" Then
    FileNameConstant = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_" & SplitB5 & "_" & SplitC5 & "_RackID_" & RackNumberB6 & "," & RackNumberC6 & "_Results"
Else
    FileNameConstant = (Format(Now, "YYYYMMDD")) & "_UTI_RackDate_YYYYMMDD_RackID_X,X_Results"
End If


LigoExportsWS.Range("A1:N7000").Clear

WSLastRow = ws.Cells(Rows.Count, "D").End(xlUp).Row
LigoExportsLastRow = LigoExportsWS.Cells(Rows.Count, "D").End(xlUp).Row
Set LigoExportsDest = LigoExportsWS.Cells(Rows.Count, "D").End(xlUp).Offset(1, 0)               'Range("A1:A" & LigoExportsLastRow).Cells
Set LigoExportsDest2 = LigoExportsWS.Cells(Rows.Count, "M").End(xlUp).Offset(1, 0)

With LigoExportsWS
    .Range("A1").Value = "Well"
    .Range("B1").Value = "UTI"
    .Range("D1").Value = "Sample"
    .Range("E1").Value = "Target"
    .Range("M1").Value = "Cq"
    .Range("N1").Value = ws.Range("N10").Value
End With

For i = 11 To WSLastRow Step 2

    Set LigoRanges = Application.Union(ws.Range("D" & i), ws.Range("E" & i))
    Set LigoRanges2 = Application.Union(ws.Range("M" & i), ws.Range("N" & i))
    
    For Each RangeValue2 In LigoRanges2
        With LigoExportsDest2
            .Value = RangeValue2.Value
            .NumberFormat = "0.00"
        End With
        
        If RangeValue2.Interior.Color = RGB(255, 255, 255) Then
        
        Else
            With LigoExportsDest2
                .Interior.Color = RangeValue2.Interior.Color
            End With
        End If
        
        Set LigoExportsDest2 = LigoExportsDest2.Offset(0, 1)
    Next RangeValue2
        Set LigoExportsDest2 = LigoExportsDest2.Offset(1, -2)
    
    For Each RangeValue In LigoRanges
        
        With LigoExportsDest
            .Value = RangeValue.Value
            .NumberFormat = "0.00"
        End With
        If RangeValue.Interior.Color = RGB(255, 255, 255) Then
            
            Else
            With LigoExportsDest
                .Interior.Color = RangeValue.Interior.Color
            End With
        End If
            Set LigoExportsDest = LigoExportsDest.Offset(0, 1)
            
    Next RangeValue
    Set LigoExportsDest = LigoExportsDest.Offset(1, -2)
Next i

With LigoExportsWS.Range("D:E", "M:N")
    .HorizontalAlignment = xlHAlignCenter
    .Columns.AutoFit
End With


'LigoExportsWS.Copy
'With ActiveWorkbook             'intially save in duplicate file location and copy file to Analysis Results
'    SaveDuplicateFile = Application.GetSaveAsFilename(InitialFileName:=DuplicateFileLocation & FileNameConstant, FileFilter:="Saving Duplicate LigoExports File(*.csv),*.csv")
'    If SaveDuplicateFile = "False" Then
'        Exit Sub
'    Else
'        .SaveAs fileName:=SaveDuplicateFile, FileFormat:=xlCSV
'        .Close False
'            MsgBox ("Duplicate file has been saved to " & DuplicateFileLocation & " folder")
'        Set UTIFSO = CreateObject("Scripting.Filesystemobject")
'            DupSourceLoc = SaveDuplicateFile
'            DestinationLoc = AnalysisResultsPath
            
'            UTIFSO.copyfile Source:=DupSourceLoc, Destination:=DestinationLoc
'            MsgBox ("Duplicate file as been copied and place in Analysis Results Folder")
'    End If
'End With

Call OptimizeCode_End

End Sub
