Attribute VB_Name = "E_Import_Result_Files"
Sub Import_Result_Files()
          'DATE ADDED 3-5-22
'              Call OptimizeCode_Begin
Dim FileToOpen As Variant
Dim ClearStartRangeLastRowFung As Long, ClearStartRangeFung As Range
Dim ResultFile As Variant, QSResultFileWB As Workbook, QSResultFileWS As Worksheet, TotalRows As Long, Counter As Long
Dim SampleNameLastRow As Long, QSResultFileWSLastUsedColumn As Long, DlastRow As Long
Dim SampleName As Range, SampleNameStart As Range, DEColumnRng As Range, SerialNumberInput As Range, CRTcolumnRng As Range, cqConfRng As Range
Dim ResultFileSerialNumber As String
Dim rng As Range, rng2 As Range, rng3 As Range
Dim NameTargetArray As Variant, crtArray As Variant, cqConfidenceArray As Variant
Dim AMRorPathRNG As Range, AMRorPathstr As String


     FileToOpen = Application.GetOpenFilename(FileFilter:="Excel Files (*.XLSX), *.XLSX", Title:="Select all files needing analyzed", MultiSelect:=True)       'if file types change to csv or something else, this needs changed
          If Not IsArray(FileToOpen) Then
               isExit = True
               Exit Sub
          Else
          
          End If
          
          With OAdataWS
                 ClearStartRangeLastRowFung = .Cells(.Rows.count, "D").End(xlUp).Row
                 Set ClearStartRangeFung = OAdataWS.Range("D10:M" & ClearStartRangeLastRowFung)
                 ClearStartRangeFung.Clear
          End With
          
          With OAdataWS
                    .Range("D10").Value = "Sample Name"
                    .Range("E10").Value = "Target Name"
                    .Range("F10").Value = "Crt"
                    .Range("G10").Value = "Crt Avg"
                    .Range("H10").Value = "Crt SD"
                    .Range("I10").Value = "Cq Confidence"
                    .Range("J10").Value = "Min Cq Value"
                    .Range("K10").Value = "Full Quantitation"
                    .Range("L10").Value = "Infection %"
                    .Range("M10").Value = "Serial Number"
               With OAdataWS.Range("D10:O10")
                    .HorizontalAlignment = xlHAlignCenter
                    .Font.Size = 14
                    .Font.Bold = True
               End With
          End With
     
          AMRorPathstr = ""

     'select all result files at once
     For Each ResultFile In FileToOpen   '---------------------------------Import Result Files (Start)----------------------------
          Set QSResultFileWB = Workbooks.Open(ResultFile)
          Set QSResultFileWS = QSResultFileWB.Sheets("Results")
               TotalRows = 0
               Counter = 0
          With QSResultFileWS
               Set SampleName = .Range("A1:Q50").Find("Sample Name")
                    SampleNameLastRow = .Cells(.Rows.count, SampleName.Column).End(xlUp).Row
               Set SampleNameStart = .Range("D" & SampleName.Row).Offset(1, 0)
                    QSResultFileWSLastUsedColumn = .Cells(20, Columns.count).End(xlToLeft).Column   '------------------------------Sort Data to get Targets Grouped Together(Start)-----------------------------
     
               .Sort.SortFields.Clear
               .Sort.SortFields.Add2 Key:=Range("D21:D" & SampleNameLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
               .Sort.SortFields.Add2 Key:=Range("E21:E" & SampleNameLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
               With QSResultFileWS.Sort
                    .SetRange QSResultFileWS.Range(QSResultFileWS.Cells(20, 1), QSResultFileWS.Cells(SampleNameLastRow, QSResultFileWSLastUsedColumn))
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
               End With                                                                                                 '------------------------------Sort Data to get Targets Grouped Together(End)-----------------------------
              
          With QSResultFileWS.Range("D" & SampleNameStart.Row, "D" & SampleNameLastRow)
               Dim r As Range
               Dim RemoveThese As String, firstTarget As Range
                    
               For Each r In .Rows
                    RemoveThese = Right(r.Offset(0, -2).Value, 2)               'grabs the last 2 characters from the "Well Position" column - these are unique identifiers for the 2nd set of Finegoldia magna targets
                    Set firstTarget = r.Offset(0, 1)
                    
                    If r.Offset(0, 1).Value = "P. magnus_APTZ9PA" And (RemoveThese = "a8" Or RemoveThese = "b6" Or RemoveThese = "b8") Then
                         With r
                              .Value = vbNullString
                         End With
                    End If
                    If r.Value = vbNullString Then
                         GoTo CountNextRow
                    End If
                    If r.Offset(0, 1).Value = vbNullString Then     'This doesn't seem right
                         r.Value = vbNullString
                    End If
                    If Application.CountA(r) <> 0 Then      'if QSResultFileWS Sample Name value <> empty then
                         Counter = Counter + 1
                              'MODIFIED - 3/14/22
                         If AMRorPathstr = "" Then                    'if this string is empty, then neither Path nor AMR has been assigned yet
                              Set AMRorPathRNG = variableStor.Range("A1:D40").Find(firstTarget.Value, LookIn:=xlValues, lookAt:=xlWhole)      'find whatever the "firstTarget" is on OAdataWS
                              If Not (AMRorPathRNG Is Nothing) Then
                                   If AMRorPathRNG.Column = 3 Then                  'pathogen result file names column = 21
                                        AMRorPathstr = "Path"                             'set string = Path so I know to name Xeno as Path-Xeno
                                   ElseIf AMRorPathRNG.Column = 1 Then              'amr result file names column = 19
                                        AMRorPathstr = "AMR"                              'set string = AMR so I know to name Xeno as AMR-Xeno
                                   End If
                              Else
                                   MsgBox "Could not find target information on Variable Storage Worksheet"     'if nothing is found then the targets in Columns S:V were deleted - PROTECT
                                   isExit = True
                                   Exit Sub
                              End If
                         End If
                    End If
CountNextRow:            Next r
               TotalRows = Counter
          End With
               ResultFileSerialNumber = QSResultFileWS.Range("B1").Value
               Set rng = .Range("D21:E" & SampleNameLastRow)     'Sample Name and Target Name
               Set rng2 = .Range("I21:I" & SampleNameLastRow)    'CRT
               Set rng3 = .Range("M21:M" & SampleNameLastRow)    'Cq Confidence
               NameTargetArray = rng.Worksheet.Evaluate("FILTER(" & rng.Address & "," & rng.Columns(1).Address & "<>"""")")
               crtArray = rng2.Worksheet.Evaluate("FILTER(" & rng2.Address & "," & rng.Columns(1).Address & "<>"""")")
               cqConfidenceArray = rng3.Worksheet.Evaluate("FILTER(" & rng3.Address & "," & rng.Columns(1).Address & "<>"""")")
          End With

          With OAdataWS
               DlastRow = OAdataWS.Cells(Rows.count, "D").End(xlUp).Row
               Set DEColumnRng = OAdataWS.Range("D" & DlastRow).Offset(1, 0)
               Set SerialNumberInput = OAdataWS.Range("M" & DlastRow).Offset(1, 0)
               Set CRTcolumnRng = OAdataWS.Range("F" & DlastRow).Offset(1, 0)
               Set cqConfRng = OAdataWS.Range("I" & DlastRow).Offset(1, 0)
               OAdataWS.Range(DEColumnRng, "E" & (DEColumnRng.Row + TotalRows) - 1).Value = NameTargetArray
               OAdataWS.Range(CRTcolumnRng, "F" & (DEColumnRng.Row + TotalRows) - 1).Value = crtArray
               OAdataWS.Range(cqConfRng, "I" & (DEColumnRng.Row + TotalRows) - 1).Value = cqConfidenceArray
               OAdataWS.Range(SerialNumberInput, "M" & (DEColumnRng.Row + TotalRows) - 1).Value = ResultFileSerialNumber
          End With
               Erase NameTargetArray
               Erase crtArray
               Erase cqConfidenceArray
               If AMRorPathstr = "Path" Then           'after first file is placed on OAdataWS, check string - will tell me if the file was a pathogen or AMR file
                    Call Change_PathogenNames          'if pathogen, follow pathogen name change
                    AMRorPathstr = ""                       'clear string value in case more than 1 file was selected
               ElseIf AMRorPathstr = "AMR" Then
                    Call Change_AMRNames               'if AMR, follow AMR name change
                    AMRorPathstr = ""                       'clear string value in case more than 1 file was selected
               End If
          QSResultFileWS.Parent.Close False
     Next ResultFile                    '---------------------------------Import Result Files (End)-------------------------------


          'clear data on reruns to pull WS
     With PullReruns
          .Range("A9:C1000").Clear
     End With

          'apply cell formatting to imported data
Dim afterImportLastRow As Long
     With OAdataWS
          afterImportLastRow = .Cells(.Rows.count, "D").End(xlUp).Row
          
          With .Range("D11:E" & afterImportLastRow, "M11:M" & afterImportLastRow)
               .NumberFormat = "@"           'apply text formatting to ranges
          End With
          With .Range("F11:J" & afterImportLastRow)
               .NumberFormat = "0.000"       'apply 3 decimal places formatting to ranges
          End With
          With .Range("K11:K" & afterImportLastRow)
               .NumberFormat = "0.00E+00"
          End With
          With .Range("L11:L" & afterImportLastRow)
               .NumberFormat = "0.00%"
          End With
          With .Range("D10:E" & afterImportLastRow, "G10:O" & afterImportLastRow)
               .HorizontalAlignment = xlHAlignCenter
               .Columns.AutoFit
          End With
     End With
     
'     Call OptimizeCode_End
End Sub
