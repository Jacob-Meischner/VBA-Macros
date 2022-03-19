Attribute VB_Name = "G_Full_Quant_Interp"
Public Sub Full_Quant_Interpretation()

'Dim ST As Single
'ST = timer

'     Call OptimizeCode_Begin

     'MIN CQ & FULL QUANTITATIVE RESULT
Dim DlastRow As Long, i As Long
Dim accNum As Range, firstTarget As Range, cRT As Range, minCqRng As Range, fullQuant As Range, crtReplicate As Range
Dim CqArr As Variant, cqArrCount As Long, replicateCqConf As Range
     Dim minCQ As Double, FullQuantResult As Double
Dim crtCutoff As Double, cqConfCutoff As Double, ndCounter As Long
     'TARGET CHECKS + INVALIDS
Dim pathogenArr As Variant, pathogenMatch As Variant, amrArr As Variant, amrMatch As Variant, xenoArr As Variant, xenoMatch As Variant
Dim invalidArr As Variant, allINVALIDresults As Range, j As Long
     'PERCENT INFECTION
Dim percentArr As Variant, percentCount As Long, rowArr As Variant
Dim SumAllQuantResults As Double, dynamicQuantResults As Double, finalPercentCalc As Double, infectionPercent As Range


     pathogenArr = Array(Aci_bau.Value, Ana_prev.Value, Bact_frag.Value, Bact_vulg.Value, Can_albi.Value, Cit_fre.Value, Clost_perf.Value, Clost_sept.Value, Coryn_str.Value, Entero_aero.Value, Entero_cloac.Value, _
                                   E_faecalis.Value, E_faecium.Value, E_coli.Value, F_magna.Value, Kleb_oxy.Value, Kleb_pneu.Value, _
                                   Pept_asa.Value, Pept_ana.Value, Prev_bivia.Value, Prev_loe.Value, Pro_mir.Value, Pro_vul.Value, Pseud_aer.Value, Sal_M.Value, Sal_N.Value, Serr_marc.Value, _
                                   Staph_aur.Value, Staph_epid.Value, Staph_haem.Value, Staph_lugd.Value, Staph_sapro.Value, Strep_agalac.Value, Strep_pneu.Value, Strep_pyo.Value)
     
     amrArr = Array(ACC.Value, ampC.Value, BILlatCMY.Value, CTXpool.Value, dfrA5A1.Value, DHA.Value, FOX.Value, GES.Value, IMPpool.Value, KPC.Value, mcr1.Value, mecA.Value, moxCMY.Value, nfsA.Value, OXApool.Value, _
                         OXA1.Value, PER12.Value, QnrASB.Value, SHV.Value, Sul12.Value, TEM.Value, tetBMS.Value, vanA12B.Value, VEB.Value, VIM.Value)
     
     xenoArr = Array(Path_Xeno.Value, AMR_Xeno.Value)
     
     
     With OAdataWS
          DlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
     End With

          'CUTOFFS
     crtCutoff = 30
     cqConfCutoff = 0.7
          'INFECTION %
     percentCount = 1
     ReDim percentArr(1 To 36) As Variant                                            'searching every 3rd cell, which means the # of possibilties = total number of targets (35) - leave extra space just in case
     ReDim rowArr(1 To 36) As String
          'NOT DETECTED/INVALID
     ndCounter = 0
     ReDim invalidArr(1 To 37)
     
          'setting Min Cq Value and Full Quantitative fields
     For i = 11 To DlastRow Step 3
     
          Set accNum = OAdataWS.Cells(i, 4)
          Set firstTarget = OAdataWS.Cells(i, 5)
          Set cRT = OAdataWS.Cells(i, 6)
          Set minCqRng = OAdataWS.Cells(i, 10)
          Set fullQuant = OAdataWS.Cells(i, 11)
          Set infectionPercentLoader = OAdataWS.Cells(i, 12)
          
          pathogenMatch = Application.match(firstTarget.Value, pathogenArr, 0)        'see if firstTarget = any pathogen target
          amrMatch = Application.match(firstTarget.Value, amrArr, 0)                       'see if firstTarget = any AMR target
          xenoMatch = Application.match(firstTarget.Value, xenoArr, 0)                     'see if firstTarget = either Xeno (results will be applied the same)
          
          If Not IsError(pathogenMatch) Then         '--------------------------------------------------------Pathogens--------------------------------------------------------------
               cqArrCount = 1
               ReDim CqArr(1 To 3) As Variant          'resets array every time target changes
               
               For Each crtReplicate In OAdataWS.Range(cRT, cRT.Offset(2, 0)).Cells
                    If crtReplicate <= crtCutoff Then
                         Set replicateCqConf = crtReplicate.Offset(0, 3)        'set range to grab the replicate specific Cq Confidence
                         If replicateCqConf.Value >= cqConfCutoff Then         'only isolate single cq value if it's <= threshold and confidence >= 0.7
                              CqArr(cqArrCount) = crtReplicate.Value
                              cqArrCount = cqArrCount + 1
                         End If
                    End If
               Next crtReplicate
               
               minCQ = Application.WorksheetFunction.Min(CqArr)
               If minCQ <> 0# Then      'if at least 1/3 replicates has qualifying Cq Value & Cq Conf then take what lowest cq value from CqArr
                         'MIN CQ VALUE
                    With minCqRng
                         .Value = minCQ
                         .Interior.Color = RGB(0, 255, 0)
                    End With
                         'FULL QUANT RESULT
                    FullQuantResult = FindFullQuantResult(firstTarget.Value, minCQ)       'plug Cq value into correct pathogen specific calculation
                    With fullQuant
                         .Value = FullQuantResult
                    End With
                         'PERCENT INFECTION
                    percentArr(percentCount) = fullQuant.Value                       'Store quantitative value for particular target
                    rowArr(percentCount) = infectionPercentLoader.Address       'Offset 1 column where quantitative results were found and store this address (this location = destination for Infection % for this target) - when the TargetName = Path-Xeno, it's the end of the patients data, sum all numerical values in array and divide individual numbers/sum and return % to OAdataWS.cells(row,12)
                    percentCount = percentCount + 1
               Else
                    With minCqRng
                         .Value = "Not Detected"                 'else if minCq = 0.000 then set minCq value = Not Detected
                    End With
                    ndCounter = ndCounter + 1                         'keeping track of total # not detected results
                    invalidArr(ndCounter) = minCqRng.Address     'and their locations
               End If
               
          ElseIf Not IsError(amrMatch) Then  '--------------------------------------------------------AMR--------------------------------------------------------------
               cqArrCount = 1
               ReDim CqArr(1 To 3) As Variant
               
                For Each crtReplicate In OAdataWS.Range(cRT, cRT.Offset(2, 0)).Cells
                    If crtReplicate <= crtCutoff Then
                         Set replicateCqConf = crtReplicate.Offset(0, 3)
                         If replicateCqConf.Value >= cqConfCutoff Then
                              CqArr(cqArrCount) = crtReplicate.Value
                              cqArrCount = cqArrCount + 1
                         End If
                    End If
               Next crtReplicate
               
               minCQ = Application.WorksheetFunction.Min(CqArr)
               If minCQ <> 0# Then
                    With minCqRng
                         .Value = minCQ
                         .Interior.Color = RGB(0, 255, 0)
                    End With
               Else
                    With minCqRng
                         .Value = "Not Detected"
                    End With
                    ndCounter = ndCounter + 1
                    invalidArr(ndCounter) = minCqRng.Address
               End If
               
          ElseIf Not IsError(xenoMatch) Then           '--------------------------------------------------------Xeno--------------------------------------------------------------
               cqArrCount = 1
               ReDim CqArr(1 To 3) As Variant
               
                For Each crtReplicate In OAdataWS.Range(cRT, cRT.Offset(2, 0)).Cells
                    If crtReplicate <= crtCutoff Then
                         Set replicateCqConf = crtReplicate.Offset(0, 3)
                         If replicateCqConf.Value >= cqConfCutoff Then
                              CqArr(cqArrCount) = crtReplicate.Value
                              cqArrCount = cqArrCount + 1
                         End If
                    End If
               Next crtReplicate
               
               minCQ = Application.WorksheetFunction.Min(CqArr)
               If minCQ <> 0# Then
                    With minCqRng
                         .Value = minCQ
                         .Interior.Color = RGB(0, 255, 0)
                    End With
               Else
                    With minCqRng
                         .Value = "Not Detected"
                    End With
                    ndCounter = ndCounter + 1
                    invalidArr(ndCounter) = minCqRng.Address
               End If
          End If
          
               '----------------------Always check if the last resulted target was either AMR-Xeno or Path-Xeno - Must reset counters after every accession number changes--------------------
               
          If firstTarget.Value = AMR_Xeno.Value Then
               If ndCounter = 26 Then            ' And accNum <> amrNTC)  '26 = total number of AMR targets + Xeno - if everything = Not detected (IE. ndCounter = 26) then set all MinCq fields to yellow
                    ReDim Preserve invalidArr(1 To ndCounter)
                    For j = LBound(invalidArr, 1) To UBound(invalidArr, 1)
                         Set allINVALIDresults = OAdataWS.Range(invalidArr(j))
                              With allINVALIDresults
                                   .Interior.Color = RGB(255, 255, 0)
                              End With
                    Next j
                    ReDim invalidArr(1 To 37)
                    rerunAccNum accNum.Value, firstTarget.Value
                    
               End If
               ndCounter = 0            'if the code reaches this point, that means AMR-Xeno result was just set, meaning the patient is changing on next iteration (custom sorting assures Xeno will always be the last target for every patient)
          ElseIf firstTarget.Value = Path_Xeno.Value Then
               If ndCounter = 36 Then             '36 = total number of Pathogen targets + Xeno - if everything = Not detected (IE. ndCounter = 36) then set all MinCq fields to yellow
                    ReDim Preserve invalidArr(1 To ndCounter)
                    For j = LBound(invalidArr, 1) To UBound(invalidArr, 1)
                         Set allINVALIDresults = OAdataWS.Range(invalidArr(j))
                              With allINVALIDresults
                                   .Interior.Color = RGB(255, 255, 0)
                              End With
                    Next j
                    ReDim invalidArr(1 To 37)
                    rerunAccNum accNum.Value, firstTarget.Value
                    
               ElseIf percentCount > 1 Then
                    ReDim Preserve percentArr(1 To percentCount - 1)            'trim down array to only include all quantitative results that were stored for every patient
                    ReDim Preserve rowArr(1 To percentCount - 1)                'apply equal trimming to rowArr - there should always be the same number of locations stored & quantitative results stored
                    SumAllQuantResults = Application.WorksheetFunction.Sum(percentArr)          'add all quantitative results in percentArr

                    For j = LBound(percentArr, 1) To UBound(percentArr, 1)           'starting from the bottom, search through all levels of percentArr
                         dynamicQuantResults = percentArr(j)                              'set stored quantitative value = dynamicQuantResults
                         finalPercentCalc = dynamicQuantResults / SumAllQuantResults      'divide individual quantitative results / sum of all quantitative results (per patient)
                         Set infectionPercent = OAdataWS.Range(rowArr(j))                 'set the final destination = to the corresponding rowArr level (this only works because both were stored at the same time)
                              With infectionPercent
                                   .Value = finalPercentCalc
                              End With
                    Next j
               End If
               
               ndCounter = 0       'if the code reaches this point, that means Path-Xeno result was just set, meaning the patient is changing on next iteration (custom sorting assures Xeno will always be the last target for every patient)
               ReDim percentArr(1 To 36) As Variant                   'resize the array back to original size
               ReDim rowArr(1 To 36) As String                        'resize the array back to original size
               percentCount = 1                                            'set count back to 1 to accurately count the next patient's targets
               dynamicQuantResults = 0                                'clear variables holding current patients values so there's no confusion
               SumAllQuantResults = 0
               
          End If
     Next i
     
'     MsgBox "Macro took: " & timer - ST & " seconds to complete."
'     Call OptimizeCode_End
End Sub
