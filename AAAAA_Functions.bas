Attribute VB_Name = "AAAAA_Functions"
Function getUniques(a, Optional ZeroBased As Boolean = True)

Dim tmp: tmp = Application.Transpose(WorksheetFunction.Unique(a))
If ZeroBased Then ReDim Preserve tmp(0 To UBound(tmp) - 1)
getUniques = tmp

End Function

Function FindFullQuantResult(target As String, minCQ As Double)

     Select Case target
          Case Aci_bau
               FindFullQuantResult = 10 ^ ((minCQ - 40.641) / -3.594)      'A. baumannii  1
          Case Bact_frag
               FindFullQuantResult = 10 ^ ((minCQ - 39.328) / -3.5846)     'B. fragilis  2
          Case Bact_vulg
               FindFullQuantResult = 10 ^ ((minCQ - 37.215) / -3.2223)     'B. vulgatus  3
          Case Cit_fre
               FindFullQuantResult = 10 ^ ((minCQ - 39.583) / -3.7755)     'C. freundii  4
          Case Clost_perf
               FindFullQuantResult = 10 ^ ((minCQ - 40.197) / -3.6059)     'C. perfringens  5
          Case Clost_sept
               FindFullQuantResult = 10 ^ ((minCQ - 41.505) / -3.901)      'C. septicum  6
          Case Coryn_str
               FindFullQuantResult = 10 ^ ((minCQ - 40.675) / -3.6658)     'C. striatium  7
          Case Entero_aero
               FindFullQuantResult = 10 ^ ((minCQ - 39.887) / -3.6172)     'E. aerogenes  8
          Case Entero_cloac
               FindFullQuantResult = 10 ^ ((minCQ - 39.661) / -3.4956)     'E. cloacae  9
          Case E_faecalis
               FindFullQuantResult = 10 ^ ((minCQ - 39.305) / -3.6476)     'E. faecalis  10
          Case E_faecium
               FindFullQuantResult = 10 ^ ((minCQ - 40.105) / -3.6276)     'E. faecium  11
          Case E_coli
               FindFullQuantResult = 10 ^ ((minCQ - 39.568) / -3.5068)     'E. coli  12
          Case F_magna
               FindFullQuantResult = 10 ^ ((minCQ - 40.014) / -3.64)       'P. magnus  13
          Case Kleb_pneu
               FindFullQuantResult = 10 ^ ((minCQ - 39.278) / -3.5939)     'K. pneumoniae  14
          Case Kleb_oxy
               FindFullQuantResult = 10 ^ ((minCQ - 40.603) / -3.6387)     'K. oxytoca  15
          Case Pept_ana
               FindFullQuantResult = 10 ^ ((minCQ - 41.301) / -3.7569)     'P. anaerobius  16
          Case Pept_asa
               FindFullQuantResult = 10 ^ ((minCQ - 40.001) / -3.3932)     'P. asaccharolyticus  17
          Case Ana_prev
               FindFullQuantResult = 10 ^ ((minCQ - 41.032) / -3.6336)     'Anaerococcus prevotii  18
          Case Prev_bivia
               FindFullQuantResult = 10 ^ ((minCQ - 40.028) / -3.678)      'P. bivia  19
          Case Prev_loe
               FindFullQuantResult = 10 ^ ((minCQ - 40.952) / -3.7278)     'P. loescheii  20
          Case Pro_mir
               FindFullQuantResult = 10 ^ ((minCQ - 40.556) / -3.6602)     'P. mirabilis  21
          Case Pro_vul
               FindFullQuantResult = 10 ^ ((minCQ - 39.635) / -3.6576)     'P. vulgaris  22
          Case Pseud_aer
               FindFullQuantResult = 10 ^ ((minCQ - 40.236) / -3.6581)     'P. aeruginosa  23
          Case Sal_M
               FindFullQuantResult = 10 ^ ((minCQ - 40.958) / -3.6496)     'S. montevideo  24
          Case Sal_N
               FindFullQuantResult = 10 ^ ((minCQ - 40.915) / -3.70335)    'S. enterica  25
          Case Serr_marc
               FindFullQuantResult = 10 ^ ((minCQ - 39.433) / -3.5346)     'S. marcescens  26
          Case Staph_aur
               FindFullQuantResult = 10 ^ ((minCQ - 38.18) / -3.4242)      'S. aureus  27
          Case Staph_epid
               FindFullQuantResult = 10 ^ ((minCQ - 38.371) / -3.5261)     'S. epidermidis  28
          Case Staph_haem
               FindFullQuantResult = 10 ^ ((minCQ - 40.915) / -3.5905)     'S. haemolyticus  29
          Case Staph_lugd
               FindFullQuantResult = 10 ^ ((minCQ - 40.063) / -3.5843)     'S. lugdunensis  30
          Case Staph_sapro
               FindFullQuantResult = 10 ^ ((minCQ - 38.058) / -3.3907)     'S. saprophyticus  31
          Case Strep_agalac
               FindFullQuantResult = 10 ^ ((minCQ - 39.489) / -3.5024)     'S. agalactiae  32
          Case Strep_pneu
               FindFullQuantResult = 10 ^ ((minCQ - 38.637) / -3.5075)     'S. pneumoniae  33
          Case Strep_pyo
               FindFullQuantResult = 10 ^ ((minCQ - 40.903) / -3.5915)     'S. pyogenes  34
          Case Can_albi
               FindFullQuantResult = 10 ^ ((minCQ - 39.12) / -3.6163)      'C. albicans  35
     End Select
End Function
Function ActualControlLB(actualControl As String) As Variant

Dim filteredRng As Range, minCqControl As Range, detectedTarget As Range
Dim conActualCount As Long, actualDetected As Variant

     With OAdataWS
          DlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
          Set filteredRng = .Range("J11:J" & DlastRow)
     End With
     
     ReDim actualDetected(1 To 37) As Variant
     conActualCount = 1

     For Each minCqControl In filteredRng.SpecialCells(xlCellTypeVisible).Cells
          If (minCqControl.Value <> vbNullString And IsNumeric(minCqControl) = True) Then
               Set detectedTarget = minCqControl.Offset(0, -5)
               actualDetected(conActualCount) = detectedTarget.Value
               conActualCount = conActualCount + 1
          End If
     Next minCqControl

     Select Case actualControl
          
          Case pathPTC
               If conActualCount > 1 Then
               ReDim Preserve actualDetected(1 To conActualCount - 1)
                    
                   With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)          'sort listbox alphabetically
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
               
               
          Case pathPEC
                If conActualCount > 1 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If

          Case pathNEC
               If conActualCount > 1 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
               
          Case pathNTC
               If conActualCount >= 2 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
          
               'AMRS
          Case amrPTC
                If conActualCount >= 2 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
               
          Case amrPEC
                If conActualCount > 1 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
               
          Case amrNEC
                If conActualCount > 1 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
               
          Case amrNTC
               If conActualCount > 1 Then
                    ReDim Preserve actualDetected(1 To conActualCount - 1)
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .List = actualDetected
                         Call UserFormSortAZ(WoundQuickFilterByID.LB_Actual_Detected)
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = (conActualCount - 1)
                    End With
               Else
                    With WoundQuickFilterByID.LB_Actual_Detected
                         .Clear
                    End With
                    With WoundQuickFilterByID.LBL_Actual_Total
                         .Caption = 0
                    End With
               End If
     End Select
     
          
End Function
Function ExpectedControlLB(expectedControl As String)

Dim pathPTCarr As Variant, pathPECarr As Variant, pathNECarr As Variant
Dim amrPTCarr As Variant, amrPECNECarr As Variant

          'PATHOGENS
     pathPTCarr = Array(Aci_bau.Value, Ana_prev.Value, Bact_frag.Value, Bact_vulg.Value, Can_albi.Value, Cit_fre.Value, Clost_perf.Value, Clost_sept.Value, Coryn_str.Value, Entero_aero.Value, Entero_cloac.Value, _
                                   E_faecalis.Value, E_faecium.Value, E_coli.Value, F_magna.Value, Kleb_oxy.Value, Kleb_pneu.Value, Pept_asa.Value, Pept_ana.Value, Prev_bivia.Value, Prev_loe.Value, Pro_mir.Value, _
                                   Pro_vul.Value, Pseud_aer.Value, Sal_M.Value, Sal_N.Value, Serr_marc.Value, Staph_aur.Value, Staph_epid.Value, Staph_haem.Value, Staph_lugd.Value, Staph_sapro.Value, _
                                   Strep_agalac.Value, Strep_pneu.Value, Strep_pyo.Value, Path_Xeno.Value)
     pathPECarr = Array(Can_albi.Value, Path_Xeno.Value)
     pathNECarr = Array(Path_Xeno.Value)

          'AMRS
     amrPTCarr = Array(ACC.Value, ampC.Value, BILlatCMY.Value, CTXpool.Value, dfrA5A1.Value, DHA.Value, FOX.Value, GES.Value, IMPpool.Value, KPC.Value, mcr1.Value, mecA.Value, moxCMY.Value, _
                                   nfsA.Value, OXApool.Value, OXA1.Value, PER12.Value, QnrASB.Value, SHV.Value, Sul12.Value, TEM.Value, tetBMS.Value, vanA12B.Value, VEB.Value, VIM.Value, AMR_Xeno.Value)
     amrPECNECarr = Array(AMR_Xeno.Value)


     Select Case expectedControl
     
          Case pathPTC
                With WoundQuickFilterByID.LB_Expected_Detected
                    .List = pathPTCarr
                    Call UserFormSortAZ(WoundQuickFilterByID.LB_Expected_Detected)
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
               
          Case pathPEC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .List = pathPECarr
                    Call UserFormSortAZ(WoundQuickFilterByID.LB_Expected_Detected)
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
          
          
          Case pathNEC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .List = pathNECarr
                    Call UserFormSortAZ(WoundQuickFilterByID.LB_Expected_Detected)
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
               
          Case pathNTC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .Clear
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
               
               'AMRS
          Case amrPTC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .List = amrPTCarr
                    Call UserFormSortAZ(WoundQuickFilterByID.LB_Expected_Detected)
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
          
          Case amrPEC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .List = amrPECNECarr
                    Call UserFormSortAZ(WoundQuickFilterByID.LB_Expected_Detected)
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
          
          Case amrNEC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .List = amrPECNECarr
                    Call UserFormSortAZ(WoundQuickFilterByID.LB_Expected_Detected)
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
          
          Case amrNTC
               With WoundQuickFilterByID.LB_Expected_Detected
                    .Clear
               End With
               For i = 0 To WoundQuickFilterByID.LB_Expected_Detected.ListCount - 1
                    expectedCount = expectedCount + 1
               Next i
               
               If expectedCount > 0 Then
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = expectedCount
                    End With
               Else
                    With WoundQuickFilterByID.LBL_Expected_Total
                         .Caption = 0
                    End With
               End If
     End Select
     
               With WoundQuickFilterByID.LB_Expected_Detected
                    .ForeColor = RGB(0, 150, 0)
               End With
End Function
Public Sub UserFormSortAZ(myListBox As MSForms.ListBox, Optional resetMacro As String)

'Create variables
Dim j As Long
Dim i As Long
Dim temp As Variant

''Reset the listBox into standard order
'If resetMacro <> "" Then
'    Run resetMacro, myListBox
'End If

'Use Bubble sort method to put listBox in A-Z order
With myListBox
    For j = 0 To .ListCount - 2
        For i = 0 To .ListCount - 2
            If LCase(.List(i)) > LCase(.List(i + 1)) Then
                temp = .List(i)
                .List(i) = .List(i + 1)
                .List(i + 1) = temp
            End If
        Next i
    Next j
End With

End Sub


