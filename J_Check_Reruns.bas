Attribute VB_Name = "J_Check_Reruns"
Function rerunAccNum(accNum As String, targ As String)

'     Call OptimizeCode_Begin

Dim controlsArr As Variant, controlMatch As Variant
          'CHECK IF ACCNUM IS ANY CONTROLS - IF IT IS THEN EXIT FUNCTION - DON'T WANT CONTROLS TO POPULATE ON RERUNS TO PULL
     controlsArr = Array(pathPTC, pathPEC, pathNTC, pathNEC, amrPTC, amrPEC, amrNTC, amrNEC)
     controlMatch = Application.match(accNum, controlsArr, 0)
     
     If Not IsError(controlMatch) Then
          Exit Function
     Else
          
     End If

Dim OAplate As String, pathogenArr As Variant, amrArr As Variant
Dim pathMatch As Variant, amrMatch As Variant
Dim dupPT As Variant, dupSearchRng As Range, combineRNG As Range
Dim FindFullID As Range, Rerun2PullLastRow As Long, Rerun2PullDest As Range
Dim Counter As Long, checkTotalTests As String

     pathogenArr = Array(Aci_bau.Value, Ana_prev.Value, Bact_frag.Value, Bact_vulg.Value, Can_albi.Value, Cit_fre.Value, Clost_perf.Value, Clost_sept.Value, Coryn_str.Value, Entero_aero.Value, Entero_cloac.Value, _
                                   E_faecalis.Value, E_faecium.Value, E_coli.Value, F_magna.Value, Kleb_oxy.Value, Kleb_pneu.Value, Pept_asa.Value, Pept_ana.Value, Prev_bivia.Value, Prev_loe.Value, Pro_mir.Value, _
                                   Pro_vul.Value, Pseud_aer.Value, Sal_M.Value, Sal_N.Value, Serr_marc.Value, Staph_aur.Value, Staph_epid.Value, Staph_haem.Value, Staph_lugd.Value, Staph_sapro.Value, Strep_agalac.Value, _
                                   Strep_pneu.Value, Strep_pyo.Value, Path_Xeno.Value)
                                   
     amrArr = Array(ACC.Value, ampC.Value, BILlatCMY.Value, CTXpool.Value, dfrA5A1.Value, DHA.Value, FOX.Value, GES.Value, IMPpool.Value, KPC.Value, mcr1.Value, mecA.Value, moxCMY.Value, nfsA.Value, OXApool.Value, _
                         OXA1.Value, PER12.Value, QnrASB.Value, SHV.Value, Sul12.Value, TEM.Value, tetBMS.Value, vanA12B.Value, VEB.Value, VIM.Value, AMR_Xeno.Value)
                         
     pathMatch = Application.match(targ, pathogenArr, 0)
     amrMatch = Application.match(targ, amrArr, 0)
     
     If Not IsError(pathMatch) Then
          OAplate = "Pathogen"
     ElseIf Not IsError(amrMatch) Then
          OAplate = "AMR"
     End If
     
     Set FindFullID = importInfoWS.Range("B12:B59").Find(accNum, LookIn:=xlValues, lookAt:=xlPart)

     If Not (FindFullID Is Nothing) Then
          If FindFullID.Borders.Weight = xlThin Then
          
               With PullReruns
                    Rerun2PullLastRow = .Cells(.Rows.count, "A").End(xlUp).Row
                    Set dupSearchRng = PullReruns.Range("A1:A" & Rerun2PullLastRow)        'range to search for duplicates on reruns to pull - need to know if col C should have both Path & AMR
                    Set Rerun2PullDest = .Range("A" & Rerun2PullLastRow).Offset(1, 0)
               End With
               
               dupPT = Application.match(FindFullID.Value, dupSearchRng, 0)
               
               If Not IsError(dupPT) Then                                          'if duplicate is found on reruns to pull then
                    Set combineRNG = PullReruns.Cells(dupPT, 3)            'check the value in col C for selected accession number
                    
                         For Counter = 1 To Len(combineRNG.Value)
                              checkTotalTests = Mid(combineRNG.Value, Counter, 1)    'should detected if a patient has been already marked to be rerun for both Pathogens & AMR
                              If checkTotalTests = "&" Then
                                   MsgBox "Patient already marked to be rerun for both Pathogens & AMR."
                                   Exit Function
                              Else

                              End If
                         Next Counter

                    If combineRNG <> OAplate Then                               'if current col C value <> OAplate
                         With combineRNG
                              .Value = combineRNG.Value & " & " & OAplate    'change col C value to existing value & OAplate
                         End With
                    Else
                         MsgBox "Patient is already marked to be rerun.  See 'Reruns To Pull worksheet'.", vbCritical
                         Exit Function
                    End If
               Else
                    With Rerun2PullDest
                         .Value = FindFullID.Value
                    End With
                    With Rerun2PullDest.Offset(0, 1)
                         .Value = Trim(FindFullID.Offset(0, -1).Value)
                    End With
                    With Rerun2PullDest.Offset(0, 2)
                         .Value = OAplate
                         .Interior.Color = RGB(255, 255, 0)
                    End With
               End If
          End If
     Else
          MsgBox ("Patient name: " & accNum & " was not found on 'Import Patient Information' Worksheet." & vbNewLine & vbNewLine)
          Exit Function
     End If
    
Dim postRerunlastRow As Long

     With PullReruns
          postRerunlastRow = .Cells(.Rows.count, "A").End(xlUp).Row
          
          With .Range("A9:A" & postRerunlastRow, "C9:C" & postRerunlastRow)
                .ColumnWidth = 30
                .Font.Size = 14
                .WrapText = True
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
           End With
           With .Range("B9:B" & postRerunlastRow)
                .Font.Size = 14
                .Font.Bold = True
                .ColumnWidth = 12
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlVAlignCenter
           End With
     End With
          
'     Call OptimizeCode_End
          
End Function
