Attribute VB_Name = "I_Prepare_Worklist_Setup"
Sub Prepare_Worklist()
          'DATE ADDED - 3/13/22
          
'     Call OptimizeCode_Begin

Dim targetArr As Variant, pathName As Variant, pathHeaderDest As Range
Dim worklistBlastRow As Long, ClearNonHeaders As Range
Dim postClearlastRow As Long, worklistAccNum As Range
Dim controlsArr As Variant, controlMatch As Variant
     
     targetArr = Array(Aci_bau.Value, Ana_prev.Value, Bact_frag.Value, Bact_vulg.Value, Can_albi.Value, Cit_fre.Value, Clost_perf.Value, Clost_sept.Value, Coryn_str.Value, Entero_aero.Value, Entero_cloac.Value, _
                                   E_faecalis.Value, E_faecium.Value, E_coli.Value, F_magna.Value, Kleb_oxy.Value, Kleb_pneu.Value, _
                                   Pept_asa.Value, Pept_ana.Value, Prev_bivia.Value, Prev_loe.Value, Pro_mir.Value, Pro_vul.Value, Pseud_aer.Value, Sal_M.Value, Sal_N.Value, Serr_marc.Value, _
                                   Staph_aur.Value, Staph_epid.Value, Staph_haem.Value, Staph_lugd.Value, Staph_sapro.Value, Strep_agalac.Value, Strep_pneu.Value, Strep_pyo.Value, Path_Xeno.Value, _
                                   ACC.Value, ampC.Value, BILlatCMY.Value, CTXpool.Value, dfrA5A1.Value, DHA.Value, FOX.Value, GES.Value, IMPpool.Value, KPC.Value, mcr1.Value, mecA.Value, moxCMY.Value, _
                                   nfsA.Value, OXApool.Value, OXA1.Value, PER12.Value, QnrASB.Value, SHV.Value, Sul12.Value, TEM.Value, tetBMS.Value, vanA12B.Value, VEB.Value, VIM.Value, AMR_Xeno.Value)
     
     controlsArr = Array(pathPTC, pathPEC, pathNTC, pathNEC, amrPTC, amrPEC, amrNTC, amrNEC)
     
     
          'prepare Worklist View WS for new incoming data
     With WorklistView
          worklistBlastRow = .Cells(.Rows.count, "B").End(xlUp).Row
          Set ClearNonHeaders = WorklistView.Range("B4:EF" & worklistBlastRow).Offset(1, 0)                            '"WorklistView.Cells(5, 2), WorklistView.Cells(worklistBlastRow, 120))
               With ClearNonHeaders
                    .Clear
               End With
          Set pathHeaderDest = WorklistView.Range("C3")
     End With
     
          'populate target headers based on constant range variables - order arrangement inside targetArr matters!
     For Each pathName In targetArr
          With pathHeaderDest
               .Value = pathName
          End With
          Set pathHeaderDest = pathHeaderDest.Offset(0, 1)
     Next pathName
     
     With WorklistView
          postClearlastRow = .Cells(.Rows.count, "B").End(xlUp).Row
          Set worklistAccNum = WorklistView.Range("B" & postClearlastRow).Offset(1, 0)
     End With

      'populate accession numbers from OAdataWS (Col D) to Worklist View (Col B)
Dim tblVar As ListObject, uniqueAccNumRNG As Range
Dim uniqueArr As Variant, uniqueVal As Variant, removeControls As New Collection, clsCounter As Long, MyObject As Variant

     Set tblVar = OAdataWS.ListObjects("Table1")
     Set uniqueAccNumRNG = tblVar.ListColumns(1).DataBodyRange
     uniqueArr = uniqueAccNumRNG
     uniqueArr = getUniques(uniqueArr)            'Pass all values to getUniques function to remove all duplicates - controls are still located in this array, must remove them
     
     For Each uniqueVal In uniqueArr              'loop through all unique values stored in uniqueArr
          Dim Inst As New Class1                       'create a new instance for each item
          clsCounter = clsCounter + 1                  'create counter - whatever this counter is will be used as the reference key when stored into the collection
          controlMatch = Application.match(uniqueVal, controlsArr, 0)
          If IsError(controlMatch) Then                'if uniqueVal <> any of the controls in controlsArr
               Inst.InstanceName = uniqueVal           'then add this value in the object instance
               removeControls.Add Inst, CStr(clsCounter)    'add named object to collection using the key of clsCounter
               Set Inst = Nothing                                'clear current instance to make room for a new one
          Else
               clsCounter = clsCounter - 1                  'else if a control is found then subtract 1 from clsCounter to retain accurate number of added instances
          End If
     Next uniqueVal
     
     For Each MyObject In removeControls
           With worklistAccNum
               .Value = MyObject.InstanceName
          End With
          Set worklistAccNum = worklistAccNum.Offset(1, 0)
     Next MyObject
     
     For clsCounter = 1 To removeControls.count
          removeControls.Remove 1
     Next
     
'     With WorklistView.Range("B:B")
'          .HorizontalAlignment = xlHAlignCenter
'          .VerticalAlignment = xlVAlignCenter
'     End With
     
'     Call OptimizeCode_End
End Sub

Public Sub setWorklistViewValues()

'     Call OptimizeCode_Begin

     'Worklist View variables
Dim wVBlastrow, wVAccSearch As Range
Dim wVTargetName As Range, wVTargetSearch As Range, wVTargetlastCol As Long
Dim wVAccNumMatch As Variant, headerSearchRNG As Range
Dim wVMinCq As Range, wVQuantResult As Range, wVInfection As Range
     'OA Data variables
Dim i As Long, OAdataDlastRow As Long, OAAccNum As Range, OATarget As Range, OAminCq As Range, OAfullQuant As Range, OAinfection As Range

     With WorklistView
          wVBlastrow = .Cells(.Rows.count, "B").End(xlUp).Row
          Set wVAccSearch = WorklistView.Range("B1:B" & wVBlastrow)
          wVTargetlastCol = .Cells(3, .Columns.count).End(xlToLeft).Column
          Set headerSearchRNG = WorklistView.Range(WorklistView.Cells(3, 1), WorklistView.Cells(3, wVTargetlastCol))
     End With
     
     With OAdataWS
          OAdataDlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
     End With

     For i = 11 To OAdataDlastRow
          Set OAAccNum = OAdataWS.Cells(i, 4)
          Set OATarget = OAdataWS.Cells(i, 5)
          Set OAminCq = OAdataWS.Cells(i, 10)
          Set OAfullQuant = OAdataWS.Cells(i, 11)
          Set OAinfection = OAdataWS.Cells(i, 12)
          
          If OAminCq <> vbNullString Then
          wVAccNumMatch = Application.match(OAAccNum, wVAccSearch, 0)
               If Not IsError(wVAccNumMatch) Then
                    Set wVTargetSearch = headerSearchRNG.Find(OATarget.Value, LookIn:=xlValues, lookAt:=xlWhole)
                    If Not (wVTargetSearch Is Nothing) Then
                         If IsNumeric(OAminCq) = True Then
                              Set wVMinCq = WorklistView.Cells(wVAccNumMatch, wVTargetSearch.Column)
                              Set wVQuantResult = wVMinCq.Offset(0, 1)
                              Set wVInfection = wVQuantResult.Offset(0, 1)
                              With wVMinCq
                                   .Value = OAminCq.Value
                                   .NumberFormat = "0.000"
                                   .HorizontalAlignment = xlHAlignCenter
                                   .VerticalAlignment = xlVAlignCenter
                              End With
                              If (OAfullQuant <> vbNullString And OAinfection <> vbNullString) Then
                                   With wVQuantResult
                                        .Value = OAfullQuant.Value
                                        .NumberFormat = "0.00E+00"
                                        .HorizontalAlignment = xlHAlignCenter
                                        .VerticalAlignment = xlVAlignCenter
                                   End With
                                   With wVInfection
                                        .Value = OAinfection
                                        .NumberFormat = "0.00%"
                                        .HorizontalAlignment = xlHAlignCenter
                                        .VerticalAlignment = xlVAlignCenter
                                   End With
                              Else
                              
                              End If
                         
                         ElseIf IsNumeric(OAminCq) = False Then
                              Set wVMinCq = WorklistView.Cells(wVAccNumMatch, wVTargetSearch.Column)
                              With wVMinCq
                                   .Value = OAminCq.Value
                                   .HorizontalAlignment = xlHAlignCenter
                                   .VerticalAlignment = xlVAlignCenter
                              End With
                         End If
                    Else
                         MsgBox "Could not find target on Worklist View Worksheet."
                    End If
               End If
          End If
     Next i
     
     With WorklistView.Range(WorklistView.Columns(2), WorklistView.Columns(wVTargetlastCol))
          .Columns.AutoFit
     End With
     
'     Call OptimizeCode_End
End Sub
