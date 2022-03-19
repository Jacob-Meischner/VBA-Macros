Attribute VB_Name = "K_Ligo_Result_File"
Sub Create_Ligo_Exports_File()
     'DATE ADDED - 3/14/22
     
Call OptimizeCode_Begin

Dim OAdataDlastRow As Long, i As Long, j As Long
Dim oaDataaccNum As Range, oaDatatargName As Range, oaDataMinCq As Range, oaDataFullQuant As Range, oaDataInf As Range
Dim combinedRNG As Variant, destRng As Range, ligoWSdlastRow As Long

     With OAdataWS
          OAdataDlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
     End With
     
     With LigoExpWS
          .Range("A1:T6000").Clear
          .Range("A1").Value = "Well"
          .Range("B1").Value = "Wound"
          .Range("D1").Value = "Sample"
          .Range("E1").Value = "Target"
          .Range("M1").Value = "Min Cq"
          .Range("R1").Value = "Full Quantitative Result"
          .Range("S1").Value = "Infection %"
          ligoWSdlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
     End With
     
     Set destRng = LigoExpWS.Range("D" & ligoWSdlastRow).Offset(1, 0)
     
     ReDim combinedRNG(1 To 5)

     For i = 11 To OAdataDlastRow Step 3
          Set oaDataaccNum = OAdataWS.Cells(i, 4)
          Set oaDatatargName = OAdataWS.Cells(i, 5)
          Set oaDataMinCq = OAdataWS.Cells(i, 10)
          Set oaDataFullQuant = OAdataWS.Cells(i, 11)
          Set oaDataInf = OAdataWS.Cells(i, 12)
          
          combinedRNG = Array(oaDataaccNum.Value, oaDatatargName.Value, oaDataMinCq.Value, oaDataFullQuant.Value, oaDataInf.Value)
          
          For j = LBound(combinedRNG, 1) To UBound(combinedRNG, 1)
               If Not IsEmpty(combinedRNG(j)) Then
                    With destRng
                         .Value = combinedRNG(j)
                    End With
               Else
                    'do nothing
               End If
               
               If destRng.Column = 5 Then              'target column - E
                    Set destRng = destRng.Offset(0, 8)      'Min Cq col - M
               ElseIf destRng.Column = 13 Then
                    Set destRng = destRng.Offset(0, 5)
               ElseIf destRng.Column = 19 Then
                    Set destRng = destRng.Offset(1, -15)     'Infection % col - O
               Else
                    Set destRng = destRng.Offset(0, 1)
               End If
          Next j
     Next i

Dim postImplastRow As Long

With LigoExpWS
     postImplastRow = .Cells(.Rows.count, "D").End(xlUp).Row
End With

     With LigoExpWS
          With .Range("M2:M" & postImplastRow)
               .NumberFormat = "0.000"       '3 decimal places formatting to ranges
          End With
          With .Range("R2:R" & postImplastRow)
               .NumberFormat = "0.00E+00"    'scientific notation formatting
          End With
          With .Range("S2:S" & postImplastRow)
               .NumberFormat = "0.00%"       'percent formatting
          End With
          .Range("A:T").Columns.AutoFit
     End With

Call OptimizeCode_End

End Sub
