Attribute VB_Name = "F_Change_Target_Names"
Sub Change_PathogenNames()
          'DATE ADDED 3-5-22
          
Dim pathName As Range, pathNameChange As Variant
Dim eLastRow As Long, ClastRow As Long
Dim PreConvertedNames As Range

     With OAdataWS
          eLastRow = .Cells(.Rows.count, "E").End(xlUp).Row
     End With
     
     With variableStor
          ClastRow = .Cells(.Rows.count, "C").End(xlUp).Row
          Set PreConvertedNames = .Range("C1:C" & ClastRow)
     End With
     
     For Each pathName In OAdataWS.Range("E11:E" & eLastRow).Cells
          pathNameChange = Application.match(pathName.Value, PreConvertedNames, 0)
          If Not IsError(pathNameChange) Then
               With pathName
                    .Value = variableStor.Cells(pathNameChange, 4)
               End With
          End If
          If pathName.Offset(0, -1).Value = PTC Then
               With pathName.Offset(0, -1)
                    .Value = pathPTC
               End With
          ElseIf pathName.Offset(0, -1).Value = PEC Then
               With pathName.Offset(0, -1)
                    .Value = pathPEC
               End With
          ElseIf pathName.Offset(0, -1).Value = NEC Then
               With pathName.Offset(0, -1)
                    .Value = pathNEC
               End With
          ElseIf pathName.Offset(0, -1).Value = NTC Then
               With pathName.Offset(0, -1)
                    .Value = pathNTC
               End With
          End If
     Next pathName
     
End Sub
Sub Change_AMRNames()

Dim amrName As Range, amrNameChange As Variant
Dim AlastRow As Long, eLastRow As Long
Dim amrNameChangeCol As Range

     With OAdataWS
          eLastRow = .Cells(.Rows.count, "E").End(xlUp).Row
     End With
     
     With variableStor
          AlastRow = .Cells(.Rows.count, "A").End(xlUp).Row
          Set amrNameChangeCol = .Range("A1:A" & AlastRow)
     End With

     For Each amrName In OAdataWS.Range("E11:E" & eLastRow).Cells
          amrNameChange = Application.match(amrName.Value, amrNameChangeCol, 0)
          If Not IsError(amrNameChange) Then
               With amrName
                    .Value = variableStor.Cells(amrNameChange, 2)
               End With
          End If
          If amrName.Offset(0, -1).Value = PTC Then
               With amrName.Offset(0, -1)
                    .Value = amrPTC
               End With
          ElseIf amrName.Offset(0, -1).Value = PEC Then
               With amrName.Offset(0, -1)
                    .Value = amrPEC
               End With
          ElseIf amrName.Offset(0, -1).Value = NEC Then
               With amrName.Offset(0, -1)
                    .Value = amrNEC
               End With
          ElseIf amrName.Offset(0, -1).Value = NTC Then
               With amrName.Offset(0, -1)
                    .Value = amrNTC
               End With
          End If
     Next amrName

End Sub
