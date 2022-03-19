Attribute VB_Name = "H_Create_Table_Update_Pivot"
Sub createTable()

'     Call OptimizeCode_Begin

Dim DlastRow As Long, tableRNG As Range

     With OAdataWS
          DlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
          Set tableRNG = OAdataWS.Range("D10:M" & DlastRow)
          .ListObjects.Add(xlSrcRange, tableRNG, , xlYes).Name = "Table1"
     End With

'     With ThisWorkbook
'          .RefreshAll
'     End With

'Call OptimizeCode_End

End Sub
