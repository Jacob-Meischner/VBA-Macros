VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WoundQuickFilterByID 
   Caption         =   "Control Comparison"
   ClientHeight    =   11985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "WoundQuickFilterByID.frx":0000
End
Attribute VB_Name = "WoundQuickFilterByID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function DlastRow() As Long
     With OAdataWS
          DlastRow = .Cells(.Rows.count, "D").End(xlUp).Row
     End With
End Function
Public Sub UserForm_Initialize()

     Call OptimizeCode_Begin

Dim tbl As ListObject, tblSN As Range
Dim SerialNum As Range, SerialNumArray As Variant, SerialNumCount As Long

    Me.Top = Application.Top + (Application.UsableHeight / 1) - (Me.Height / 1.35)         'loads userform to left side of screen where excel is located - MOVE LEFT - Number goes towards 0
    Me.Left = Application.Left + (Application.UsableWidth / 1) - (Me.Width / 0.85)     'adjusting me.height & me.width manipulates height + width. - MOVE LEFT - Towards 0

     With LB_serialNum
          .Clear
     End With

     
     Set tbl = OAdataWS.ListObjects("Table1")
          tbl.AutoFilter.ShowAllData                                                                     'unfilter data in table
     Set tblSN = tbl.ListColumns(10).DataBodyRange.SpecialCells(xlCellTypeVisible)     'entire filtered range of Column 1 of Table (Col D - Sample Name)
          
          SerialNumArray = tblSN                                      'set entire serial number range = array
          SerialNumArray = getUniques(SerialNumArray)       'send getUniques function to remove duplicates - set SerialNumArray = to what gets returned

          With LB_serialNum
               .List = SerialNumArray
          End With

'          With LB_Expected_Detected
'               .ColumnCount = 2
'               .ColumnWidths = "155"
'          End With
'
'          For i = 1 To (conActualCount - 1)
'               LB_Expected_Detected.AddItem
'               LB_Expected_Detected.List(i - 1, 1) = actualDetected(i)
'
'          Next i

     
Call OptimizeCode_End

End Sub
Private Sub LB_serialNum_Click()
     
     Call OptimizeCode_Begin

Dim tblRng As ListObject, filteredRng As Range, serialNumberRNG As Range, accNum As Range
Dim selectedSerialNum As String
Dim controlCount As Integer, ControlsArray As Variant, isFiltered As Boolean
Dim controlsArr As Variant, controlMatch As Variant, filteredMatch As Variant

     controlsArr = Array(pathPTC, pathPEC, pathNTC, pathNEC, amrPTC, amrPEC, amrNTC, amrNEC)

     ReDim ControlsArray(1 To 10) As Variant
     controlCount = 1
     
     Set tblRng = OAdataWS.ListObjects("Table1")
     
          isFiltered = TestFiltered
          If isFiltered = True Then
                    'clear all existing filters in table
               tblRng.AutoFilter.ShowAllData
                    'de-select everything in ControlsListBox
               With ControlsListBox
                    .ListIndex = -1
               End With
                    'clear all listboxes
               With LB_Expected_Detected
                    .Clear
               End With
               With LB_Actual_Detected
                    .Clear
               End With
                    'clear all captions
               With LBL_Expected_Total
                    .Caption = ""
               End With
               With LBL_Actual_Total
                    .Caption = ""
               End With
          Else

          End If
          
          If LB_serialNum.Value <> vbNullString Then
               selectedSerialNum = LB_serialNum.Value
               
               With OAdataWS
                    .ListObjects("Table1").Range.AutoFilter Field:=10, Criteria1:=selectedSerialNum
               End With
               
               Set filteredRng = tblRng.ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeVisible)     'entire filtered range of Column 1 of Table (Col D - Sample Name)
               
               For Each controlMatch In controlsArr                                            'for each control in controlsArr
                    filteredMatch = Application.match(controlMatch, filteredRng, 0)  'find a match in the filtered table range
                    If Not IsError(filteredMatch) Then                                         'if a match is found then
                         Set accNum = filteredRng.Cells(filteredMatch, 1)                 'set accNum = the match found
                         Set serialNumberRNG = accNum.Offset(0, 9)                        'offset 9 columns within filtered table range and set serial number
                         
                         If serialNumberRNG.Value = selectedSerialNum Then                'if serialNumberRNG = serial number selected in LB_serialNum then
                              ControlsArray(controlCount) = accNum.Value                  'add this accNum to array
                              controlCount = controlCount + 1                                  'add 1 to counter
                         Else
                              MsgBox "CRITICAL: Issue filtering data correctly by serial number.", vbCritical      'if serialNumberRNG <> selectedSerialNum then there's a major issue with the filtering and I need contacted"
                         End If
                    End If
               Next controlMatch
          End If
          
          If controlCount > 1 Then
               ReDim Preserve ControlsArray(1 To controlCount - 1)
               With ControlsListBox
                    .List = ControlsArray
               End With
          Else
               With ControlsListBox
                    .Clear
               End With
          End If
     
     Call OptimizeCode_End
     
End Sub
Private Sub ControlsListBox_Click()

     Call OptimizeCode_Begin
     
Dim selectedControl As String, filteredRng As Range

     'expected
Dim expectedCount As Long

     If ControlsListBox.Value <> vbNullString Then
          selectedControl = ControlsListBox.Value
          
          With OAdataWS
               .ListObjects("Table1").Range.AutoFilter Field:=1, Criteria1:=selectedControl
          End With
          
               'PATHOGEN START
          If selectedControl = pathPTC Then
               ActualControlLB (pathPTC)
               ExpectedControlLB (pathPTC)
          
          ElseIf selectedControl = pathPEC Then
               ActualControlLB (pathPEC)
               ExpectedControlLB (pathPEC)
               
          ElseIf selectedControl = pathNEC Then
               ActualControlLB (pathNEC)
               ExpectedControlLB (pathNEC)
               
          ElseIf selectedControl = pathNTC Then
               ActualControlLB (pathNTC)
               ExpectedControlLB (pathNTC)
          
               'AMR START
          ElseIf selectedControl = amrPTC Then
               ActualControlLB (amrPTC)
               ExpectedControlLB (amrPTC)
               
          ElseIf selectedControl = amrPEC Then
               ActualControlLB (amrPEC)
               ExpectedControlLB (amrPEC)
               
          ElseIf selectedControl = amrNEC Then
               ActualControlLB (amrNEC)
               ExpectedControlLB (amrNEC)
               
          ElseIf selectedControl = amrNTC Then
               ActualControlLB (amrNTC)
               ExpectedControlLB (amrNTC)
          
          End If
     End If
     
     If LBL_Expected_Total = LBL_Actual_Total Then
          With LB_Actual_Detected
               .ForeColor = RGB(0, 150, 0)
          End With
     Else
          With LB_Actual_Detected
               .ForeColor = RGB(150, 0, 0)
          End With
     End If
     Call OptimizeCode_End

End Sub
Private Sub Clear_Filter_Click()

Dim tblVar As ListObject: Set tblVar = OAdataWS.ListObjects("Table1")

   tblVar.AutoFilter.ShowAllData
     
     With LB_serialNum
          .ListIndex = -1
     End With
     With ControlsListBox
          .ListIndex = -1
          .Clear
     End With
     With LB_Expected_Detected
          .Clear
     End With
     With LB_Actual_Detected
          .Clear
     End With
     With LBL_Expected_Total
          .Caption = ""
     End With
     With LBL_Actual_Total
          .Caption = ""
     End With
     With OAdataWS.Range("A1")
          .Activate
     End With
End Sub
Public Function TestFiltered() As Boolean

Dim filterArea As Range
Dim rowsCount As Long, cellsCount As Long, columnsCount As Long

     Set filterArea = OAdataWS.ListObjects("Table1").Range

     rowsCount = filterArea.Rows.count
     columnsCount = filterArea.Columns.count

     cellsCount = filterArea.SpecialCells(xlCellTypeVisible).count

     If (rowsCount * columnsCount) > cellsCount Then
          TestFiltered = True
     Else
          TestFiltered = False
     End If

End Function

Private Sub RerunPatient_Click()

Dim accNum As String, target As String

accNum = Selection
target = Selection.Offset(0, 1)

     rerunAccNum accNum, target

End Sub
