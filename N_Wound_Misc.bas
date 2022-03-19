Attribute VB_Name = "N_Wound_Misc"
Sub OptimizeCode_Begin()
     Application.ScreenUpdating = False
     Application.EnableEvents = False
     Application.Calculation = xlCalculationManual
End Sub
Sub OptimizeCode_End()
     Application.ScreenUpdating = True
     Application.EnableEvents = True
     Application.Calculation = xlCalculationAutomatic
End Sub
Sub WoundOpenRerunSheet()

Call OptimizeCode_Begin
Dim Ret, checkifOpen As Workbook

ChDrive "X"

    Set checkifOpen = Workbooks.Open(rrFilePath)
    Ret = IsWorkBookOpenNow(rrFilePath)
        
        isExit = False
        
        If checkifOpen.ReadOnly = True Then
            MsgBox ("Rerun file is currently opened by another user." & vbNewLine & vbNewLine & "Have the other user close the file and try again.")
                isExit = True
            Exit Sub
        ElseIf Ret = True Then
            isExit = False
        Else
            Workbooks.Open (rrFilePath)
        End If
Call OptimizeCode_End

End Sub
Sub WoundRerunSheetSave()
        With Workbooks(rrFileName)
            .Save
        End With
End Sub
