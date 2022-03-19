Attribute VB_Name = "D_Full_Quant_Results"
Sub Import_Full_Quant_Results()
          'DATE ADDED 3-5-22

Call OptimizeCode_Begin

Dim ST As Single
ST = timer

     isExit = False      'everytime the Import QS Raw Data is clicked - set this to false, if there are errors in other modules, then this will be set to true
     
     If Dir(ResultFilePath, vbDirectory) <> "" Then
          ChDir ResultFilePath
     Else
          ChDrive "C"
     End If

          Call Import_Result_Files                     'import sorted data
               If isExit = True Then Exit Sub
          
          Call createTable                             'adds Table1 to OpenArray Raw Data and updates pivot table
          
          Call Prepare_Worklist                         'prepare the worklist / clear ranges & load accession to worklist view - post file import

          Call Full_Quant_Interpretation          'Full Interpretation - Min Cq, Full Quant, Infection %
               If isExit = True Then Exit Sub

          Call setWorklistViewValues              'returns Min Cq, Full Quant Result, and Infection % values for all patients to the worklist view WS


     With OAdataWS.OLEObjects("Control_Comparison")
          If .Height <> 50.25 Then
               .Height = 50.25
          Else
               .Height = 51
          End If
     End With

     With OAdataWS.OLEObjects("Create_Transfer_Ligo_File")
          If .Height <> 50.25 Then
               .Height = 50.25
          Else
               .Height = 51
          End If
     End With
     
     OAdataWS.Activate
'     WoundQuickFilterByID.UserForm_Initialize
'     WoundQuickFilterByID.Show (0)
     

             MsgBox "Macro took: " & timer - ST & " seconds to complete."
Call OptimizeCode_End
     

End Sub
