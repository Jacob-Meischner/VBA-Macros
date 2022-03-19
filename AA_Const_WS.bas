Attribute VB_Name = "AA_Const_WS"
Public Property Get importInfoWS() As Worksheet
     Set importInfoWS = ThisWorkbook.Sheets("Import Patient Information")
End Property
Public Property Get OAdataWS() As Worksheet
     Set OAdataWS = ThisWorkbook.Sheets("OpenArray Raw Data")
End Property
Public Property Get WorklistView() As Worksheet
     Set WorklistView = ThisWorkbook.Sheets("Worklist View")
End Property
Public Property Get PullReruns() As Worksheet
     Set PullReruns = ThisWorkbook.Sheets("Reruns To Pull")
End Property
Public Property Get LigoExpWS() As Worksheet
     Set LigoExpWS = ThisWorkbook.Sheets("Ligo Exports")
End Property
Public Property Get variableStor() As Worksheet
     Set variableStor = ThisWorkbook.Sheets("Variable Storage")
End Property
Property Get ResultFilePath() As String
     ResultFilePath = "C:\Users\jacob\OneDrive\Desktop\Wound 3-5-22"
End Property
Property Get LigoExportsPath() As String
     LigoExportsPath = "C:\Users\jacob\OneDrive\Desktop\Wound 3-5-22"
End Property
