Attribute VB_Name = "AAAA_Const_AMR_Ranges"
     'ACC.value, ampC.value, BILlatCMY.value, CTXpool.value, dfrA5A1.value, DHA.value, FOX.value, GES.value, IMPpool.value, KPC.value, mcr1.value, mecA.value, moxCMY.value, nfsA.value, OXApool.value,
     'OXA1.value, PER12.value, QnrASB.value, SHV.value, Sul12.value, TEM.value, tetBMS.value, vanA12B.value, VEB.value, VIM.value, AMR_Xeno.value
Function ampC() As Range
     Set ampC = variableStor.Range("B1:B38").Find("ampC", LookIn:=xlValues)
'     MsgBox ampC.Address
End Function
Function nfsA() As Range
     Set nfsA = variableStor.Range("B1:B38").Find("nfsA", LookIn:=xlValues)
'     MsgBox nfsA.Address
End Function
Function BILlatCMY() As Range
     Set BILlatCMY = variableStor.Range("B1:B38").Find("BIL/LAT/CMY", LookIn:=xlValues)
'     MsgBox BILlatCMY.Address
End Function
Function ACC() As Range
     Set ACC = variableStor.Range("B1:B38").Find("ACC", LookIn:=xlValues)
'     MsgBox ACC.Address
End Function
Function moxCMY() As Range
     Set moxCMY = variableStor.Range("B1:B38").Find("MOX/CMY", LookIn:=xlValues)
'     MsgBox moxCMY.Address
End Function
Function CTXpool() As Range
     Set CTXpool = variableStor.Range("B1:B38").Find("CTX pool", LookIn:=xlValues)
'     MsgBox CTXpool.Address
End Function
Function IMPpool() As Range
     Set IMPpool = variableStor.Range("B1:B38").Find("IMP pool", LookIn:=xlValues)
'     MsgBox IMPpool.Address
End Function
Function OXA1() As Range
     Set OXA1 = variableStor.Range("B1:B38").Find("OXA1", LookIn:=xlValues)
'     MsgBox OXA1.Address
End Function
Function PER12() As Range
     Set PER12 = variableStor.Range("B1:B38").Find("PER1/PER2", LookIn:=xlValues)
'     MsgBox PER12.Address
End Function
Function SHV() As Range
     Set SHV = variableStor.Range("B1:B38").Find("SHV", LookIn:=xlValues)
'     MsgBox SHV.Address
End Function
Function DHA() As Range
     Set DHA = variableStor.Range("B1:B38").Find("DHA", LookIn:=xlValues)
'     MsgBox DHA.Address
End Function
Function FOX() As Range
     Set FOX = variableStor.Range("B1:B38").Find("FOX", LookIn:=xlValues)
'     MsgBox FOX.Address
End Function
Function GES() As Range
     Set GES = variableStor.Range("B1:B38").Find("GES", LookIn:=xlValues)
'     MsgBox GES.Address
End Function
Function KPC() As Range
     Set KPC = variableStor.Range("B1:B38").Find("KPC", LookIn:=xlValues)
'     MsgBox KPC.Address
End Function
Function mcr1() As Range
     Set mcr1 = variableStor.Range("B1:B38").Find("mcr-1", LookIn:=xlValues)
'     MsgBox mcr1.Address
End Function
Function mecA() As Range
     Set mecA = variableStor.Range("B1:B38").Find("mecA", LookIn:=xlValues)
'     MsgBox mecA.Address
End Function
Function OXApool() As Range
     Set OXApool = variableStor.Range("B1:B38").Find("OXA pool", LookIn:=xlValues)
'     MsgBox OXApool.Address
End Function
Function QnrASB() As Range
     Set QnrASB = variableStor.Range("B1:B38").Find("QnrA/QnrS/QnrB", LookIn:=xlValues)
'     MsgBox QnrASB.Address
End Function
Function Sul12() As Range
     Set Sul12 = variableStor.Range("B1:B38").Find("Sul1/Sul2", LookIn:=xlValues)
'     MsgBox Sul12.Address
End Function
Function dfrA5A1() As Range
     Set dfrA5A1 = variableStor.Range("B1:B38").Find("dfrA5/dfrA1", LookIn:=xlValues)
'     MsgBox dfrA5A1.Address
End Function
Function TEM() As Range
     Set TEM = variableStor.Range("B1:B38").Find("TEM", LookIn:=xlValues)
'     MsgBox TEM.Address
End Function
Function tetBMS() As Range
     Set tetBMS = variableStor.Range("B1:B38").Find("tetB/tetM/tetS", LookIn:=xlValues)
'     MsgBox tetBMS.Address
End Function
Function vanA12B() As Range
     Set vanA12B = variableStor.Range("B1:B38").Find("vanA1/vanA2/vanB", LookIn:=xlValues)
'     MsgBox vanA12B.Address
End Function
Function VEB() As Range
     Set VEB = variableStor.Range("B1:B38").Find("VEB", LookIn:=xlValues)
'     MsgBox VEB.Address
End Function
Function VIM() As Range
     Set VIM = variableStor.Range("B1:B38").Find("VIM", LookIn:=xlValues)
'     MsgBox VIM.Address
End Function
Function AMR_Xeno() As Range
     Set AMR_Xeno = variableStor.Range("B1:B38").Find("AMR-Xeno", LookIn:=xlValues)
'     MsgBox AMR_Xeno.Address
End Function
