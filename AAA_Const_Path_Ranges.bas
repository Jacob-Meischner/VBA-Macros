Attribute VB_Name = "AAA_Const_Path_Ranges"
     'Aci_bau, Bact_frag, Bact_vulg, Cit_fre, Clost_perf, Clost_sept, Coryn_str, Entero_aero, Entero_cloac, E_faecalis, E_faecium, E_coli, F_magna, Kleb_pneu, Kleb_oxy, Pept_ana, Pept_asa, Ana_prev, Prev_bivia, Prev_loe, Pro_mir, Pro_vul, Pseud_aer, Sal_M, Sal_N, Serr_marc,
     'Staph_aur, Staph_epid, Staph_haem, Staph_lugd, Staph_sapro, Strep_agalac, Strep_pneu, Strep_pyo, Path_Xeno, Can_albi
Function Aci_bau() As Range
     Set Aci_bau = variableStor.Range("D1:D38").Find("Acinetobacter baumannii", LookIn:=xlValues)
'     MsgBox Aci_bau.Value
End Function
Function Ana_prev() As Range
     Set Ana_prev = variableStor.Range("D1:D38").Find("Anaerococcus prevotii", LookIn:=xlValues)
'     MsgBox Ana_prev.Value
End Function
Function Bact_frag() As Range
     Set Bact_frag = variableStor.Range("D1:D38").Find("Bacteroides fragilis", LookIn:=xlValues)
'     MsgBox Bact_frag.Value
End Function
Function Bact_vulg() As Range
     Set Bact_vulg = variableStor.Range("D1:D38").Find("Bacteroides vulgatus", LookIn:=xlValues)
'     MsgBox Bact_vulg.Value
End Function
Function Can_albi() As Range
     Set Can_albi = variableStor.Range("D1:D38").Find("Candida albicans", LookIn:=xlValues)
'     MsgBox Can_albi.Value
End Function
Function Cit_fre() As Range
     Set Cit_fre = variableStor.Range("D1:D38").Find("Citrobacter freundii", LookIn:=xlValues)
'     MsgBox Cit_fre.Value
End Function
Function Clost_perf() As Range
     Set Clost_perf = variableStor.Range("D1:D38").Find("Clostridium perfringens", LookIn:=xlValues)
'     MsgBox Clost_perf.Value
End Function
Function Clost_sept() As Range
     Set Clost_sept = variableStor.Range("D1:D38").Find("Clostridium septicum", LookIn:=xlValues)
'     MsgBox Clost_sept.Value
End Function
Function Coryn_str() As Range
     Set Coryn_str = variableStor.Range("D1:D38").Find("Corynebacterium striatum", LookIn:=xlValues)
'     MsgBox Coryn_str.Value
End Function
Function Entero_aero() As Range
     Set Entero_aero = variableStor.Range("D1:D38").Find("Enterobacter aerogenes", LookIn:=xlValues)
'     MsgBox Entero_aero
End Function
Function Entero_cloac() As Range
     Set Entero_cloac = variableStor.Range("D1:D38").Find("Enterobacter cloacae", LookIn:=xlValues)
'     MsgBox Entero_cloac.Value
End Function
Function E_faecalis() As Range
     Set E_faecalis = variableStor.Range("D1:D38").Find("Enterococcus faecalis", LookIn:=xlValues)
'     MsgBox E_faecalis.Value
End Function
Function E_faecium() As Range
     Set E_faecium = variableStor.Range("D1:D38").Find("Enterococcus faecium", LookIn:=xlValues)
'     MsgBox E_faecium.Value
End Function
Function E_coli() As Range
     Set E_coli = variableStor.Range("D1:D38").Find("Escherichia coli", LookIn:=xlValues)
'     MsgBox E_coli.Value
End Function
Function F_magna() As Range
     Set F_magna = variableStor.Range("D1:D38").Find("Finegoldia magna", LookIn:=xlValues)
'     MsgBox F_magna.Value
End Function
Function Kleb_oxy() As Range
     Set Kleb_oxy = variableStor.Range("D1:D38").Find("Klebsiella oxytoca", LookIn:=xlValues)
'     MsgBox Kleb_oxy.Value
End Function
Function Kleb_pneu() As Range
     Set Kleb_pneu = variableStor.Range("D1:D38").Find("Klebsiella pneumoniae", LookIn:=xlValues)
'     MsgBox Kleb_pneu.Value
End Function
Function Pept_asa() As Range
     Set Pept_asa = variableStor.Range("D1:D38").Find("Peptoniphilus asaccharolyticus", LookIn:=xlValues)
'     MsgBox Pept_asa.Value
End Function
Function Pept_ana() As Range
     Set Pept_ana = variableStor.Range("D1:D38").Find("Peptostreptococcus anaerobius", LookIn:=xlValues)
'     MsgBox Pept_ana.Value
End Function
Function Prev_bivia() As Range
     Set Prev_bivia = variableStor.Range("D1:D38").Find("Prevotella bivia", LookIn:=xlValues)
'     MsgBox Prev_bivia.Value
End Function
Function Prev_loe() As Range
     Set Prev_loe = variableStor.Range("D1:D38").Find("Prevotella loescheii", LookIn:=xlValues)
'     MsgBox Prev_loe.Value
End Function
Function Pro_mir() As Range
     Set Pro_mir = variableStor.Range("D1:D38").Find("Proteus mirabilis", LookIn:=xlValues)
'     MsgBox Pro_mir.Value
End Function
Function Pro_vul() As Range
     Set Pro_vul = variableStor.Range("D1:D38").Find("Proteus vulgaris", LookIn:=xlValues)
'     MsgBox Pro_vul.Value
End Function
Function Pseud_aer() As Range
     Set Pseud_aer = variableStor.Range("D1:D38").Find("Pseudomonas aeruginosa", LookIn:=xlValues)
'     MsgBox Pseud_aer.Value
End Function
Function Sal_M() As Range
     Set Sal_M = variableStor.Range("D1:D38").Find("Salmonella enterica Serovar Montevideo", LookIn:=xlValues)
'     MsgBox Sal_M.Value
End Function
Function Sal_N() As Range
     Set Sal_N = variableStor.Range("D1:D38").Find("Salmonella enterica Serovar Newport", LookIn:=xlValues)
'     MsgBox Sal_N.Value
End Function
Function Serr_marc() As Range
     Set Serr_marc = variableStor.Range("D1:D38").Find("Serratia marcescens", LookIn:=xlValues)
'     MsgBox Serr_marc.Value
End Function
Function Staph_aur() As Range
     Set Staph_aur = variableStor.Range("D1:D38").Find("Staphylococcus aureus", LookIn:=xlValues)
'     MsgBox Staph_aur.Value
End Function
Function Staph_epid() As Range
     Set Staph_epid = variableStor.Range("D1:D38").Find("Staphylococcus epidermidis", LookIn:=xlValues)
'     MsgBox Staph_epid.Value
End Function
Function Staph_haem() As Range
     Set Staph_haem = variableStor.Range("D1:D38").Find("Staphylococcus haemolyticus", LookIn:=xlValues)
'     MsgBox Staph_haem.Value
End Function
Function Staph_lugd() As Range
     Set Staph_lugd = variableStor.Range("D1:D38").Find("Staphylococcus lugdunensis", LookIn:=xlValues)
'     MsgBox Staph_lugd.Value
End Function
Function Staph_sapro() As Range
     Set Staph_sapro = variableStor.Range("D1:D38").Find("Staphylococcus saprophyticus", LookIn:=xlValues)
'     MsgBox Staph_sapro.Value
End Function
Function Strep_agalac() As Range
     Set Strep_agalac = variableStor.Range("D1:D38").Find("Streptococcus agalactiae", LookIn:=xlValues)
'     MsgBox Strep_agalac.Value
End Function
Function Strep_pneu() As Range
     Set Strep_pneu = variableStor.Range("D1:D38").Find("Streptococcus pneumoniae", LookIn:=xlValues)
'     MsgBox Strep_pneu.Value
End Function
Function Strep_pyo() As Range
     Set Strep_pyo = variableStor.Range("D1:D38").Find("Streptococcus pyogenes", LookIn:=xlValues)
'     MsgBox Strep_pyo.Value
End Function
Function Path_Xeno() As Range
     Set Path_Xeno = variableStor.Range("D1:D38").Find("Path-Xeno", LookIn:=xlValues)
'     MsgBox Path_Xeno.Value
End Function
