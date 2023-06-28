Attribute VB_Name = "kozos"
'Vakablak és Szerkesztõ v2

'Fájlcsatornák:
'               Projektor - 1
'               Szerkesztõ - 2
'               Beállítások - 3
'               Nyelv   - 4
'               Ponthatárok - 5



'Közös, megosztott elérésû modul
'Az aktív módot meghatározó objektum: Szulo



Option Explicit

Public Vakterkep As Program
Public Szulo As Object
'Public Szovegek(1 To 100) As String

'Hibatár
    
Type Program
    Verzio As String
    Konyvtar As String
End Type

Public Type KerdesValasz
    Kerdes As String
    Valasz As String
End Type

Public Type elem
    Cime As String
    tipp As String
    Kovetkezo As Byte
    kerdesek(1 To 256) As KerdesValasz
    Bal As String
    Felso As String
    pont As Double
    Kitalalta As Boolean
End Type

Public KozosSzovegek(1 To 50) As String


Sub Main()
On Error GoTo Hiba
    'If True Then debuger.Show
    Vakterkep.Verzio = "Vakablak " & App.Major & "." & App.Minor ' & " b" & App.Revision
    If Len(App.Path) = 3 Then
            Vakterkep.Konyvtar = App.Path
        Else
            Vakterkep.Konyvtar = App.Path & "\"
    End If
    
    Select Case LCase(Utasitas(Command$))
        Case "-szerkeszto", "-sz", "-s"
            Set Szulo = szerkeszto
            magyar_nyelv
            szerkeszto.megnyitas ("$vt\vakablak.ini")
            szerkeszto.Show
            szerkeszto.torol
            If Ertek(Command$) <> "" Then
                    szerkeszto.megnyitas (Ertek(Command$))
                Else
                    'tulajdonsagok.megse.Enabled = False
                    tulajdonsagok.Mutat (0)
            End If
            
        Case "-beall"
            'Beállítások nélkül
            Terkepeles
            terkep.beall_mnu.Enabled = True
        Case ""
            Terkepeles
        Case Else
            Terkepeles
            terkep.megnyitas (Command$)
            Exit Sub
    End Select
    'Szulo.Show
Exit Sub
Hiba:
    Select Case Err.Number
        Case 339
            MsgBox "Your system does not support ActiveX UserControls. Please reinstall", vbCritical, "UserControl error"
            MsgBox Err.Description, vbCritical, "ERR: " & Err.Number
            Megsemmisit
        Case 374
            MsgBox "ActiveX compatibility error.", vbCritical, "UserControl error"
            MsgBox Err.Description, vbCritical, "ERR: " & Err.Number
            Megsemmisit
        Case Else
            MsgBox Err.Description
    End Select
End Sub

Public Function Konyvtara(Fajlnev As String)
    Dim i As Integer, j As Integer
    j = 0
    For i = 1 To Len(Fajlnev)
        If Mid(Fajlnev, i, 1) = "\" Then j = i
    Next i
    Konyvtara = Mid(Fajlnev, 1, j)
End Function

Public Function Atalakit(Adat As String, Optional egyeb As Variant)
    Dim i As Integer, uj As String
    uj = ""
    i = 1
    Do While i <= Len(Adat)
        If Mid(Adat, i, 3) = "$vt" Then
            uj = uj & Mid(Vakterkep.Konyvtar, 1, Len(Vakterkep.Konyvtar) - 1)
            i = i + 3
        End If
        
        If Mid(Adat, i, 1) = "\" And i = 1 Then
            uj = uj & egyeb
            i = i + 1
        End If
        
        If Mid(Adat, i, 3) = "$sz" Then
            uj = uj & egyeb
            i = i + 3
        End If
        
        If Mid(Adat, i, 6) = "$jelsz" Then
            uj = uj & terkep.jel(egyeb).Width
            i = i + 6
        End If
        
        If Mid(Adat, i, 5) = "$jelm" Then
            uj = uj & terkep.jel(egyeb).Height
            i = i + 5
        End If
        
        If Mid(Adat, i, 6) = "$szovm" Then
            uj = uj & terkep.jel_szoveg(egyeb).Height
            i = i + 6
        End If
        
        If Mid(Adat, i, 7) = "$szovsz" Then
            uj = uj & terkep.jel_szoveg(egyeb).Width
            i = i + 7
        End If
        
        If Mid(Adat, i, 6) = "$<br>$" Then
            uj = uj & vbCrLf
            i = i + 6
        End If
        
        uj = uj & Mid(Adat, i, 1)
        i = i + 1
    Loop
    Atalakit = uj
End Function

Public Function Nagyobb(Egyik As Single, Masik As Single) As Single
    If Egyik > Masik Then
                    Nagyobb = Egyik
                Else
                    Nagyobb = Masik
            End If
End Function
Public Sub megjelol(Index As Integer)
Szulo.tabulalo = Index
On Error Resume Next
    Dim i As Integer
    For i = 1 To Szulo.jel_szoveg.Count - 1
        Szulo.jel_szoveg(i).BorderStyle = 0
    Next i
    Szulo.jel_szoveg(Index).BorderStyle = 1
    If Szulo.Name = "szerkeszto" Then
        szerkeszto.Passzint Index
    End If
End Sub
Public Function V(Kifejezes As Integer) As Byte
    V = Abs(Int(Kifejezes))
End Function
Public Sub Megsemmisit()
    Unload beallitasok
    Unload jelmagyarazat
    Unload nevjegy
    Unload ertekeles
    Unload nyomtatas
    Unload tulajdonsagok
    Unload szerkeszto
    Unload terkep
    Set Szulo = Nothing
    End
End Sub

'Public Function uzenet(Index As Integer, Beszurando As String) As String
 '   If InStr(1, Szovegek(Index), "$sz") Then
 '           uzenet = Atalakit(Szovegek(Index), Beszurando)
 '       Else
 '           uzenet = Szovegek(Index)
 '   End If
'End Function

Public Function CsakANeve(Eleres As String)
    CsakANeve = Mid(Eleres, Len(Konyvtara(Eleres)) + 1, Len(Eleres))
End Function
Public Function aze() As Boolean
On Error GoTo nincsis
    If TypeOf Szulo.obj Is jel Then
        aze = True
    Else
        aze = False
    End If
Exit Function
nincsis:
    aze = False
End Function
Public Function RelativEleres(Mappa As String, FajlElerese As String) As String
    If LCase(Mid(Konyvtara(FajlElerese), 1, Len(Mappa))) = LCase(Mappa) Then
        RelativEleres = Mid(FajlElerese, Len(Mappa) + 1, Len(FajlElerese))
    Else
        RelativEleres = ""
    End If
End Function
Public Sub HHSugo(Melyiket As String)
On Error GoTo Hiba
    If Szulo.Name = "terkep" Then
            Melyiket = "vakablak.chm::/html/" & Melyiket
        Else
            Melyiket = "szerkeszto.chm::/html/" & Melyiket
    End If
    Shell "hh.exe " & Vakterkep.Konyvtar & Melyiket, vbNormalFocus
Exit Sub
Hiba:
    MsgBox KozosSzovegek(36), vbCritical, KozosSzovegek(37)
End Sub
Public Sub Pozicional()
On Error Resume Next
Dim X As Single, Y As Single, pw As Single, ph As Single

If Szulo.fl.Visible Then
        pw = Szulo.fl.Width
    Else
        pw = 0
End If

If Szulo.jb.Visible Then
        ph = Szulo.jb.Height
    Else
        ph = 0
End If


X = (Szulo.ScaleWidth - Szulo.terulet.Width - pw) / 2
Y = (Szulo.ScaleHeight - Szulo.terulet.Height - ph) / 2


Szulo.fl.Move Szulo.ScaleWidth - Szulo.fl.Width, 0, Szulo.fl.Width, Szulo.ScaleHeight - Szulo.fl.Width
Szulo.jb.Move 0, Szulo.ScaleHeight - Szulo.jb.Height, Szulo.ScaleWidth - Szulo.jb.Height, Szulo.jb.Height
Szulo.terulet.Move X, Y

If Szulo.ScaleWidth - Szulo.terulet.Width - pw < 0 Then
        Szulo.jb.SmallChange = Int(Szulo.ScaleWidth - Szulo.terulet.Width / 100)
        Szulo.jb.LargeChange = Int(Szulo.ScaleWidth - Szulo.terulet.Width / 10)
        Szulo.jb.Max = Szulo.ScaleWidth - Szulo.terulet.Width - 2 * Szulo.fl.Width
        Szulo.jb.Min = Szulo.fl.Width
        Szulo.jb.Visible = True
        Szulo.jb.Value = Szulo.terulet.Left
    Else
        Szulo.jb.Visible = False
End If

If Szulo.ScaleHeight - Szulo.terulet.Height - ph < 0 Then
        Szulo.fl.SmallChange = Int(Szulo.ScaleHeight - Szulo.terulet.Height / 100)
        Szulo.fl.LargeChange = Int(Szulo.ScaleHeight - Szulo.terulet.Height / 10)
        Szulo.fl.Max = Szulo.ScaleHeight - Szulo.terulet.Height - 2 * Szulo.jb.Height
        Szulo.fl.Min = Szulo.jb.Height
        Szulo.fl.Visible = True
        Szulo.fl.Value = Szulo.terulet.Top
    Else
        Szulo.fl.Visible = False
End If

If (Szulo.fl.Visible Or Szulo.jb.Visible) Then
        Szulo.gomb.Move Szulo.fl.Left, Szulo.jb.Top
        Szulo.gomb.Visible = True
        If Szulo.Name = "terkep" Then Szulo.gomb.Enabled = Szulo.jelm_mnu.Enabled
Else
        Szulo.gomb.Visible = False
        
End If
End Sub
Public Sub UzenetAblak(szoveg As String, Optional Cime As String)
    If Cime = "" Then Cime = KozosSzovegek(38)
    If szoveg <> "" Then MsgBox szoveg, vbExclamation, Cime
End Sub
Public Sub KozosHibak(Hiba As Integer)
    Select Case Hiba
        Case 70
            UzenetAblak KozosSzovegek(40), KozosSzovegek(39)
        Case 71
            UzenetAblak KozosSzovegek(41), KozosSzovegek(39)
        Case 75
            UzenetAblak KozosSzovegek(42), KozosSzovegek(39)
        Case Else
            UzenetAblak Err.Number & ": " & Err.Description, KozosSzovegek(15)
    End Select
    
End Sub
Public Sub NyelvAlkalmazasa(Nyelvfajl As String)
On Error GoTo Hiba

    Dim sor As String, parancs As String, Index As Integer
    Open Vakterkep.Konyvtar & "nyelvek\" & Nyelvfajl For Input As 4
    
    If Szulo.Name = "terkep" Then
            Do While Not EOF(4)
                
                Line Input #4, sor
                
                If sor = "" Or Mid(sor, 1, 1) = ";" Or Mid(sor, 1, 1) = "#" Or Mid(sor, 1, 1) = "/" Or Mid(sor, 1, 1) = "[" Then GoTo kihagy
                parancs = Utasitas(sor)
                
                Select Case parancs
                ' ######## TÉRKÉP OLVASÓ ABLAK
                    ' ###### FÁJL MENÜ
                    Case "fajl"
                        terkep.fajl_mnu.Caption = Ertek(sor)
                        
                    Case "fajl.ujra_kezd"
                        terkep.ujra_mnu.Caption = Ertek(sor)
                        
                    Case "fajl.megnyit"
                        terkep.megnyit_mnu.Caption = Ertek(sor)
                        
                    Case "fajl.nyomtat"
                        terkep.nyomtat.Caption = Ertek(sor)
                        'nyomtatas.Caption = Ertek(sor)
                        
                    Case "fajl.szerkeszt"
                        terkep.szerkesztes_mnu.Caption = Ertek(sor)
                        
                    Case "fajl.kilepes"
                        terkep.kilep_mnu.Caption = Ertek(sor)
                    
                    '##### ESZKÖZÖK MENÜ
                    Case "eszkozok"
                        terkep.eszkozok_mnu.Caption = Ertek(sor)
                        
                    Case "eszkozok.jelmagyarazat"
                        terkep.jelm_mnu.Caption = Ertek(sor)
                        jelmagyarazat.Caption = Ertek(sor)
                        
                    Case "eszkozok.ertekeles"
                        terkep.ertekel_mnu.Caption = Ertek(sor)
                        ertekeles.ertekeles = Ertek(sor)
                        
                    Case "eszkozok.beallitas"
                        terkep.beall_mnu.Caption = Ertek(sor)
                        beallitasok.Caption = Ertek(sor)
                    
                    Case "eszkozok.behuzo"
                        terkep.behuzo_mnu.Caption = Ertek(sor)
                        behuzo.Caption = Ertek(sor)
                        
                    '###### SÚGÓ MENÜ
                    Case "sugo"
                        terkep.sugo_mnu.Caption = Ertek(sor)
                    Case "sugo.segitseg"
                        terkep.help_mnu.Caption = Ertek(sor)
                        
                    Case "sugo.nevjegy"
                        terkep.nevjegy_mnu.Caption = Ertek(sor)
                    
                    'Névjegy
                    Case "nevjegy.forditas"
                        nevjegy.cimke(4).Caption = Ertek(sor)
                        
                    Case "nevjegy.fordito"
                        nevjegy.cimke(3).Caption = Ertek(sor)
                        
                    Case "nevjegy.koszonet"
                        terkep.koszonet_mnu.Caption = Ertek(sor)
                        
                    '#### GOMBOK ELNEVEZÉSE
                    Case "gombok.megse"
                        beallitasok.megse.Caption = Ertek(sor)
                        'tulajdonsagok.megse.Caption = Ertek(sor)
                        nyomtatas.megse.Caption = Ertek(sor)
                        
                    Case "gombok.ok"
                        beallitasok.ok.Caption = Ertek(sor)
                        'tulajdonsagok.ok.Caption = Ertek(sor)
                        koszonet.ok.Caption = Ertek(sor)
                        
                    Case "gombok.mentes"
                        beallitasok.alkalmaz.Caption = Ertek(sor)
                        
                    Case "gombok.nyomtat"
                        nyomtatas.ok.Caption = Ertek(sor)
                        
                    Case "gombok.sugo"
                        beallitasok.sugo.Caption = Ertek(sor)
                        'tulajdonsagok.sugo.Caption = Ertek(sor)
                        nyomtatas.sugo.Caption = Ertek(sor)
                        
                    Case "gombok.bezar"
                        nyomtatas.megse.Caption = Ertek(sor)
                '##### BEÁLLÍTÁSOK ABLAK
                    'ÁLTALÁNOS FÜL
                    Case "beallitasok.altalanos"
                        beallitasok.beall_mnu(0).Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.beall_enged"
                        beallitasok.enged.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.egyeni_enged"
                        beallitasok.egyeni.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.szerkesztes_enged"
                        beallitasok.szerkesztes.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.behuzas_enged"
                        beallitasok.behuzas.Caption = Ertek(sor)
                    
                    Case "beallitasok.altalanos.kerdesek_enged"
                        beallitasok.kerdesek.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.tippek_enged"
                        beallitasok.tippek.Caption = Ertek(sor)
                    
                    Case "beallitasok.altalanos.potlo"
                        beallitasok.cetli(5).Caption = Ertek(sor)
                    'Étékelés fül
                    
                    Case "beallitasok.ertekeles"
                        beallitasok.beall_mnu(1).Caption = Ertek(sor)
                        
                    Case "beallitasok.ertekeles.feltetel"
                        beallitasok.cetli(2).Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.akkor"
                        beallitasok.cetli(0).Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.megnevezes"
                        beallitasok.cetli(1).Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.felvesz"
                        beallitasok.felvesz.Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.torol"
                        beallitasok.torol.Caption = Ertek(sor)
                        
                        
                    Case "beallitasok.ertekeles.pont"
                        beallitasok.cetli(4).Caption = Ertek(sor)
                    
                    Case "beallitasok.ertekeles.levonas"
                        beallitasok.cetli(10).Caption = Ertek(sor)
                    'Ellenõrzés
                    Case "beallitasok.ellenorzes"
                        beallitasok.beall_mnu(2).Caption = Ertek(sor)
                        
                    Case "beallitasok.ellenorzes.kapitalis"
                        beallitasok.kisbetus.Caption = Ertek(sor)
                    
                    Case "beallitasok.ellenorzes.felesleges"
                        beallitasok.szokoz.Caption = Ertek(sor)
                        
                    Case "beallitasok.ellenorzes.kotojelek"
                        beallitasok.nincskotojel.Caption = Ertek(sor)
                    
                    Case "beallitasok.ellenorzes.szinek"
                        beallitasok.megoldas(0).Caption = Ertek(sor)
                    
                    Case "beallitasok.ellenorzes.eltuntet"
                        beallitasok.megoldas(1).Caption = Ertek(sor)
                'Nyelvi modul
                    Case "beallitasok.nyelv"
                        beallitasok.beall_mnu(3).Caption = Ertek(sor)
                    Case "beallitasok.nyelv.magyar"
                        beallitasok.nyelv(0).Caption = Ertek(sor)
                    Case "beallitasok.nyelv.idegen"
                        beallitasok.nyelv(1).Caption = Ertek(sor)
                    Case "beallitasok.nyelv.betolt"
                        beallitasok.betoltnyelv.Caption = Ertek(sor)
                'Nyomtatás
                    'Case "nyomtatas"
                    '    nyomtatas.Caption = Ertek(sor)
                        
                    Case "nyomtatas.szelesseg"
                        nyomtatas.cimke(1).Caption = Ertek(sor)
                    
                    Case "nyomtatas.magassag"
                        nyomtatas.cimke(2).Caption = Ertek(sor)
                'Értékelés ablak
                    Case "ertekeles.helyes"
                        ertekeles.cetli(1).Caption = Ertek(sor)
                    Case "ertekeles.hibas"
                        ertekeles.cetli(0).Caption = Ertek(sor)
                    Case "ertekeles.eredmeny"
                        ertekeles.cetli(2).Caption = Ertek(sor)
                    Case "ertekeles.pontok"
                        ertekeles.cetli(5).Caption = Ertek(sor)
                    Case "ertekeles.jegy"
                        ertekeles.cetli(4).Caption = Ertek(sor)
                End Select
                
                'Közös szövegek
                If IsNumeric(parancs) Then
                    Index = CByte(parancs)
                    KozosSzovegek(Index) = Ertek(sor)
                End If
kihagy:
            Loop
            
            'Nyelv Bejelölése
            Dim i As Long
            For i = 0 To beallitasok.nyelvek.ListCount - 1
                If beallitasok.nyelvek.List(i) = Nyelvfajl Then beallitasok.nyelvek.ListIndex = i
            Next i
        Else '########## Szerkesztõ
          Do While Not EOF(4)
                Line Input #4, sor
                If sor = "" Or Mid(sor, 1, 1) = ";" Or Mid(sor, 1, 1) = "#" Or Mid(sor, 1, 1) = "/" Or Mid(sor, 1, 1) = "[" Then GoTo kihagySZ
                
                parancs = Utasitas(sor)
                
                Select Case parancs
                    'Projekt mnu
                    Case "projekt"
                        szerkeszto.file.Caption = Ertek(sor)
                    Case "projekt.uj"
                        szerkeszto.uj_mnu.Caption = Ertek(sor)
                    Case "projekt.megnyit"
                        szerkeszto.megnyit_mnu.Caption = Ertek(sor)
                    Case "projekt.ment"
                        szerkeszto.ment_mnu.Caption = Ertek(sor)
                    Case "projekt.ment_mint"
                        szerkeszto.ment_mint_mnu.Caption = Ertek(sor)
                    Case "projekt.tulajdonsagok"
                        szerkeszto.olda_mnu.Caption = Ertek(sor)
                    Case "projekt.megtekint"
                        szerkeszto.megtekint_mnu.Caption = Ertek(sor)
                    Case "projekt.beallitasok"
                        szerkeszto.beall_mnu.Caption = Ertek(sor)
                    Case "projekt.kilep"
                        szerkeszto.kilepes_mnu.Caption = Ertek(sor)
                    'Súgó mnu
                    Case "sugo"
                        szerkeszto.sugo_mnu.Caption = Ertek(sor)
                    Case "sugo.segitseg"
                        szerkeszto.sugo_mnup.Caption = Ertek(sor)
                        
                    Case "sugo.nevjegy"
                        szerkeszto.nevjegy_mnu.Caption = Ertek(sor)
                    
                    'Névjegy
                    Case "nevjegy.forditas"
                        nevjegy.cimke(4).Caption = Ertek(sor)
                        
                    Case "nevjegy.fordito"
                        nevjegy.cimke(3).Caption = Ertek(sor)
                        
                    Case "nevjegy.koszonet"
                        szerkeszto.koszonet_mnu.Caption = Ertek(sor)
                        
                    Case "elem.uj"
                        szerkeszto.uj_elem_mnu.Caption = Ertek(sor)
                    Case "elem.torol"
                        szerkeszto.torles_mnu.Caption = Ertek(sor)
                    Case "elem.tulajdonsagok"
                        szerkeszto.tulajdonsag_mnu.Caption = Ertek(sor)
                        
                                 '##### BEÁLLÍTÁSOK ABLAK
                    'ÁLTALÁNOS FÜL
                    Case "beallitasok.altalanos"
                        beallitasok.beall_mnu(0).Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.beall_enged"
                        beallitasok.enged.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.egyeni_enged"
                        beallitasok.egyeni.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.szerkesztes_enged"
                        beallitasok.szerkesztes.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.behuzas_enged"
                        beallitasok.behuzas.Caption = Ertek(sor)
                    
                    Case "beallitasok.altalanos.kerdesek_enged"
                        beallitasok.kerdesek.Caption = Ertek(sor)
                        
                    Case "beallitasok.altalanos.tippek_enged"
                        beallitasok.tippek.Caption = Ertek(sor)
                    
                    Case "beallitasok.altalanos.potlo"
                        beallitasok.cetli(5).Caption = Ertek(sor)
                    'Étékelés fül
                    
                    Case "beallitasok.ertekeles"
                        beallitasok.beall_mnu(1).Caption = Ertek(sor)
                        
                    Case "beallitasok.ertekeles.feltetel"
                        beallitasok.cetli(2).Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.akkor"
                        beallitasok.cetli(0).Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.megnevezes"
                        beallitasok.cetli(1).Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.felvesz"
                        beallitasok.felvesz.Caption = Ertek(sor)
                    Case "beallitasok.ertekeles.torol"
                        beallitasok.torol.Caption = Ertek(sor)
                        
                        
                    Case "beallitasok.ertekeles.pont"
                        beallitasok.cetli(4).Caption = Ertek(sor)
                    
                    Case "beallitasok.ertekeles.levonas"
                        beallitasok.cetli(10).Caption = Ertek(sor)
                    'Ellenõrzés
                    Case "beallitasok.ellenorzes"
                        beallitasok.beall_mnu(2).Caption = Ertek(sor)
                        
                    Case "beallitasok.ellenorzes.kapitalis"
                        beallitasok.kisbetus.Caption = Ertek(sor)
                    
                    Case "beallitasok.ellenorzes.felesleges"
                        beallitasok.szokoz.Caption = Ertek(sor)
                        
                    Case "beallitasok.ellenorzes.kotojelek"
                        beallitasok.nincskotojel.Caption = Ertek(sor)
                    
                    Case "beallitasok.ellenorzes.szinek"
                        beallitasok.megoldas(0).Caption = Ertek(sor)
                    
                    Case "beallitasok.ellenorzes.eltuntet"
                        beallitasok.megoldas(1).Caption = Ertek(sor)
                    'Nyelvi modul
                    Case "beallitasok.nyelv"
                        beallitasok.beall_mnu(3).Caption = Ertek(sor)
                    Case "beallitasok.nyelv.magyar"
                        beallitasok.nyelv(0).Caption = Ertek(sor)
                    Case "beallitasok.nyelv.idegen"
                        beallitasok.nyelv(1).Caption = Ertek(sor)
                    Case "beallitasok.nyelv.betolt"
                        beallitasok.betoltnyelv.Caption = Ertek(sor)
                        
                    'Tipus almenü
                    Case "elem.tipus"
                        szerkeszto.tipus_mnu.Caption = Ertek(sor)
                    Case "elem.tipus.kerdezendo"
                        szerkeszto.tip_mnu(0).Caption = Ertek(sor)
                    Case "elem.tipus.megjegyzes"
                        szerkeszto.tip_mnu(2).Caption = Ertek(sor)
                    Case "elem.tipus.jelmagyarazat"
                        szerkeszto.tip_mnu(1).Caption = Ertek(sor)
                    
                    'Szöveg igazítás
                    Case "elem.igazitas"
                        szerkeszto.igazitas_mnu.Caption = Ertek(sor)
                    Case "elem.igazitas.ala"
                        szerkeszto.szoveg_igazit(0).Caption = Ertek(sor)
                    Case "elem.igazitas.fole"
                        szerkeszto.szoveg_igazit(1).Caption = Ertek(sor)
                    Case "elem.igazitas.kozepre"
                        szerkeszto.szoveg_igazit(2).Caption = Ertek(sor)
                    Case "elem.igazitas.jobbra"
                        szerkeszto.szoveg_igazit(3).Caption = Ertek(sor)
                    Case "elem.igazitas.balra"
                        szerkeszto.szoveg_igazit(4).Caption = Ertek(sor)
                    
                    Case "elem.meretez"
                        szerkeszto.meretez_mnu.Caption = Ertek(sor)
                        
                    'Tulajdonságok
                    '#### GOMBOK ELNEVEZÉSE
                    Case "gombok.megse"
                        tulajdonsagok.megse.Caption = Ertek(sor)
                        beallitasok.megse.Caption = Ertek(sor)
                        
                     Case "gombok.mentes"
                        beallitasok.alkalmaz.Caption = Ertek(sor)
                        
                    Case "gombok.ok"
                        tulajdonsagok.ok.Caption = Ertek(sor)
                        beallitasok.ok.Caption = Ertek(sor)
                        koszonet.ok.Caption = Ertek(sor)
                        
                    Case "gombok.sugo"
                        tulajdonsagok.sugo.Caption = Ertek(sor)
                        beallitasok.sugo.Caption = Ertek(sor)
                    'Általános fül
                    Case "tulajdonsagok.altalanos"
                        tulajdonsagok.ful(0).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.cim"
                        tulajdonsagok.cimke(0).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.kep"
                        tulajdonsagok.cimke(1).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.nagyitas"
                        tulajdonsagok.cimke(10).Caption = Ertek(sor)
                        
                    Case "tulajdonsagok.altalanos.nev"
                        tulajdonsagok.cimke(3).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.tipp"
                        tulajdonsagok.cimke(4).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.masolo"
                        tulajdonsagok.formatum_masolo.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.kikerdezendo"
                        tulajdonsagok.tipus(0).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.jelmagyarazat"
                        tulajdonsagok.tipus(1).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.altalanos.megjegyzes"
                        tulajdonsagok.tipus(2).Caption = Ertek(sor)
                    'jel fül
                    Case "tulajdonsagok.jel"
                        tulajdonsagok.ful(1).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.jel.jel"
                        tulajdonsagok.cimke(2).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.jel.kitoltes"
                        tulajdonsagok.kitolte.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.jel.keret"
                        tulajdonsagok.cimke(6).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.jel.vastagsag"
                        tulajdonsagok.cimke(7).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.jel.lathatatlan"
                        tulajdonsagok.elrejt.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.jel.vtajolas"
                        tulajdonsagok.cimke(12).Caption = Ertek(sor)
                    
                    'Szöveg
                    Case "tulajdonsagok.szoveg"
                        tulajdonsagok.ful(2).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.betutipus"
                        tulajdonsagok.cimke(5).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.meret"
                        tulajdonsagok.cimke(8).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.szin"
                        tulajdonsagok.valaszt(3).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.felkover"
                        tulajdonsagok.felkover.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.dolt"
                        tulajdonsagok.dolt.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.alahuzott"
                        tulajdonsagok.alahuzott.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.athuzva"
                        tulajdonsagok.athuzva.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.hatter"
                        tulajdonsagok.hatter.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.elrejt"
                        tulajdonsagok.elrejt_szov.Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.szoveg.hatterszin"
                        tulajdonsagok.valaszt(4).Caption = Ertek(sor)
                        
                    'Kérdések
                    Case "tulajdonsagok.kerdes"
                        tulajdonsagok.ful(3).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.kerdes.kerdes"
                        tulajdonsagok.cimke(9).Caption = Ertek(sor)
                    
                    Case "tulajdonsagok.kerdes.valasz"
                        tulajdonsagok.cimke(11).Caption = Ertek(sor)
                End Select
                
                If IsNumeric(parancs) Then
                    Index = CByte(parancs)
                    KozosSzovegek(Index) = Ertek(sor)
                End If
kihagySZ:
            Loop
        tulajdonsagok.UjraNyelvel
    End If
    Close 4
    
Exit Sub
Hiba:
    MsgBox "Language error - Nyelvi hiba", vbCritical, "Language error - Nyelvi hiba"
    magyar_nyelv
End Sub
Public Sub magyar_nyelv()
If Szulo.Name = "terkep" Then
        'Fájl
        terkep.fajl_mnu.Caption = "&Fájl"
        terkep.ujra_mnu.Caption = "Újra kezd"
        terkep.megnyit_mnu.Caption = "Megnyitás"
        terkep.nyomtat.Caption = "Nyomtatás..."
        'nyomtatas.Caption = Ertek(sor)
        terkep.szerkesztes_mnu.Caption = "Szerkesztés..."
        terkep.kilep_mnu.Caption = "Kilépés"
                    
        '##### ESZKÖZÖK MENÜ
        terkep.eszkozok_mnu.Caption = "&Eszközök"
        terkep.jelm_mnu.Caption = "Jelmagyarázat"
        jelmagyarazat.Caption = terkep.jelm_mnu.Caption
        ertekeles.ertekeles = "Értékelés"
        terkep.ertekel_mnu.Caption = ertekeles.ertekeles
        terkep.beall_mnu.Caption = "Beállítások..."
        beallitasok.Caption = "Beállítások..."
        terkep.behuzo_mnu.Caption = "Behúzó"
        behuzo.Caption = "Behúzó"
                        
        '###### SÚGÓ MENÜ
        terkep.sugo_mnu.Caption = "&Súgó"
        terkep.help_mnu.Caption = "Segítség"
        terkep.nevjegy_mnu.Caption = "Névjegy"
                    
        'Névjegy
        nevjegy.cimke(4).Caption = "Fordítás:"
        nevjegy.cimke(3).Caption = "Eredeti magyar nyelvû változat."
        terkep.koszonet_mnu.Caption = "Köszönet..."
                        
        '#### GOMBOK ELNEVEZÉSE
        beallitasok.megse.Caption = "&Mégse"
        'tulajdonsagok.megse.Caption = ""
        nyomtatas.megse.Caption = "&Mégse"
                        
        
        beallitasok.ok.Caption = "&OK"
        koszonet.ok.Caption = "&OK"
        
        beallitasok.alkalmaz.Caption = "Menté&s"
        nyomtatas.ok.Caption = "&Nyomtat"
        
        beallitasok.sugo.Caption = "Súgó"
        nyomtatas.sugo.Caption = "Súgó"
        nyomtatas.megse.Caption = "&Bezár"
        
        '##### BEÁLLÍTÁSOK ABLAK
        'ÁLTALÁNOS FÜL
        beallitasok.beall_mnu(0).Caption = "Általános"
        beallitasok.enged.Caption = "Beállítások menüpont engedélyezése"
        beallitasok.egyeni.Caption = "Projektfájlok egyéni beállításainak engedélyezése"
        beallitasok.szerkesztes.Caption = "Szerkesztés menü engedélyezése"
        beallitasok.behuzas.Caption = "Behúzások engedélyezése"
        beallitasok.kerdesek.Caption = "Segítõ kérdések"
        beallitasok.tippek.Caption = "Gyorstippek"
        beallitasok.cetli(5).Caption = "Pótlójel:"
        'Étékelés fül
        beallitasok.beall_mnu(1).Caption = "Értékelés"
        'beallitasok.cetli(0).Caption = "Elégséges alsó határa:"
        'beallitasok.cetli(1).Caption = "Közepes alsó határa:"
        'beallitasok.cetli(2).Caption = "Jó alsó határa:"
        'beallitasok.cetli(3).Caption = "Példás alsó határa:"
        beallitasok.cetli(2).Caption = "Ha az eredmény <="
        beallitasok.cetli(0).Caption = "%, akkor"
        beallitasok.cetli(1).Caption = "Megnevezés:"
        beallitasok.felvesz.Caption = "Felvesz"
        beallitasok.torol.Caption = "Töröl"
        beallitasok.cetli(4).Caption = "Egy feladatra adható pont:"
        beallitasok.cetli(10).Caption = "Kérdésenkénti százalék levonás:"
        'Ellenõrzés
        beallitasok.beall_mnu(2).Caption = "Ellenõrzés"
        beallitasok.kisbetus.Caption = "Kis és nagy betû nem számít"
        beallitasok.szokoz.Caption = "Feleslges szóközöket figyelmen kívül hagy"
        beallitasok.nincskotojel.Caption = "Kötõjelek figyelmen kívül hagyása"
        beallitasok.megoldas(0).Caption = "Javítás színekkel"
        beallitasok.megoldas(1).Caption = "Csak a helytelenek maradnak"
        'nyelv fül
        beallitasok.beall_mnu(3).Caption = "Nyelv"
        beallitasok.nyelv(0).Caption = "Eredeti magyar nyelv használata"
        beallitasok.nyelv(1).Caption = "Idegen nyelvi modul alkalmazása"
        beallitasok.betoltnyelv.Caption = "&Nyelv alkalmazása"
        
        
        'Nyomtatás
        'nyomtatas.Caption = "Nyomtatás"
        nyomtatas.cimke(1).Caption = "Szélesség:"
        nyomtatas.cimke(2).Caption = "Magasság:"
        
        'Értékelés
        ertekeles.cetli(1).Caption = "Helyes válaszok száma:"
        ertekeles.cetli(0).Caption = "Hibák száma:"
        ertekeles.cetli(2).Caption = "Eredmény:"
        ertekeles.cetli(5).Caption = "Pontok:"
        ertekeles.cetli(4).Caption = "Érdemjegy:"
    Else
        'Projekt
        szerkeszto.file.Caption = "&Projekt"
        szerkeszto.uj_mnu.Caption = "Új projekt"
        szerkeszto.megnyit_mnu.Caption = "Megnyitás"
        szerkeszto.ment_mnu.Caption = "Mentés"
        szerkeszto.ment_mint_mnu.Caption = "Mentés másként"
        szerkeszto.olda_mnu.Caption = "Projekt tulajdonságai"
        szerkeszto.megtekint_mnu.Caption = "Megtekintés"
        szerkeszto.beall_mnu.Caption = "Beállítások"
        szerkeszto.kilepes_mnu.Caption = "Kilépés"
        'Súgó
        szerkeszto.sugo_mnu.Caption = "&Súgó"
        szerkeszto.sugo_mnup.Caption = "Segítség"
        szerkeszto.nevjegy_mnu.Caption = "Névjegy"
                    
        'Névjegy
        nevjegy.cimke(4).Caption = "Fordítás:"
        nevjegy.cimke(3).Caption = "Eredeti magyar nyelvû változat."
        szerkeszto.koszonet_mnu.Caption = "Köszönet..."
        
        '##### BEÁLLÍTÁSOK ABLAK
        'ÁLTALÁNOS FÜL
        beallitasok.beall_mnu(0).Caption = "Általános"
        beallitasok.enged.Caption = "Beállítások menüpont engedélyezése"
        beallitasok.egyeni.Caption = "Projektfájlok egyéni beállításainak engedélyezése"
        beallitasok.szerkesztes.Caption = "Szerkesztés menü engedélyezése"
        beallitasok.behuzas.Caption = "Behúzások engedélyezése"
        beallitasok.kerdesek.Caption = "Segítõ kérdések"
        beallitasok.tippek.Caption = "Gyorstippek"
        beallitasok.cetli(5).Caption = "Pótlójel:"
        'Étékelés fül
        beallitasok.beall_mnu(1).Caption = "Értékelés"
        'beallitasok.cetli(0).Caption = "Elégséges alsó határa:"
        'beallitasok.cetli(1).Caption = "Közepes alsó határa:"
        'beallitasok.cetli(2).Caption = "Jó alsó határa:"
        'beallitasok.cetli(3).Caption = "Példás alsó határa:"
        beallitasok.cetli(2).Caption = "Ha az eredmény <="
        beallitasok.cetli(0).Caption = "%, akkor"
        beallitasok.cetli(1).Caption = "Megnevezés:"
        beallitasok.felvesz.Caption = "Felvesz"
        beallitasok.torol.Caption = "Töröl"
        beallitasok.cetli(4).Caption = "Egy feladatra adható pont:"
        beallitasok.cetli(10).Caption = "Kérdésenkénti százalék levonás:"
        'Ellenõrzés
        beallitasok.beall_mnu(2).Caption = "Ellenõrzés"
        beallitasok.kisbetus.Caption = "Kis és nagy betû nem számít"
        beallitasok.szokoz.Caption = "Feleslges szóközöket figyelmen kívül hagy"
        beallitasok.nincskotojel.Caption = "Kötõjelek figyelmen kívül hagyása"
        beallitasok.megoldas(0).Caption = "Javítás színekkel"
        beallitasok.megoldas(1).Caption = "Csak a helytelenek maradnak"
        'nyelv fül
        beallitasok.beall_mnu(3).Caption = "Nyelv"
        beallitasok.nyelv(0).Caption = "Eredeti magyar nyelv használata"
        beallitasok.nyelv(1).Caption = "Idegen nyelvi modul alkalmazása"
        beallitasok.betoltnyelv.Caption = "&Nyelv alkalmazása"
        'Nyelv
        beallitasok.megse.Caption = "&Mégse"
        beallitasok.ok.Caption = "&OK"
        beallitasok.sugo.Caption = "&Súgó"
        beallitasok.alkalmaz.Caption = "Menté&s"
        
        
        'Elem menu
        szerkeszto.uj_elem_mnu.Caption = "Új elem"
        szerkeszto.torles_mnu.Caption = "Töröl"
        szerkeszto.tulajdonsag_mnu.Caption = "Tulajdonságok"
        
        'Tipus almenü
        szerkeszto.tipus_mnu.Caption = "Típusa"
        szerkeszto.tip_mnu(0).Caption = "Kikérdezendõ"
        szerkeszto.tip_mnu(2).Caption = "Megjegyzés"
        szerkeszto.tip_mnu(1).Caption = "Jelmagyarázat"
                    
        'Szöveg igazítás
        szerkeszto.igazitas_mnu.Caption = "Szöveg igazítása"
        szerkeszto.szoveg_igazit(0).Caption = "Alá"
        szerkeszto.szoveg_igazit(1).Caption = "Fölé"
        szerkeszto.szoveg_igazit(2).Caption = "Középre"
        szerkeszto.szoveg_igazit(3).Caption = "Jobbra"
        szerkeszto.szoveg_igazit(4).Caption = "Balra"
        szerkeszto.meretez_mnu.Caption = "Jel méretezése"
                        
        'Tulajdonságok
        '#### GOMBOK ELNEVEZÉSE
        tulajdonsagok.megse.Caption = "&Mégse"
        tulajdonsagok.ok.Caption = "&OK"
        beallitasok.ok.Caption = "&OK"
        koszonet.ok.Caption = "&OK"
        tulajdonsagok.sugo.Caption = "&Súgó"
                    
        'Általános fül
        tulajdonsagok.ful(0).Caption = "&Általános"
        tulajdonsagok.cimke(0).Caption = "Cím:"
        tulajdonsagok.cimke(1).Caption = "Kép:"
        tulajdonsagok.cimke(10).Caption = "Nagyítás:"
        tulajdonsagok.cimke(3).Caption = "Név:"
        tulajdonsagok.cimke(4).Caption = "Tipp szövege:"
        tulajdonsagok.formatum_masolo.Caption = "Formátum másoló"
        tulajdonsagok.tipus(0).Caption = "Kikérdezendõ"
        tulajdonsagok.tipus(1).Caption = "Jelmagyarázat (külön ablakban)"
        tulajdonsagok.tipus(2).Caption = "Megjegyzés, felirat"
        'jel fül
        tulajdonsagok.ful(1).Caption = "&Jel"
        tulajdonsagok.cimke(2).Caption = "Jel típusa:"
        tulajdonsagok.kitolte.Caption = "Kitöltés:"
        tulajdonsagok.cimke(6).Caption = "Keret típusa:"
        tulajdonsagok.cimke(7).Caption = "Vastagsága:"
        tulajdonsagok.elrejt.Caption = "Láthatatlan alakzat"
        tulajdonsagok.cimke(12).Caption = "Vonal tájolása:"
        'Szöveg
        tulajdonsagok.ful(2).Caption = "&Szöveg"
        tulajdonsagok.cimke(5).Caption = "Betûtípus:"
        tulajdonsagok.cimke(8).Caption = "Méret:"
        tulajdonsagok.valaszt(3).Caption = "Betûszín"
        tulajdonsagok.felkover.Caption = "Félkövér"
        tulajdonsagok.dolt.Caption = "Dõlt"
        tulajdonsagok.alahuzott.Caption = "Aláhúzott"
        tulajdonsagok.athuzva.Caption = "Áthúzva"
        tulajdonsagok.hatter.Caption = "Van háttérszine"
        tulajdonsagok.elrejt_szov.Caption = "Szöveg elrejtése"
        tulajdonsagok.valaszt(4).Caption = "Háttér"
                        
        'Kérdések
        tulajdonsagok.ful(3).Caption = "&Kérdések"
        tulajdonsagok.cimke(9).Caption = "Kérdés:"
        tulajdonsagok.cimke(11).Caption = "Válasz:"
        
        'Alakzatok nevei
        'tul_alakzat_szov(1) = "Négyzet"
End If
'Közös stringek
    KozosSzovegek(1) = "Ön letiltotta a Beállítások menüt. Ahhoz, hogy újra el tudja érni a beállításokat, indítsa a programot a -beall kapcsolóval az alábbi módon: 'vakablak.exe -beall'."
    KozosSzovegek(2) = "Az új beállítások érvénybelépéséhez újra kell indítanom az alkalmazást. Kívánja most megtenni?"
    KozosSzovegek(3) = "Projekt megnyitása..."
    KozosSzovegek(4) = "Vakablak fájlok"
    KozosSzovegek(5) = "Projekt mentése másként ..."
    KozosSzovegek(6) = "Projekt mentése ..."
    KozosSzovegek(7) = "Szerkesztõ mód"
    KozosSzovegek(8) = "Névtelen projekt"
    KozosSzovegek(9) = "Az alapkép elérési útja nem adható meg relatívan. Kívánja, hogy a képet a projekt mellé másoljam? Ha a nemet választja, akkor a teljes elérési út lesz elmentve."
    KozosSzovegek(10) = "$sz térképelem egy másik fájlra hivatkozik. Kívánja, hogy ezt a fájlt a projekt mellé másoljam?"
    KozosSzovegek(11) = "Külsõ fájlok kezelése:"
    KozosSzovegek(12) = "A '$sz' fájl nem tölthetõ be."
    KozosSzovegek(13) = "A szerkesztésre megadott fájl nem található!"
    KozosSzovegek(14) = "A megadott projekt hibás bejegyzéseket tartalmaz, ami bizonytalanná teheti a program futását. Kívánja folytatni töltést?"
    KozosSzovegek(15) = "Ismeretlen hiba"
    KozosSzovegek(16) = "A megadott kép hibás, ismeretlen tömörítésû vagy nem található a megadott helyen. $sz Kívánja folytatni a töltést a hiba javításához?"
    KozosSzovegek(17) = "Képbetöltési hiba"
    KozosSzovegek(18) = "A '$sz' módosításait nem mentette el. Kívánja most menteni azt?"
    KozosSzovegek(19) = "Módosítások mentése"
    KozosSzovegek(20) = "Végsõ eredmény"
    KozosSzovegek(21) = "A megadott fájl nem található!"
    KozosSzovegek(22) = "A megadott projekt hibás bejegyzéseket tartalmaz, ami bizonytalanná teheti a program futását. A töltés e hiba miatt megszakad."
    KozosSzovegek(23) = "$sz. kérdés:"
    KozosSzovegek(24) = "$sz tulajdonságai"
    KozosSzovegek(25) = "Kép keresése ..."
    KozosSzovegek(26) = "Minden támogatott képformátum"
    KozosSzovegek(27) = "Négyzet"
    KozosSzovegek(28) = "Ellipszis"
    KozosSzovegek(29) = "Kör"
    KozosSzovegek(30) = "Kerekített négyszög"
    KozosSzovegek(31) = "Kerekített négyzet"
    KozosSzovegek(32) = "Kép"
    KozosSzovegek(33) = "Vonal"
    KozosSzovegek(34) = "Nincs"
    KozosSzovegek(35) = "Elem színével"
    
    'hibák
    KozosSzovegek(36) = "A témakörhöz tartozó súgót nem lehet megnyitni!"
    KozosSzovegek(37) = "Súgó hiba:"
    KozosSzovegek(38) = "Megnyitási hiba"
    KozosSzovegek(39) = "Mentési hiba"
    KozosSzovegek(40) = "Nem tudok a meghajtóra írni, mert írásvédett."
    KozosSzovegek(41) = "A lemez nincs a meghajtóban, vagy nincs formázva."
    KozosSzovegek(42) = "A fájl nem írható, mert írásvédett."
    KozosSzovegek(43) = "Nyomtatási hiba"
    KozosSzovegek(44) = "Nem találom a kiválasztott nyomtatót!"
    KozosSzovegek(45) = "A projekt nyomtatása meghiúsult, mert hibás nyomtatási méretet adott meg."
    KozosSzovegek(46) = "A projekt nem nyomtatható, mert túl nagy nyomtatási méretet adott meg."
    KozosSzovegek(47) = "Négyszög" 'Kimaradt
    
    If Szulo.Name = "szerkeszto" Then
        tulajdonsagok.UjraNyelvel
    End If
End Sub
Private Sub Terkepeles()
    Set Szulo = terkep
    Szulo.jel(0).jel = 0
    magyar_nyelv
    terkep.megnyitas ("$vt\vakablak.ini")
    terkep.Show
End Sub
