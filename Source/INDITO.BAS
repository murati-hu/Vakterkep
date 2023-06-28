Attribute VB_Name = "kozos"
Option Explicit
Public Vakterkep As Program
Dim Szulo As Object
Dim szovegek(1 To 100) As String

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

Sub Main()
On Error GoTo Hiba
    'If True Then debuger.Show
    Vakterkep.Verzio = App.Major & "." & App.Minor & " Build: " & App.Revision
    If Len(App.Path) = 3 Then
            Vakterkep.Konyvtar = App.Path
        Else
            Vakterkep.Konyvtar = App.Path & "\"
    End If
    Set Szulo = terkep
    Szulo.jel(0).jel = 0
    
    Select Case LCase(Utasitas(Command$))
        Case "-szerkeszto", "-sz", "-s"
            Set Szulo = szerkeszto
            szerkeszto.Show
            If Ertek(Command$) <> "" Then
                    szerkeszto.megnyitas (Ertek(Command$))
                Else
                    'tulajdonsagok.megse.Enabled = False
                    'tulajdonsagok.Mutat (0)
            End If
        Case "-beallitas_nelkul", "-nobea", "-n"
            'Beállítások nélkül
            Set Szulo = terkep
            terkep.Show
        'Case "-megnyit", "-m"
         '   terkep.megnyitas ("$vt\vakterkep.ini")
          '  terkep.megnyitas (Ertek(Command$))
           ' terkep.Show
        Case ""
            Set Szulo = terkep
            terkep.megnyitas ("$vt\vakterkep.ini")
            terkep.Show
        Case Else
            Set Szulo = terkep
            terkep.megnyitas ("$vt\vakterkep.ini")
            terkep.Show
            terkep.megnyitas (Command$)
            Exit Sub
    End Select
    'Szulo.Show
Exit Sub
Hiba:
    Select Case Err.Number
        Case 339
            MsgBox "Kibõvített jel ActiveX vezérlõ nincs telepítve a rendszerén. Kérem telepítse újra az alkalmazást.", vbCritical, "Hiányzó 'jelek.ocx' vezérlõ!"
            MsgBox Err.Description, vbCritical, "A hibáról bõvebben(" & Err.Number & ")"
            Megsemmisit
        Case 374
            MsgBox "Egy olyan ActiveX vezérlõ van telepítve, amely nem kompatibilis a program ezen verziójával. Kérem telepítse újra az alkalmazás.", vbCritical, "Inkompatibilis ActiveX vezérlõ"
            MsgBox Err.Description, vbCritical, "A hibáról bõvebben(" & Err.Number & ")"
            Megsemmisit
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

Public Function Atalakit(Adat As String, egyeb As Variant)
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
    Set Szulo = Nothing
    Unload beallitasok
    Unload jelmagyarazat
    Unload nevjegy
    Unload ertekeles
    Unload nyomtatas
    Unload szerkeszto
    Unload terkep
    Unload tulajdonsagok
    End
End Sub
Public Function uzenet(Index As Integer, Beszurando As String) As String
    If InStr(1, szovegek(Index), "$sz") Then
            uzenet = Atalakit(szovegek(Index), Beszurando)
        Else
            uzenet = szovegek(Index)
    End If
End Function
Public Sub magyar()
    
End Sub
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
            Melyiket = "vakterkep.chm::/page/" & Melyiket
        Else
            Melyiket = "szerkeszto.chm::/page/" & Melyiket
    End If
    Shell "hh.exe " & Vakterkep.Konyvtar & Melyiket, vbNormalFocus
Exit Sub
Hiba:
    MsgBox "A témakörhöz tartozó súgót nem lehet megnyitni!", vbCritical, "Súgó hiba:"
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
