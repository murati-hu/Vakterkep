Attribute VB_Name = "modul"
Option Explicit
Public Konyvtar As String

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long




Sub Main()
    If Len(App.Path) = 3 Then
            Konyvtar = App.Path
        Else
            Konyvtar = App.Path & "\"
    End If
    With csomagolo
    If Trim(Command$) = "" Then
            .Show
        Else
            Dim nev As String
            nev = CsakANeve(Command$)
            nev = Mid(nev, 1, Len(nev) - 4)
            Osszepakol Command$, Konyvtara(Command$) & nev & ".zip"
            MsgBox "Kész."
            End
    End If
    End With
End Sub
Public Function utasitas(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    utasitas = Mid(Adatsor, 1, i - 1)
        End If
        i = i + 1
    Loop
    If Not megvan Then utasitas = Adatsor
End Function
Public Function Ertek(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    Ertek = Mid(Adatsor, i + 1, Len(Adatsor) - i)
        End If
        i = i + 1
    Loop
    If Not megvan Then Ertek = ""
End Function
Public Function Korulmetel(Szoveg As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Szoveg) And Not megvan
        If Mid(Szoveg, i, 1) <> Chr(9) And Mid(Szoveg, i, 1) <> " " Then
            megvan = True
            Szoveg = Mid(Szoveg, i, Len(Szoveg) - i + 1)
        End If
        i = i + 1
    Loop
    megvan = False
    i = Len(Szoveg)
    Do While i >= 1 And Not megvan
        If Mid(Szoveg, i, 1) <> Chr(9) And Mid(Szoveg, i, 1) <> " " Then
            megvan = True
            Szoveg = Mid(Szoveg, 1, i)
        End If
        i = i - 1
    Loop
    Korulmetel = Szoveg
End Function
Public Function Atalakit(Adat As String, Optional egyeb As Variant)
    Dim i As Integer, uj As String
    uj = ""
    i = 1
    Do While i <= Len(Adat)
        If Mid(Adat, i, 3) = "$vt" Then
            uj = uj & Mid(Konyvtar, 1, Len(Konyvtar) - 1)
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
        
        uj = uj & Mid(Adat, i, 1)
        i = i + 1
    Loop
    Atalakit = uj
End Function
Public Function Konyvtara(Fajlnev As String)
    Dim i As Integer, j As Integer
    j = 0
    For i = 1 To Len(Fajlnev)
        If Mid(Fajlnev, i, 1) = "\" Then j = i
    Next i
    Konyvtara = Mid(Fajlnev, 1, j)
End Function
Public Function CsakANeve(Eleres As String)
    CsakANeve = Mid(Eleres, Len(Konyvtara(Eleres)) + 1, Len(Eleres))
End Function
Public Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lRet As Long
    
    sBuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sBuffer)
    
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    GetTempPathName = sBuffer
    
End Function
Public Function Osszepakol(forras As String, cel As String) As Long
On Error GoTo hiba:
    Dim temp As String
    Dim fajl As String
    Dim sor As String
    Dim parameter As String

    'ideiglenes könyvtár létrehozása
    temp = GetTempPathName
    Randomize
    temp = temp & "vakablak-temp" & Second(Time()) & CLng(Rnd() * 100) & "\"
    MkDir temp
    csomagolo.loggol temp
    
    'Mappa létrehozása a fájlnév alapján
    fajl = CsakANeve(forras)
    fajl = Mid(fajl, 1, Len(fajl) - 4)
    fajl = SpecNelkul(fajl)
    MkDir temp & fajl & "\"
    'csomagolo.loggol "Könyvtár létrehozása: " & temp & fajl & "\"
    
    'Célfájl megnyitása írásra #1
    Open temp & fajl & ".vtk" For Output As 1
        'Forrásfájl megnyitása olvasásra #2
        Open forras For Input As 2
            Do While Not EOF(2)
                Line Input #2, sor
                
                parameter = Korulmetel(Ertek(sor))
                
                Select Case UCase(utasitas(Korulmetel(sor)))
                    Case "CIM"
                        Print #1, "cim=" & parameter
                        
                    Case "KEP"
                        'Alapkép megkeresése és másolása
                        parameter = Atalakit(parameter, Konyvtara(forras))
                        FileCopy parameter, temp & fajl & "\" & SpecNelkul(CsakANeve(parameter))
                        Print #1, "kep=\" & fajl & "\" & SpecNelkul(CsakANeve(parameter))
                        'csomagolo.loggol "Alapkép OK"
                    Case "IKON"
                        parameter = Atalakit(parameter, Konyvtara(forras))
                        FileCopy parameter, temp & fajl & "\" & SpecNelkul(CsakANeve(parameter))
                        Print #1, "         ikon=\" & fajl & "\" & SpecNelkul(CsakANeve(parameter))
                        'csomagolo.loggol "Csatolt kép OK"
                    Case Else
                        Print #1, sor
                End Select
            Loop
        Close 2
    Close 1
    
    csomagolo.loggol "Projekt összekészítve"
    'csomagolo.loggol "----------------------"
    csomagolo.loggol "Tömörítés..."
    
    FileCopy Konyvtar & "info-zip-bin\zip23.exe", temp & "zip23.exe"
    ChDir temp
    'csomagolo.loggol temp & "zip23.exe " & fajl & ".zip " & fajl & ".vtk " & fajl & "\*.*"
    Shell temp & "zip23.exe " & cel & " .\*.vtk .\*\*.*", vbNormalFocus
    
    'Do While 1 = 1
    'Loop
    'ChDir App.Path
    'MsgBox "van"
    'FileCopy temp & fajl & ".zip", Cel
    'Masold temp & fajl & ".zip", Cel
    csomagolo.loggol "Kész."
    
    Osszepakol = 0
Exit Function
hiba:
    Osszepakol = Err.Number
    csomagolo.loggol "HIBA: " & Err.Description
End Function
Private Sub Masold(Mit As String, Hova As String)
On Error GoTo hiba
    FileCopy Mit, Hova
Exit Sub
'Addig próbálgatni, amíg nincs kész
hiba:
    'Resume
End Sub
Private Function SpecNelkul(Mit As String) As String
Dim i As Integer
SpecNelkul = ""
Mit = LCase(Mit)
    For i = 1 To Len(Mit)
        Select Case Mid(Mit, i, 1)
            Case "ö", "õ", "ó"
                SpecNelkul = SpecNelkul & "o"
            Case "ü", "û", "ú"
                SpecNelkul = SpecNelkul & "u"
            Case "é"
                SpecNelkul = SpecNelkul & "e"
            Case "á"
                SpecNelkul = SpecNelkul & "a"
            Case "í"
                SpecNelkul = SpecNelkul & "i"
            Case "?", "@", " "
                SpecNelkul = SpecNelkul & ""
            Case Else
                SpecNelkul = SpecNelkul & Mid(Mit, i, 1)
        End Select
    Next i
End Function
