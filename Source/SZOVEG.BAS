Attribute VB_Name = "szoveg"
Option Explicit
Public Function Utasitas(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    Utasitas = Mid(Adatsor, 1, i - 1)
        End If
        i = i + 1
    Loop
    If Not megvan Then Utasitas = Adatsor
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

Public Function Kicsontoz(szoveg As String, Elvalaszto As String, Melyiket As Byte) As String
    Dim i As Integer, j As Integer, ker As Integer
    Dim cella(0 To 255) As String
    i = 1
    j = 0
        
    For ker = 1 To Len(szoveg)
        If Mid(szoveg, ker, 1) = Elvalaszto Then
            cella(j) = Mid(szoveg, i, ker - i)
            i = ker + 1
            j = j + 1
        End If
    Next
    cella(j) = Mid(szoveg, i, Len(szoveg) + 1 - i)
    Kicsontoz = cella(Melyiket)
End Function
Public Function Korulmetel(szoveg As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(szoveg) And Not megvan
        If Mid(szoveg, i, 1) <> Chr(9) And Mid(szoveg, i, 1) <> " " Then
            megvan = True
            szoveg = Mid(szoveg, i, Len(szoveg) - i + 1)
        End If
        i = i + 1
    Loop
    megvan = False
    i = Len(szoveg)
    Do While i >= 1 And Not megvan
        If Mid(szoveg, i, 1) <> Chr(9) And Mid(szoveg, i, 1) <> " " Then
            megvan = True
            szoveg = Mid(szoveg, 1, i)
        End If
        i = i - 1
    Loop
    Korulmetel = szoveg
End Function

Public Function VanEBenne(Miben As String, Mi As String) As Boolean
    If InStr(1, Miben, Mi) = 0 Then
        VanEBenne = False
    Else
        VanEBenne = True
    End If
End Function
