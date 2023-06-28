Attribute VB_Name = "input"
Public i As Integer, j As Double, eleres As String, k As String
Sub Main()
Dim parancs As String, ertek As String
parancs = Trim(Command$)
If Len(App.Path) = 3 Then eleres = Mid(App.Path, 1, 2) Else eleres = App.Path
Select Case UCase(parancs)
    Case "/SZERK"
        szerk.Show
        Exit Sub
        
    Case "/BEALL"
        terkep.sett.Visible = True
    
    Case Else
            If parancs <> "" Then
                If Mid(parancs, 1, 1) = Chr(34) Then
                        parancs = Mid(parancs, 2, Len(parancs) - 2)
                End If
    
                Call terkep.tolt(parancs)

            End If
End Select
Call terkep.tolt(eleres & "\vakterkep.ini")
terkep.Show
End Sub

Public Sub totalki()
    Unload jelol
    Unload bizi
    Unload szerk
    Unload terkep
    Unload tul
    End
End Sub

Public Function perenbol(szoveg As String)
Dim l As String, ker As Integer
    l = ""
    ker = 1
            Do While ker <= Len(szoveg)
                If Mid(szoveg, ker, 2) = "\n" Then
                    l = l & vbCrLf
                    ker = ker + 2
                Else
                    l = l & Mid(szoveg, ker, 1)
                    ker = ker + 1
                End If
            Loop
            perenbol = l
End Function

Public Function perenbe(szoveg As String)
Dim l As String, ker As Integer
    l = ""
    ker = 1
    Do While ker <= Len(szoveg)
                If Mid(szoveg, ker, 1) = Chr(13) Then
                        l = l & "\n"
                        ker = ker + 2
                Else
                        l = l & Mid(szoveg, ker, 1)
                        ker = ker + 1
                End If
    Loop
    perenbe = l
End Function
