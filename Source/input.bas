Attribute VB_Name = "input"
Public i As Integer, j As Integer
Sub Main()
Dim parancs As String, ertek As String
parancs = Trim(Command$)
    'For i = 1 To Len(parancs)
        'If Mid(parancs, i, 1) = " " Then
            'ertek = Mid(parancs, i + 1, Len(parancs) - i)
            'parancs = Mid(parancs, 1, i - 1)
        'End If
    'Next i
'MsgBox parancs
Select Case UCase(parancs)
    Case "/SZERK"
        szerk.Show
        Exit Sub
    Case Else
            If parancs <> "" Then
                ' Ha a windows küldte, akkor az idézõjel levágása
                If Mid(parancs, 1, 1) = Chr(34) Then
                        parancs = Mid(parancs, 2, Len(parancs) - 2)
                End If
    
                Call terkep.tolt(parancs)

            End If
End Select
terkep.Show
End Sub

Public Sub totalki()
    Unload sysmon
    Unload bizi
    Unload szerk
    Unload terkep
    Unload tul
    End
End Sub


