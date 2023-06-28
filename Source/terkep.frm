VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form terkep 
   BackColor       =   &H8000000A&
   Caption         =   "Vaktérkép"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6615
   Icon            =   "terkep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer kuldo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.HScrollBar jb 
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   4080
      Width           =   5535
   End
   Begin VB.VScrollBar fl 
      Height          =   4575
      Left            =   5640
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin MSComDlg.CommonDialog pb 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox terulet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   720
      ScaleHeight     =   4020
      ScaleWidth      =   5220
      TabIndex        =   0
      Top             =   480
      Width           =   5250
      Begin VB.Frame opciok 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   -120
         TabIndex        =   6
         Top             =   4680
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox hatarok 
            Height          =   285
            Index           =   4
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   17
            Text            =   "91"
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox hatarok 
            Height          =   285
            Index           =   3
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "75"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox hatarok 
            Height          =   285
            Index           =   2
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   15
            Text            =   "60"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox hatarok 
            Height          =   285
            Index           =   1
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "52"
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox pont1 
            Height          =   285
            Left            =   1800
            MaxLength       =   1
            TabIndex        =   13
            Text            =   "1"
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label cetli 
            Caption         =   "pont"
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   12
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label cetli 
            Caption         =   "Egy feladatra adható"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label cetli 
            Caption         =   "Példás alsó határa:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label cetli 
            Caption         =   "Jó alsó határa:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label cetli 
            Caption         =   "Közepes alsó határa:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label cetli 
            Caption         =   "Elégséges alsó határa:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.TextBox szoveg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label fedo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Címke"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   3810
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape pont 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   1200
         Shape           =   3  'Circle
         Top             =   3840
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Menu file 
      Caption         =   "&Fájl"
      Begin VB.Menu open 
         Caption         =   "Térkép megnyitása"
         Shortcut        =   ^O
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Kilépés"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu ellenorzes 
      Caption         =   "&Ellenõrzés"
      Begin VB.Menu szigor 
         Caption         =   "Szigor"
         Visible         =   0   'False
      End
      Begin VB.Menu ertekeles 
         Caption         =   "Értékelés"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu sugo 
      Caption         =   "&Súgó"
      Begin VB.Menu help 
         Caption         =   "Súgó"
         Shortcut        =   {F1}
      End
      Begin VB.Menu v2 
         Caption         =   "-"
      End
      Begin VB.Menu nevjegy 
         Caption         =   "Névjegy"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "terkep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aktualis As Integer, koordinatak As Boolean
Dim nevek(0 To 256) As String, hibak(0 To 256) As Integer, teljes As Integer
Dim eleres As String, ellenorzo(1 To 256) As Integer

Private Sub cimke_Click(Index As Integer)
aktualis = Index
beiro
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub fedo_Click(Index As Integer)
aktualis = Index
beiro
End Sub


Private Sub ertekeles_Click()
On Error Resume Next
Dim talalatok As Integer
Dim k As Integer
For i = 1 To teljes
    If Trim(UCase(nevek(i))) = Trim(UCase(cimke(i).Caption)) Then
            talalatok = talalatok + 1
            cimke(i).BackStyle = 1
            cimke(i).BackColor = vbGreen
        Else
            hibak(i) = 1
            cimke(i).BackStyle = 1
            cimke(i).BackColor = vbRed
            
            
    End If
    
Next i

bizi.helyes.Caption = talalatok
bizi.hibak.Caption = teljes - talalatok
bizi.szazalek = Format(talalatok / teljes * 100, "##,##")
bizi.pontok = talalatok * pont1
bizi.maxpont = "/ " & teljes * pont1
Select Case CByte(bizi.szazalek)
    Case 0 To hatarok(1) - 1
        bizi.jegy = 1
        bizi.neve = "Elégtelen"
        'If bizi.szazalek = 0 Then bizi.jegy = bizi.jegy & ",": bizi.neve = bizi.neve & " alá"
        'If bizi.szazalek = hatarok(1) - 1 Then bizi.jegy = bizi.jegy & "*": bizi.neve = "Csillagos " & bizi.neve
    
    Case hatarok(1) To hatarok(2) - 1
        bizi.jegy = 2
        bizi.neve = "Elégséges"
        'If bizi.szazalek = hatarok(1) Then bizi.jegy = bizi.jegy & ",": bizi.neve = bizi.neve & " alá"
        'If bizi.szazalek = hatarok(2) - 1 Then bizi.jegy = bizi.jegy & "*": bizi.neve = "Csillagos " & bizi.neve
    
    Case hatarok(2) To hatarok(3) - 1
        bizi.jegy = 3
        bizi.neve = "Közepes"
        'If bizi.szazalek = hatarok(2) Then bizi.jegy = bizi.jegy & ",": bizi.neve = bizi.neve & " alá"
        'If bizi.szazalek = hatarok(3) - 1 Then bizi.jegy = bizi.jegy & "*": bizi.neve = "Csillagos " & bizi.neve

    Case hatarok(3) To hatarok(4) - 1
        bizi.jegy = 4
        bizi.neve = "Jó"
        'If bizi.szazalek = hatarok(3) Then bizi.jegy = bizi.jegy & ",": bizi.neve = bizi.neve & " alá"
        'If bizi.szazalek = hatarok(4) - 1 Then bizi.jegy = bizi.jegy & "*": bizi.neve = "Csillagos " & bizi.neve
        
    Case hatarok(4) To 100
        bizi.jegy = 5
        bizi.neve = "Példás"
        'If bizi.szazalek = hatarok(4) Then bizi.jegy = bizi.jegy & ",": bizi.neve = bizi.neve & " alá"
        'If bizi.szazalek = 100 Then bizi.jegy = bizi.jegy & "*": bizi.neve = "Csillagos " & bizi.neve

        
End Select
bizi.Show vbModal

End Sub

Private Sub fl_Change()
terulet.Top = fl.Value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim ful  As String
'MsgBox KeyAscii
If szoveg.Visible Then
    Select Case KeyAscii
        Case 13
            cimke(aktualis).Caption = szoveg.Text
            'cimke(aktualis).Width = (Len(szoveg.Text)) * 100 + 200
            szoveg.Visible = False
            aktualis = 0
        Case 27
            szoveg.Visible = False
            aktualis = 0
    End Select
Else
    Select Case KeyAscii
        Case 126
            koordinatak = Not koordinatak
        Case 123
            Me.Caption = Me.Caption & " - Automatikus kitöltés..."
            kitolt
        Case 244
            alaphelyzet
        Case 232
            Me.Caption = Me.Caption & " - Memória listázva volt..."
            kuldo.Enabled = Not kuldo.Enabled
            sysmon.Visible = Not sysmon.Visible
    End Select
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
terkep.terulet.Picture = sysmon.kep.Picture
Me.Caption = "Vaktérkép " & App.Major & "." & App.Minor
If Len(App.Path) = 3 Then eleres = Mid(App.Path, 1, 2) Else eleres = App.Path
'MkDir eleres & "\megnyitva"
koordinatak = False
ertekeles.Enabled = False
'jb.Visible = True

End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim X As Integer, Y As Integer
X = (terkep.ScaleWidth - terulet.Width) / 2
Y = (terkep.ScaleHeight - terulet.Height) / 2

fl.Move terkep.ScaleWidth - fl.Width, 0, fl.Width, terkep.ScaleHeight - fl.Width
jb.Move 0, terkep.ScaleHeight - jb.Height, terkep.ScaleWidth - jb.Height, jb.Height
terulet.Move X, Y

If terkep.ScaleWidth - terulet.Width < 0 Then
        jb.SmallChange = Int(terkep.ScaleWidth - terulet.Width / 100)
        jb.LargeChange = Int(terkep.ScaleWidth - terulet.Width / 10)
        jb.Max = terkep.ScaleWidth - terulet.Width
        jb.Min = 0
        jb.Visible = True
    Else
        jb.Visible = False
End If

If terkep.ScaleHeight - terulet.Height < 0 Then
        fl.SmallChange = Int(terkep.ScaleHeight - terulet.Height / 100)
        fl.LargeChange = Int(terkep.ScaleHeight - terulet.Height / 10)
        fl.Max = terkep.ScaleHeight - terulet.Height
        fl.Min = 0
        fl.Visible = True
    Else
        fl.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'Kill eleres & "\megnyitva\*.*"
'RmDir eleres & "\megnyitva"
totalki
End Sub

Private Sub jb_Change()
terulet.Left = jb.Value
End Sub

Private Sub kuldo_Timer()
Dim kep As String
kep = "Memória Állapota:" & vbCrLf
kep = kep & "-----------------" & vbCrLf
For i = 1 To teljes
    kep = kep & "cimke(" & i & ").Caption=" & nevek(i) & vbCrLf
Next i
kep = kep & "Összesen:" & teljes & " objektum" & vbCrLf
kep = kep & "Koordináta listázás:" & koordinatak & vbCrLf
kep = kep & "Aktív objektum:" & aktualis
sysmon.monitor = kep
End Sub

Private Sub nevjegy_Click()
frmAbout.Show vbModal
End Sub

Private Sub open_Click()
On Error GoTo hiba
pb.DialogTitle = "Térkép megnyitása ..."
pb.Filter = "Térkép projektek (*.vtk)|*.vtk"
pb.FileName = eleres & "\*.vtk"
pb.ShowOpen
'tomorit pb.FileName
Call alaphelyzet
tolt (pb.FileName)
hiba:
End Sub

Public Sub beiro()
'MsgBox cimke(aktualis).Caption, vbCritical, aktualis
szoveg.Width = (Len(cimke(aktualis).Caption)) * 100 + 200
szoveg.Text = cimke(aktualis).Caption
szoveg.Move cimke(aktualis).Left, cimke(aktualis).Top
szoveg.Visible = True
End Sub
Public Sub tolt(fajlnev As String)
Dim parancs As String, ertek As String, kod As Integer, sor As String, ker As Integer, i As Integer
Dim tipus As Byte, X As Integer, Y As Integer, nev As String, szin As ColorConstants
Dim kover As Boolean, dolt As Boolean, alahuzott As Boolean, feltolt(1 To 10) As String
Dim meret As Byte, konyvtar As String, jobbra As Boolean

'konyvtar meghatározása
j = 0
    For i = 1 To Len(fajlnev)
        If Mid(fajlnev, i, 1) = "\" Then j = i
    Next i
konyvtar = Mid(fajlnev, 1, j)
'MsgBox konyvtar

'MsgBox vbBlack & "  " & vbRed & "   " & vbBlue
meret = 9
szin = vbBlack
kod = 0
dolt = False
kover = False
alahuzott = False
jobbra = False

On Error GoTo fajlhiba
Open fajlnev For Input As 1
    Do While Not EOF(1)
        Line Input #1, sor
            parancs = ""
            ertek = ""
            On Error GoTo kephiba
            'On Error Resume Next
            For ker = 1 To Len(sor)
            If Mid(sor, ker, 1) = "=" Then
                parancs = Mid(sor, 1, ker - 1)
                ertek = Mid(sor, ker + 1, Len(sor) - ker)
                GoTo gyorski
            End If
        Next ker
    If parancs = "" Then GoTo ki
gyorski:
    Select Case parancs
        Case "cim"
            Me.Caption = ertek & " - Vaktérkép " & App.Major & "." & App.Minor
        Case "terkep"
            If Mid(ertek, 1, 1) = "\" Then
                    ertek = Mid(ertek, 2, Len(ertek) - 1)
                    terulet.Picture = LoadPicture(konyvtar & ertek)
                Else
                    terulet.Picture = LoadPicture(ertek)
            End If
        Case "szin"
            szin = ertek
        Case "kover"
            kover = ertek
        Case "dolt"
            dolt = ertek
        Case "alahuzott"
            alahuzott = ertek
        Case "meret"
            meret = ertek
        Case "jobbra"
            jobbra = ertek
        Case "elem"
            kod = kod + 1
            For i = 1 To 10
                feltolt(i) = ""
            Next i
            i = 1
            j = 1
        
                    For ker = 1 To Len(ertek)
                        If Mid(ertek, ker, 1) = "," Then
                            feltolt(j) = Mid(ertek, i, ker - i)
                            i = ker + 1
                            j = j + 1
                        End If
                    Next
                    feltolt(j) = Mid(ertek, i, Len(ertek) + 1 - i)
                    
                    nev = feltolt(4)
                    X = feltolt(2)
                    Y = feltolt(3)
                    tipus = feltolt(1)
                    'End
                    'MsgBox tipus & "    " & X & " " & Y & " " & nev
                        
                            Load pont(kod)
                            pont(kod).Left = X
                            pont(kod).Top = Y
                            pont(kod).Shape = tipus
                            pont(kod).BorderColor = szin
                            pont(kod).FillColor = szin
                            pont(kod).Visible = True
                            
                            
                            Load cimke(kod)
                            cimke(kod).FontSize = meret
                            cimke(kod).Caption = "?"
                            cimke(kod).ForeColor = szin
                            cimke(kod).FontBold = kover
                            cimke(kod).FontItalic = dolt
                            cimke(kod).FontUnderline = alahuzott
                            cimke(kod).Alignment = Abs(CInt(jobbra))
                            cimke(kod).Top = Y - 30
                                Select Case jobbra
                                    Case 0
                                        cimke(kod).Left = pont(kod).Width + pont(kod).Left + 15
                                    Case 1
                                        cimke(kod).Left = pont(kod).Left - (cimke(kod).Width + 15)
                                End Select
                                    
                                
                            cimke(kod).Visible = True
                            nevek(kod) = nev
                            
                            
                            Load fedo(kod)
                            fedo(kod).Move pont(kod).Left, pont(kod).Top, pont(kod).Width, pont(kod).Height
                            fedo(kod).Visible = True
                        
           ' Next i
objki:
        Case "vege"
            Exit Sub
    End Select
ki:
    Loop

Close 1
teljes = kod
ertekeles.Enabled = True
Form_Resize
Exit Sub


kephiba:
    MsgBox "A projectben megadott kép nem elérhetõ, ezért a project betöltése megszakad.", vbCritical, "A kép nem elérhetõ"
        alaphelyzet
        Close 1
        Form_Resize
        Exit Sub


fajlhiba:
    MsgBox "A megadott elérési út helytelen, vagy nem Vaktérkép fájl.(" & fajlnev & ")", vbCritical, "A project nem nyitható meg..."
        alaphelyzet
        Close 1
        Exit Sub


End Sub



Private Sub szigor_Click()
opciok.Visible = Not opciok.Visible
End Sub

Private Sub szoveg_Change()
szoveg.Width = (Len(szoveg.Text) + 1) * 120 + 150
End Sub

Private Sub terulet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If koordinatak = True Then MsgBox X & "            " & Y
aktualis = 0
szoveg.Visible = False
End Sub
'Public Sub tomorit(fajl As String)
'On Error GoTo hiba
'Shell eleres & "\rar.exe x " & fajl & " megnyitva", vbHide
'Do While Dir(eleres & "\megnyitva\terkep.vtk", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = ""
'Loop
'tolt (eleres & "\megnyitva\terkep.vtk")
'Exit Sub
'hiba:
'MsgBox "Nemtalálom a külsõ tömörítõt!"
'End Sub
Public Sub kitolt()
    For i = 1 To teljes
        cimke(i).Caption = nevek(i)
         cimke(i).Width = (Len(nevek(i))) * 100 + 200
    Next i
    
End Sub
Public Sub alaphelyzet()
    For i = 1 To teljes
        nevek(i) = ""
        Unload cimke(i)
        Unload pont(i)
        Unload fedo(i)
    Next i
    teljes = 0
    terulet.Picture = Nothing
    aktualis = 0
    Form_Load
End Sub
