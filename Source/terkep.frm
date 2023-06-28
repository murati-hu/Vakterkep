VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form terkep 
   BackColor       =   &H8000000C&
   Caption         =   "Vaktérkép"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6120
   Icon            =   "terkep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton gomb 
      Caption         =   "J"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer ora 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   1560
   End
   Begin VB.HScrollBar jb 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   5295
   End
   Begin VB.VScrollBar fl 
      Height          =   4095
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox terulet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   360
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   120
      Width           =   5295
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
      Begin VB.Label segito 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000017&
         Height          =   225
         Left            =   1080
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label fedo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label cimke 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Címke"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   3810
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape pont 
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         DrawMode        =   1  'Blackness
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   1320
         Shape           =   3  'Circle
         Top             =   3840
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog pb 
      Left            =   1800
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu file 
      Caption         =   "&Fájl"
      Begin VB.Menu retrn 
         Caption         =   "Újra kezd"
         Enabled         =   0   'False
         Shortcut        =   ^U
      End
      Begin VB.Menu open 
         Caption         =   "Térkép megnyitása"
         Shortcut        =   ^M
      End
      Begin VB.Menu nyomtat 
         Caption         =   "Nyomtatás..."
         Visible         =   0   'False
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Kilépés"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu eszkozok 
      Caption         =   "&Eszközök"
      Begin VB.Menu jelm 
         Caption         =   "Jelmagyarázat"
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu ertekeles 
         Caption         =   "Értékelés"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu sett 
         Caption         =   "Beállítások..."
         Shortcut        =   ^B
         Visible         =   0   'False
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
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "terkep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aktualis As Byte, ttip As Byte
Dim segitseg(1 To 256, 0 To 4) As String, hanyadik As Byte
Dim nevek(0 To 256) As String, pontok(0 To 256) As Double
Dim megoldas(1 To 256, 0 To 4) As String, egyeni(1 To 256) As Byte
Dim gyorstip(1 To 256) As String


Private Sub cimke_Click(Index As Integer)
ttip_el
aktualis = Index
beiro
End Sub

Private Sub cimke_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    kijelol (ttip)
    ttip = Index
    ora.Enabled = True
End Sub

Private Sub exit_Click()
If retrn.Enabled Then
    i = MsgBox("Ha most kilép, minden eddigi eredménye el fog veszni. Biztosan ki akar lépni?", vbQuestion + vbYesNo, "Kilépés megerõsítése")
Else
    i = vbYes
End If
If i = vbYes Then Unload Me
End Sub

Private Sub fedo_Click(Index As Integer)
ttip_el
aktualis = Index
beiro
End Sub


Private Sub ertekeles_Click()
'On Error Resume Next
Dim jo As Integer
If hanyadik = 5 Then
    bizi.Show
    Exit Sub
End If

For i = 1 To (cimke.Count - 1)
    If pontok(i) = 0 Then 'And cimke(i).Visible = True
        If Trim(UCase(nevek(i))) = Trim(UCase(cimke(i).Caption)) Then
            cimke(i).BackStyle = 1
            cimke(i).BackColor = vbGreen
            cimke(i).Enabled = False
            fedo(i).Enabled = False
            If hanyadik = 0 Or opciok.segito = 0 Then
                pontok(i) = opciok.pont
            Else
                pontok(i) = opciok.pont * ((100 - ((egyeni(i) + 1) * opciok.szazal)) / 100)
            End If
            jo = jo + 1
        Else
            cimke(i).BackStyle = 1
            cimke(i).BackColor = 9934847 'vbRed
        End If
    Else
        jo = jo + 1
    End If
Next i

hanyadik = hanyadik + 1

j = 0
For i = 1 To (cimke.Count - 1)
    j = j + pontok(i)
    'If cimke(i).Enabled Then
        'cimke(i).ToolTipText = segitseg(i, hanyadik)
    'End If
Next i


bizi.helyes.Caption = jo
bizi.hibak.Caption = (cimke.Count - 1) - jo
bizi.pontok.Caption = j
bizi.maxpont.Caption = (cimke.Count - 1) * opciok.pont

bizi.szazalek.Caption = Format(CDbl(bizi.pontok.Caption) / CDbl(bizi.maxpont.Caption) * 100, "##,##")
If bizi.szazalek.Caption = "" Then bizi.szazalek.Caption = 0
Select Case CByte(bizi.szazalek)
    Case 0 To opciok.hatarok(1) - 1
        bizi.jegy = 1
        bizi.neve = "Elégtelen"
       
    Case opciok.hatarok(1) To opciok.hatarok(2) - 1
        bizi.jegy = 2
        bizi.neve = "Elégséges"
        
    Case opciok.hatarok(2) To opciok.hatarok(3) - 1
        bizi.jegy = 3
        bizi.neve = "Közepes"
        
    Case opciok.hatarok(3) To opciok.hatarok(4) - 1
        bizi.jegy = 4
        bizi.neve = "Jó"
        
    Case opciok.hatarok(4) To 100
        bizi.jegy = 5
        bizi.neve = "Példás"
End Select

If hanyadik = 5 Then
    ertekeles.Caption = "Értékelés mutatása"
    bizi.Caption = "Utólsó értékelés"
Else
    If hanyadik = 4 Then
        ertekeles.Caption = "Utolsó értékelés"
    Else
        ertekeles.Caption = hanyadik + 1 & ". értékelés"
    End If
    bizi.Caption = hanyadik & ". értékelés"
End If
bizi.Show vbModal
End Sub

Private Sub fedo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    kijelol (ttip)
    ttip = Index
    ora.Enabled = True
End Sub

Private Sub fl_Change()
terulet.Top = fl.Value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim ful  As String
If szoveg.Visible Then
    Select Case KeyAscii
        Case 13
            cimke(aktualis).Caption = szoveg.Text
            szoveg.Visible = False
            aktualis = 0
        Case 27
            szoveg.Visible = False
            aktualis = 0
    End Select
Else
    Select Case KeyAscii
        Case 244
            alaphelyzet
        
    End Select
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Caption = "Vaktérkép " & App.Major & "." & App.Minor
ertekeles.Enabled = False
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

If (fl.Visible Or jb.Visible) Then
        gomb.Move fl.Left, jb.Top
        gomb.Visible = True
        gomb.Enabled = jelm.Enabled
Else
        gomb.Visible = False
        
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

totalki
End Sub

Private Sub gomb_Click()
jelm_Click
End Sub

Private Sub help_Click()
On Error GoTo hiba
Shell "hh.exe " & eleres & "\vakterkep.chm", vbNormalFocus
Exit Sub
hiba:
    MsgBox "Az ön Windowsa nem képes kezelni a HTML Help fájlokat.", vbInformation, "Súgó nem tölthetõ be"
End Sub

Private Sub jb_Change()
terulet.Left = jb.Value
End Sub


Public Sub jelm_Click()
If jelol.Visible Then
    jelol.Visible = False
Else
    jelol.Show vbModeless, terkep
End If
End Sub

Private Sub nevjegy_Click()
frmAbout.Show vbModal
End Sub

Private Sub nyomtat_Click()
terkep.PrintForm
End Sub

Private Sub open_Click()
On Error GoTo hiba
pb.DialogTitle = "Térkép megnyitása ..."
pb.Filter = "Térkép projektek (*.vtk)|*.vtk"
pb.FileName = "*.vtk"
pb.ShowOpen
Call alaphelyzet
tolt (pb.FileName)
hiba:
End Sub

Public Sub beiro()
'On Error GoTo ki
Call ttip_el
i = aktualis
Call Form_KeyPress(27)
aktualis = i
If ((megoldas(aktualis, egyeni(aktualis)) = "" Or hanyadik = 0 Or cimke(aktualis).LinkTimeout = 0) And egyeni(aktualis) <= 4) Then
mehet:
    'If (cimke(aktualis).LinkTimeout = 50 And opciok.segito = 1 And hanyadik > 0) Then Exit Sub
    szoveg.Width = (Len(cimke(aktualis).Caption)) * 100 + 200
    szoveg.Text = cimke(aktualis).Caption
    szoveg.SelStart = 0
    szoveg.SelLength = Len(szoveg.Text)
    szoveg.Move cimke(aktualis).Left, cimke(aktualis).Top
    szoveg.Visible = True
    szoveg.SetFocus
    Exit Sub
Else
    k = InputBox(segitseg(aktualis, egyeni(aktualis)), egyeni(aktualis) + 1 & ". segítõ kérdés:")
    If k = "" Then
        Exit Sub
    Else
        If Trim(UCase(k)) <> Trim(UCase(megoldas(aktualis, egyeni(aktualis)))) Then
            egyeni(aktualis) = egyeni(aktualis) + 1
            If egyeni(aktualis) = 5 Then egyeni(aktualis) = 4
        Else
            cimke(aktualis).LinkTimeout = 0
            GoTo mehet
        End If
    End If
End If
ki:
End Sub
Public Sub tolt(fajlnev As String)
Dim parancs As String, ertek As String, kod As Integer, sor As String, ker As Integer, i As Integer
Dim tipus As Byte, X As Integer, Y As Integer, nev As String, szin As ColorConstants
Dim kover As Boolean, dolt As Boolean, alahuzott As Boolean, feltolt(1 To 10) As String
Dim meret As Byte, konyvtar As String, jobbra As Boolean, lathatatlan As Boolean

'konyvtar meghatározása
j = 0
    For i = 1 To Len(fajlnev)
        If Mid(fajlnev, i, 1) = "\" Then j = i
    Next i
konyvtar = Mid(fajlnev, 1, j)
'MsgBox konyvtar

meret = 9
szin = vbBlack
kod = 0
dolt = False
kover = False
alahuzott = False
jobbra = False
lathatatlan = False

If fajlnev = eleres & "\vakterkep.ini" Then
    On Error GoTo nincs
Else
    On Error GoTo fajlhiba
End If
Open fajlnev For Input As 1
    Do While Not EOF(1)
        Line Input #1, sor
            parancs = ""
            ertek = ""
            'On Error Resume Next
            For ker = 1 To Len(sor)
            If Mid(sor, ker, 1) = "=" Then
                parancs = Mid(sor, 1, ker - 1)
                ertek = Mid(sor, ker + 1, Len(sor) - ker)
                GoTo gyorski
            End If
        Next ker
    If parancs = "" Then parancs = sor
    parancs = LCase(parancs)
gyorski:
    Select Case parancs
        Case "cim"
            Me.Caption = ertek & " - Vaktérkép " & App.Major & "." & App.Minor
        Case "terkep"
            On Error GoTo kephiba
            If Mid(ertek, 1, 1) = "\" Then
                    ertek = Mid(ertek, 2, Len(ertek) - 1)
                    terulet.Picture = LoadPicture(konyvtar & ertek)
                Else
                    terulet.Picture = LoadPicture(ertek)
            End If
            On Error GoTo egyeb
        Case "szin"
            szin = ertek
        Case "kover"
            kover = True
        Case "dolt"
            dolt = True
        Case "alahuzott"
            alahuzott = True
        Case "meret"
            meret = ertek
        Case "balra"
            jobbra = True
        Case "lathatatlan"
            lathatatlan = True
        Case "elem"
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
                    
                    tipus = feltolt(1)
                    X = feltolt(2)
                    Y = feltolt(3)
                    nev = feltolt(4)
                    
                     Select Case tipus
                        Case 0 To 6
                            kod = kod + 1
                            Load pont(kod)
                            pont(kod).Left = X
                            pont(kod).Top = Y
                            pont(kod).BorderColor = szin
                            pont(kod).FillColor = szin
                            'pont(kod).Visible = True
                            
                            
                            Load cimke(kod)
                            cimke(kod).FontSize = meret
                            cimke(kod).Caption = opciok.jel
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
                            
                            If lathatatlan = False Then
                                pont(kod).Shape = tipus
                                pont(kod).Visible = True
                            End If
                       Case 7
    
                            i = jelol.pont.Count
                            i = CInt(i)
                            Load jelol.pont(i)
                            jelol.pont(i).Left = jelol.pont(i - 1).Left
                            jelol.pont(i).Top = jelol.cimke(i - 1).Top + jelol.cimke(i - 1).Height + 30
                            'jelol.pont(i).Shape = X
                            jelol.pont(i).BorderColor = szin
                            jelol.pont(i).FillColor = szin
                            'jelol.pont(i).Visible = True
                            
                            
                            Load jelol.cimke(i)
                            jelol.cimke(i).FontSize = meret
                            jelol.cimke(i).Caption = nev
                            jelol.cimke(i).ForeColor = szin
                            jelol.cimke(i).FontBold = kover
                            jelol.cimke(i).FontItalic = dolt
                            jelol.cimke(i).FontUnderline = alahuzott
                            jelol.cimke(i).Top = jelol.cimke(i - 1).Top + jelol.cimke(i - 1).Height + 30
                            jelol.cimke(i).Left = jelol.cimke(i - 1).Left
                            jelol.cimke(i).Visible = True
                            
                            If jelol.Width < jelol.cimke(i).Left + jelol.cimke(i).Width + 100 Then
                                    jelol.Width = jelol.cimke(i).Left + jelol.cimke(i).Width + 100
                            End If
                            jelol.Height = jelol.cimke(i).Top + jelol.cimke(i).Height + 500
                            jelm.Enabled = True
                            
                            If lathatatlan = False Then
                                jelol.pont(i).Shape = X
                                jelol.pont(i).Visible = True
                            End If
                        
                End Select
                meret = 9
                szin = vbBlack
                dolt = False
                kover = False
                alahuzott = False
                jobbra = False
                lathatatlan = False
                
                retrn.Enabled = True
                ertekeles.Enabled = True
                ertekeles.Caption = "Értékelés"
objki:
        Case "hatarok"
            If opciok.beall.Value = 1 Or fajlnev = eleres & "\vakterkep.ini" Then
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
                    
                    For i = 1 To 4
                        opciok.hatarok(i) = feltolt(i)
                    Next i
                End If
        Case "tipp"
            If opciok.tippek = 1 Then
                gyorstip(kod) = perenbol(ertek)
            End If
        Case "kerdes"
            If opciok.segito = 1 Then
            For i = 1 To 10
                feltolt(i) = ""
            Next i
            i = 1
            j = 1
        
                    For ker = 1 To Len(ertek)
                     If j < 3 Then
                        If Mid(ertek, ker, 1) = "," Then
                            feltolt(j) = Mid(ertek, i, ker - i)
                            i = ker + 1
                            j = j + 1
                        End If
                    Else
                        GoTo ki2
                    End If
                    Next
ki2:
                    feltolt(j) = Mid(ertek, i, Len(ertek) + 1 - i)
            
            If feltolt(2) = "" Then feltolt(2) = " "
            
            i = feltolt(1) - 1
            megoldas(kod, i) = feltolt(2)
            segitseg(kod, i) = perenbol(feltolt(3))
            End If
        
        Case "beallitas"
            opciok.beeng.Value = ertek
            sett.Visible = Abs(CInt(ertek))
            
        Case "egyeni"
            opciok.beall.Value = ertek
        Case "jel"
            opciok.jel = ertek
        Case "pont"
            opciok.pont = ertek
        Case "segito"
            opciok.segito = ertek
        Case "tippek"
            opciok.tippek = ertek
        Case "minusz"
            opciok.szazal = ertek
        Case "vege"
            Close 1
            Exit Sub
    End Select
ki:
    Loop

Close 1
Form_Resize
retrn_Click
Exit Sub


kephiba:
    MsgBox "A projektben megadott kép nem elérhetõ, ezért a projekt betöltése megszakad.", vbCritical, "A kép nem elérhetõ"
        alaphelyzet
        Close 1
        Form_Resize
        Exit Sub


fajlhiba:
    MsgBox "A megadott elérési út helytelen, vagy nem Vaktérkép fájl.(" & fajlnev & ")", vbCritical, "A projekt nem nyitható meg..."
        alaphelyzet
        Close 1
        Exit Sub
egyeb:
    MsgBox "A töltés meg fog szakadni az alábbi hiba miatt:" & vbCrLf & Err.Description, vbCritical, "Hiba(" & Err.Number & ")"
    alaphelyzet
    Form_Resize
    Close 1
    Exit Sub
nincs:
    Close 1
End Sub

Private Sub ora_Timer()
On Error GoTo ki
If gyorstip(ttip) = "" Or szoveg.Visible = True Then GoTo ki
segito.Caption = gyorstip(ttip)
segito.Move pont(ttip).Left, cimke(ttip).Top + cimke(ttip).Height + 10
segito.Visible = True
ki:
End Sub

Private Sub retrn_Click()
On Error Resume Next
    For i = 1 To (cimke.Count - 1)
        cimke(i).Caption = opciok.jel
        cimke(i).Enabled = True
        cimke(i).BackStyle = 0
        fedo(i).Enabled = True
        pontok(i) = 0
        egyeni(i) = 0
        cimke(i).LinkTimeout = 50
    Next i
Unload bizi
ertekeles.Caption = "Értékelés"
aktualis = 0
hanyadik = 0
End Sub

Private Sub sett_Click()
opciok.Show vbModal
End Sub

Private Sub szoveg_Change()
szoveg.Width = (Len(szoveg.Text) + 1) * 120 + 150
End Sub

Private Sub terulet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
aktualis = 0
szoveg.Visible = False
End Sub

Public Sub alaphelyzet()
On Error Resume Next
    For i = 1 To (cimke.Count - 1)
        nevek(i) = ""
        Unload cimke(i)
        Unload pont(i)
        Unload fedo(i)
        gyorstip(i) = ""
        For j = 0 To 4
            megoldas(i, j) = ""
            segitseg(i, j) = ""
        Next j
    Next i
    For i = 1 To jelol.cimke.Count
        Unload jelol.cimke(i)
        Unload jelol.pont(i)
    Next i
        
    terulet.Picture = Nothing
    aktualis = 0
    hanyadik = 0
    Form_Load
End Sub

Private Sub terulet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ttip_el
End Sub
Private Sub ttip_el()
ora.Enabled = False
segito.Visible = False
ttip = 0
For i = 1 To cimke.Count - 1
        pont(i).BorderStyle = 0
        cimke(i).BorderStyle = 0
Next i
End Sub
Sub kijelol(id As Integer)
    'pont(id).BorderStyle = 1
    'cimke(id).BorderStyle = 1
End Sub
