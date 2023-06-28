VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form szerk 
   BackColor       =   &H8000000A&
   Caption         =   "Vaktérkép Szerkesztõ"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7155
   Icon            =   "szerk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5310
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6315
            MinWidth        =   1235
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5786
            MinWidth        =   706
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog pb 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.HScrollBar jb 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   5535
   End
   Begin VB.VScrollBar fl 
      Height          =   4575
      Left            =   5520
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox terulet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   840
      MousePointer    =   2  'Cross
      ScaleHeight     =   3900
      ScaleWidth      =   5100
      TabIndex        =   1
      Top             =   360
      Width           =   5130
      Begin VB.TextBox szoveg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label fedo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   2040
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Címke"
         Height          =   195
         Index           =   0
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   3360
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape pont 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   3360
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Menu file 
      Caption         =   "&Térképek"
      Begin VB.Menu new 
         Caption         =   "&Új térkép"
         Shortcut        =   ^N
      End
      Begin VB.Menu v5 
         Caption         =   "-"
      End
      Begin VB.Menu open 
         Caption         =   "Térkép megnyitása"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Térkép mentése"
         Shortcut        =   ^S
      End
      Begin VB.Menu v4 
         Caption         =   "-"
      End
      Begin VB.Menu tuls 
         Caption         =   "Tulajdonságok"
         Shortcut        =   ^P
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Kilépés"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Szerkesztés"
      Visible         =   0   'False
      Begin VB.Menu uj 
         Caption         =   "Új pont"
         Begin VB.Menu elem 
            Caption         =   "Város"
            Index           =   1
         End
         Begin VB.Menu elem 
            Caption         =   "Terület"
            Index           =   2
         End
      End
      Begin VB.Menu rename 
         Caption         =   "Átnevez"
      End
      Begin VB.Menu del 
         Caption         =   "Töröl"
      End
      Begin VB.Menu replace 
         Caption         =   "Áthelyez"
      End
      Begin VB.Menu v3 
         Caption         =   "-"
      End
      Begin VB.Menu props 
         Caption         =   "Tulajdonságok"
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
Attribute VB_Name = "szerk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public aktualis As Integer, teljes As Integer, athelyez As Boolean
Dim eleres As String, ellenorzo(1 To 256) As Integer, px As Integer, py As Integer
Public kepneve As String, terkepneve As String
Public szerkesztett As Boolean




Private Sub cimke_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        aktualis = Index
        elemmenu
    End If
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub fedo_Mousedown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
    Case 1
        If athelyez = True And aktualis = Index Then replace_Click
    Case 2
        aktualis = Index
        elemmenu
End Select
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
            'cimke(aktualis).Width = (Len(szoveg.Text)) * 100 + 200
            szoveg.Visible = False
            aktualis = 0
        Case 27
            szoveg.Visible = False
            aktualis = 0
    End Select

End If
End Sub

Private Sub Form_Load()
On Error Resume Next
'szerk.terulet.Picture = sysmon.kep.Picture
terulet.Move terulet.Left, terulet.Top, 6000, 5000
Me.Caption = "Vaktérkép Szerkesztõ " & App.Major & "." & App.Minor
If Len(App.Path) = 3 Then eleres = Mid(App.Path, 1, 2) Else eleres = App.Path
'MkDir eleres & "\megnyitva"
'jb.Visible = True
szerkesztett = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sb.Panels(1).Text = "Ez nem"
sb.Panels(2).Text = "rajzterület"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim X As Integer, Y As Integer
X = (szerk.ScaleWidth - terulet.Width) / 2
Y = (szerk.ScaleHeight - terulet.Height) / 2

fl.Move szerk.ScaleWidth - fl.Width, 0, fl.Width, szerk.ScaleHeight - fl.Width - sb.Height
jb.Move 0, szerk.ScaleHeight - jb.Height - sb.Height, szerk.ScaleWidth - jb.Height, jb.Height
terulet.Move X, Y

If szerk.ScaleWidth - terulet.Width < 0 Then
        jb.SmallChange = Int(szerk.ScaleWidth - terulet.Width / 100)
        jb.LargeChange = Int(szerk.ScaleWidth - terulet.Width / 10)
        jb.Max = szerk.ScaleWidth - terulet.Width
        jb.Min = 0
        jb.Visible = True
    Else
        jb.Visible = False
End If

If szerk.ScaleHeight - terulet.Height < 0 Then
        fl.SmallChange = Int(szerk.ScaleHeight - terulet.Height / 100)
        fl.LargeChange = Int(szerk.ScaleHeight - terulet.Height / 10)
        fl.Max = szerk.ScaleHeight - terulet.Height - sb.Height
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

Private Sub nevjegy_Click()
frmAbout.Show vbModal
End Sub

Private Sub new_Click()
If szerkesztett Then
    i = MsgBox("Új project létrehozásával, minden eddigi munka el fog veszni." & vbCrLf & vbCrLf & "Kívánja menteni a jelenlegi projectet?", vbYesNoCancel + vbQuestion, "Új project létrehozása...")
    
    Select Case i
        Case vbYes
            save_Click
            alaphelyzet
        Case vbNo
            alaphelyzet
    End Select
Else
        alaphelyzet
End If
Form_Resize

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
szerkesztett = False
hiba:
End Sub

Public Sub beiro()
'MsgBox cimke(aktualis).Caption, vbCritical, aktualis
szoveg.Width = (Len(cimke(aktualis).Caption)) * 100 + 200
szoveg.Text = cimke(aktualis).Caption
szoveg.Move cimke(aktualis).Left, cimke(aktualis).Top
szoveg.Visible = True
szerkesztett = True
End Sub
Public Sub tolt(fajlnev As String)
Dim parancs As String, ertek As String, kod As Integer, sor As String, ker As Integer, i As Integer
Dim tipus As Byte, X As Integer, Y As Integer, nev As String, szin As ColorConstants
Dim kover As Boolean, dolt As Boolean, alahuzott As Boolean, meret As Byte, feltolt(1 To 10) As String
Dim konyvtar As String, jobbra As Boolean

'konyvtar meghatározása
j = 0
    For i = 1 To Len(fajlnev)
        If Mid(fajlnev, i, 1) = "\" Then j = i
    Next i
konyvtar = Mid(fajlnev, 1, j)

'alapertekek megadasa
meret = 9
szin = vbBlack
kod = 0
dolt = False
kover = False
alahuzott = False
jobbra = False

terulet.Picture = Nothing
terulet.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
Call Form_Resize
'project töltése

On Error GoTo fajlhiba
Open fajlnev For Input As 1
    Do While Not EOF(1)
        On Error Resume Next 'Hibás értékeket ugorja át
        Line Input #1, sor
            parancs = ""
            ertek = ""
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
            Me.Caption = ertek & " - Vaktérkép Szerkesztõ " & App.Major & "." & App.Minor
            terkepneve = ertek
        Case "terkep"
            If Mid(ertek, 1, 1) = "\" Then
                    ertek = Mid(ertek, 2, Len(ertek) - 1)
                    terulet.Picture = LoadPicture(konyvtar & ertek)
                    kepneve = konyvtar & ertek
                Else
                    terulet.Picture = LoadPicture(ertek)
                    kepneve = ertek
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
                    
                    'MsgBox tipus & "    " & X & " " & Y & " " & nev
                        
                            Load pont(kod)
                            pont(kod).Left = X
                            pont(kod).Top = Y
                            pont(kod).Shape = tipus
                            pont(kod).BorderColor = szin
                            pont(kod).FillColor = szin
                            pont(kod).Visible = True
                            
                            
                            Load cimke(kod)
                            cimke(kod).Caption = nev
                            cimke(kod).ForeColor = szin
                            cimke(kod).FontBold = kover
                            cimke(kod).FontItalic = dolt
                            cimke(kod).FontUnderline = alahuzott
                            cimke(kod).FontSize = meret
                            cimke(kod).Top = Y - 30
                            Call igazit(kod, Abs(CInt(jobbra)))
                               
                                
                            'cimke(kod).Width = Len(nev) * 100 + 200 /* az autosize miatt
                            cimke(kod).Visible = True
                            Load fedo(kod)
                            fedo(kod).Move pont(kod).Left, pont(kod).Top, pont(kod).Width, pont(kod).Height
                            fedo(kod).Visible = True
                       
objki:
        Case "vege"
            Exit Sub
    End Select
ki:
    Loop
Close 1
teljes = kod
Form_Resize
Exit Sub

fajlhiba:
    MsgBox "A megadott elérési út helytelen, vagy nem Vaktérkép fájl.(" & fajlnev & ")", vbCritical, "A project nem nyitható meg..."
        alaphelyzet
        Close 1
        Exit Sub


End Sub



Public Sub picopen()
On Error GoTo megse
    pb.DialogTitle = "Kép megnyitása..."
    pb.Filter = "Bitmap képek(*.bmp)|*.bmp|GIF képek(*.gif)|*.gif|Jpg képek(*.jpg)|*.jpg|JPE képek(*.jpe)|*.jpe|Jpeg képek(*.jpeg)|*.jpeg|Minden fájl(*.*)|*.*"
    pb.ShowOpen
    terulet.Picture = LoadPicture(pb.FileName)
    tul.kep = pb.FileName
megse:
    Form_Resize
    szerkesztett = True
End Sub


Private Sub rename_Click()
beiro
End Sub

Private Sub replace_Click()
athelyez = Not athelyez
If athelyez = True Then
    replace.Caption = "Rögzít"
Else
    replace.Caption = "Áthelyez"
End If

End Sub

Private Sub save_Click()
Dim konyvtar As String, kepfajl As String
On Error GoTo megse
    pb.DialogTitle = "Térkép mentése..."
    pb.Filter = "Vaktérkép project(*.vtk)|*.vtk"
    pb.FileName = terkepneve & ".vtk"
    pb.ShowSave


'belsõ struktúra
 j = 0
    For i = 1 To Len(pb.FileName)
        If Mid(pb.FileName, i, 1) = "\" Then j = i
    Next i
konyvtar = Mid(pb.FileName, 1, j) 'mentés könyvtára
j = 0
    For i = 1 To Len(kepneve)
        If Mid(kepneve, i, 1) = "\" Then j = i
    Next i
kepfajl = Mid(kepneve, j + 1, Len(kepneve) - j) ' csak a mentendõ képfájl
        
        
'mentés
On Error Resume Next
        FileCopy kepneve, konyvtar & kepfajl
       kepneve = "\" & kepfajl
        mentes pb.FileName
megse:
szerkesztett = False
End Sub

Private Sub szoveg_Change()
szoveg.Width = (Len(szoveg.Text) + 1) * 120 + 150
End Sub

Private Sub terulet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
szoveg.Visible = False
px = X
py = Y
    Select Case Button
        Case 2
            rename.Enabled = False
            del.Enabled = False
            replace.Enabled = False
            PopupMenu edit
            rename.Enabled = True
            del.Enabled = True
            replace.Enabled = True
    End Select
End Sub
'Public Sub tomorit(fajl As String)
'On Error GoTo hiba
'Shell eleres & "\rar.exe x " & fajl & " megnyitva", vbHide
'Do While Dir(eleres & "\megnyitva\szerk.vtk", vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) = ""
'Loop
'tolt (eleres & "\megnyitva\szerk.vtk")
'Exit Sub
'hiba:
'MsgBox "Nem találom a külsõ tömörítõt!"
'End Sub

Private Sub terulet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If athelyez = True Then
    pont(aktualis).Move X - 67, Y - 67
    fedo(aktualis).Move pont(aktualis).Left, pont(aktualis).Top
    cimke(aktualis).Top = pont(aktualis).Top - 30
    Call igazit(aktualis, cimke(aktualis).Alignment)
End If



sb.Panels(1).Text = "X=" & X
sb.Panels(2).Text = "Y=" & Y
End Sub
Public Sub alaphelyzet()
On Error Resume Next
    For i = 1 To teljes
        Unload cimke(i)
        Unload pont(i)
        Unload fedo(i)
    Next i
    teljes = 0
    terulet.Picture = Nothing
    aktualis = 0
    terkepneve = ""
    kepneve = ""
    Form_Load
End Sub
Private Sub elemmenu()
    uj.Enabled = False
            PopupMenu edit
    uj.Enabled = True
End Sub
Private Sub del_click()
'sb.Panels(4).Text = cimke(aktualis).Caption
'sb.Panels(6).Text = "Törlés"
i = MsgBox("Biztos törölni akarja a kijelölt elemet (" & cimke(aktualis).Caption & ")", vbYesNo + vbCritical, "Törlés megerõsítése")
    If i = vbYes Then
        Unload cimke(aktualis)
        Unload pont(aktualis)
        Unload fedo(aktualis)
    End If
'sb.Panels(4).Text = ""
'sb.Panels(6).Text = ""
aktualis = 0
szerkesztett = True
End Sub
Private Sub elem_click(Index As Integer)
px = px - 67
py = py - 67
    teljes = teljes + 1
                        Load pont(teljes)
                            pont(teljes).Left = px
                            pont(teljes).Top = py
                            pont(teljes).Visible = True
                        Load cimke(teljes)
                            cimke(teljes).Top = py - 30
                            cimke(teljes).Left = pont(teljes).Width + px + 15
                            cimke(teljes).Visible = True
                        Load fedo(teljes)
                            fedo(teljes).Move pont(teljes).Left, pont(teljes).Top, pont(teljes).Width, pont(teljes).Height
                            fedo(teljes).Visible = True
    Select Case Index
        Case 1
            pont(teljes).Shape = 3
            cimke(teljes).Caption = "Új város"
        Case 2
            pont(teljes).Shape = 1
            cimke(teljes).Caption = "Új terület"
    End Select
    aktualis = teljes
    beiro
End Sub
Private Sub props_click()
    tul.Show vbModal
End Sub

Private Sub tuls_Click()
aktualis = 0
props_click
End Sub
Public Sub mentes(fajlnev As String)
On Error GoTo atugrik
    '"masolas
    Open fajlnev For Output As 2
        Print #2, "cim=" & terkepneve
        Print #2, "terkep=" & kepneve
        For i = 1 To teljes
            Print #2, "dolt=" & Abs(CInt(cimke(i).FontItalic))
            Print #2, "alahuzott=" & Abs(CInt(cimke(i).FontUnderline))
            Print #2, "kover=" & Abs(CInt(cimke(i).FontBold))
            Print #2, "meret=" & cimke(i).FontSize
            Print #2, "szin=" & cimke(i).ForeColor
            Print #2, "jobbra=" & cimke(i).Alignment
            Print #2, "elem=" & pont(i).Shape & "," & pont(i).Left & "," & pont(i).Top & "," & cimke(i).Caption
atugrik:
        Next i
    
    Close 2
End Sub
Public Sub igazit(elem As Integer, zaras As Byte)
    Select Case zaras
        Case 0
            'ponttól jobbra
            cimke(elem).Left = pont(elem).Left + pont(elem).Width + 15
        Case 1
            'ponttól blra
            cimke(elem).Left = pont(elem).Left - (cimke(elem).Width + 15)
    End Select
cimke(elem).Alignment = zaras
End Sub
