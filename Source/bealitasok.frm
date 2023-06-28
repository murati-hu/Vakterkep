VERSION 5.00
Begin VB.Form beallitasok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beállítások"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame beall_frm 
      BorderStyle     =   0  'None
      Caption         =   "Javítás színekkel"
      Height          =   2415
      Index           =   3
      Left            =   6480
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.FileListBox nyelvek 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   810
         Left            =   360
         Pattern         =   "*.lng"
         TabIndex        =   35
         Top             =   1080
         Width           =   2775
      End
      Begin VB.OptionButton nyelv 
         Caption         =   "Idegen nyelvi modul alkalmazása"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   3375
      End
      Begin VB.OptionButton nyelv 
         Caption         =   "Eredeti magyar nyelv használata"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.CommandButton betoltnyelv 
         Caption         =   "&Nyelv alkalmazása"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   2040
         Width           =   2775
      End
   End
   Begin VB.Frame beall_frm 
      BorderStyle     =   0  'None
      Caption         =   "Javítás színekkel"
      Height          =   2415
      Index           =   2
      Left            =   5160
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.OptionButton megoldas 
         Caption         =   "Csak a helytelenek maradnak"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   3275
      End
      Begin VB.OptionButton megoldas 
         Caption         =   "Javítás színekkel"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Value           =   -1  'True
         Width           =   3275
      End
      Begin VB.CheckBox szokoz 
         Caption         =   "Felesleges szóközöket figyelmen kívül hagy"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Value           =   1  'Checked
         Width           =   3275
      End
      Begin VB.CheckBox kisbetus 
         Caption         =   "Kis és nagy betû nem számít"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Value           =   1  'Checked
         Width           =   3275
      End
      Begin VB.CheckBox nincskotojel 
         Caption         =   "Kötõjelek figyelmenkívül hagyása"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   3275
      End
   End
   Begin VB.CommandButton sugo 
      Caption         =   "Súgó"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox alap_frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   1275
      TabIndex        =   17
      Top             =   0
      Width           =   1335
      Begin VB.Label beall_mnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nyelv"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   405
      End
      Begin VB.Label beall_mnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ellenõrzés"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   720
      End
      Begin VB.Label beall_mnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Értékelés"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   660
      End
      Begin VB.Label beall_mnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Általános"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.Frame beall_frm 
      BorderStyle     =   0  'None
      Caption         =   "Értékelés"
      Height          =   2415
      Index           =   1
      Left            =   5760
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox ert_jegy 
         Height          =   285
         Left            =   2640
         TabIndex        =   41
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox ert_szoveg 
         Height          =   285
         Left            =   1440
         TabIndex        =   40
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox ert_szazalek 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "10"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton torol 
         Caption         =   "Töröl"
         Height          =   255
         Left            =   2400
         TabIndex        =   38
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton felvesz 
         Caption         =   "Felvesz"
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   840
         Width           =   855
      End
      Begin VB.ListBox ert_hatarok 
         Height          =   840
         Left            =   0
         TabIndex        =   36
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox levonas 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "10"
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox pont 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "10"
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label cetli 
         AutoSize        =   -1  'True
         Caption         =   "%, akkor"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   44
         Top             =   165
         Width           =   615
      End
      Begin VB.Label cetli 
         AutoSize        =   -1  'True
         Caption         =   "Ha az eredmény <="
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label cetli 
         AutoSize        =   -1  'True
         Caption         =   "Megnevezés:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   42
         Top             =   480
         Width           =   960
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   11
         Left            =   3120
         TabIndex        =   16
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label cetli 
         AutoSize        =   -1  'True
         Caption         =   "Kérdésenként százalék levonás:"
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   15
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label cetli 
         AutoSize        =   -1  'True
         Caption         =   "Egy feladatra adható pont:"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   14
         Top             =   1800
         Width           =   1875
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Frame beall_frm 
      BorderStyle     =   0  'None
      Caption         =   "Általános beállítások"
      Height          =   2415
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   3375
      Begin VB.CheckBox szerkesztes 
         Caption         =   "Szerkesztés menü engedélyezése"
         Height          =   375
         Left            =   0
         TabIndex        =   22
         Top             =   480
         Width           =   3275
      End
      Begin VB.CheckBox behuzas 
         Caption         =   "Behúzások engedélyezése"
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   960
         Value           =   1  'Checked
         Width           =   3275
      End
      Begin VB.CheckBox egyeni 
         Caption         =   "Projektfájlok egyéni beállításainak engedélyezése"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   3275
      End
      Begin VB.CheckBox enged 
         Caption         =   "Beállítások menüpont engedélyezése"
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Value           =   1  'Checked
         Width           =   3275
      End
      Begin VB.TextBox jel 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Text            =   "?"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CheckBox kerdesek 
         Caption         =   "Segítõ kérdések"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3275
      End
      Begin VB.CheckBox tippek 
         Caption         =   "Gyorstippek"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3275
      End
      Begin VB.Label cetli 
         AutoSize        =   -1  'True
         Caption         =   "Pótlójel:"
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   9
         Top             =   2160
         Width           =   555
      End
   End
   Begin VB.CommandButton megse 
      Caption         =   "&Mégse"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton alkalmaz 
      Caption         =   "Menté&s"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "beallitasok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub alap_frm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    beall_mnu_MouseMove beall_mnu.Count + 1, Button, Shift, X, Y
End Sub

Private Sub alkalmaz_Click()
On Error GoTo Hiba
Dim seged, i As Integer

    Open Atalakit("$vt\vakablak.ini") For Output As 3
        Print #3, "[" & Vakterkep.Verzio & "]"
        Print #3, ""
        Print #3, "Beallitasok_Engedelyezese=" & enged.Value
        Print #3, "Egyeni_Beallitasok=" & egyeni.Value
        'Print #3, "Ponthatarok=" & hatarok(1).Text & ","; hatarok(2).Text & ","; hatarok(3).Text & ","; hatarok(4).Text
        Print #3, "Helyettesito_Szoveg=" & jel.Text
        Print #3, "Pont=" & pont.Text
        Print #3, "Levonasok=" & levonas.Text
        Print #3, "Tippek_Engedelyezese=" & tippek.Value
        Print #3, "Kerdesek_Engedelyezese=" & kerdesek.Value
        Print #3, "Behuzasok_Engedelyezese=" & behuzas.Value
        Print #3, "Szerkesztes_engedelyezese=" & szerkesztes.Value
        If nyelv(1).Value Then Print #3, "Nyelv=" & nyelvek.List(nyelvek.ListIndex)
                
    Close 3
    
    Open Atalakit("$vt\ertekeles.ini") For Output As 5
        For i = 0 To ert_hatarok.ListCount - 1
            Print #5, ert_hatarok.List(i)
        Next i
    Close 5
    
    If enged.Value = 0 Then
        MsgBox kozos.KozosSzovegek(1), vbExclamation, Me.Caption
    End If
    'terkep.megnyitas ("$vt\vakablak.ini")
    Szulo.megnyitas ("$vt\vakablak.ini")
    Exit Sub
Hiba:
    KozosHibak Err.Number
    Close 3
End Sub

Private Sub beall_mnu_Click(Index As Integer)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To beall_frm.Count - 1
        beall_frm(i).Visible = False
        beall_mnu(i).FontBold = False
    Next i
    beall_frm(Index).Visible = True
    beall_mnu(Index).FontBold = True
End Sub

Private Sub beall_mnu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To beall_mnu.Count - 1
        beall_mnu(i).ForeColor = vbBlack
    Next i
    beall_mnu(Index).ForeColor = vbBlue
End Sub


Private Sub betoltnyelv_Click()
    If nyelv(0).Value Then
            magyar_nyelv
        Else
            NyelvAlkalmazasa (nyelvek.List(nyelvek.ListIndex))
    End If
End Sub

Private Sub felvesz_Click()
Dim Sorszam As Integer, sor As String
    
    sor = ert_szazalek.Text & "%>=" & ert_szoveg.Text & "=" & ert_jegy.Text
    ert_hatarok.AddItem sor
    
    Sorszam = ert_hatarok.ListCount - 1
    Do While Kisebb(Sorszam)
        ert_hatarok.RemoveItem (Sorszam)
        Sorszam = Sorszam - 1
        ert_hatarok.AddItem sor, Sorszam
    Loop
    ert_hatarok.Selected(Sorszam) = True
    
    ert_szazalek.Text = ""
    ert_szoveg.Text = ""
    ert_jegy.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            megse_Click
        Case 112
            sugo_Click
    End Select
End Sub


Private Sub Form_Load()
Dim i As Integer
On Error Resume Next
    Me.Width = 4965
    Me.Height = 3630
    For i = 0 To beall_frm.Count - 1
        beall_frm(i).Move 1440, 0
    Next i
    beall_mnu_Click (0)
    nyelvek.Path = Vakterkep.Konyvtar & "nyelvek\"

    BetoltPonthatarok
End Sub

Private Sub megse_Click()
    'terkep.megnyitas ("$vt\vakablak.ini")
    Szulo.megnyitas ("$vt\vakablak.ini")
    ok_Click
End Sub


Private Sub nyelv_Click(Index As Integer)
    nyelvek.Enabled = nyelv(1).Value
End Sub

Private Sub ok_Click()
    If Szulo.Name = "terkep" And terkep.Megnyitva <> "" Then
        If MsgBox(kozos.KozosSzovegek(2), vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            terkep.Ujratolt (terkep.Megnyitva)
        End If
    End If
    Me.Hide
End Sub

Private Sub sugo_Click()
    HHSugo ("beall.htm")
End Sub
Private Function Kisebb(Sorszam As Integer) As Boolean
Dim elozo As Byte, jelenlegi As Byte
Dim seged As String
    If Sorszam > 0 Then
            'seged = Utasitas(ert_hatarok.List(Sorszam - 1))
            'elozo = CByte(Mid(seged, 1, Len(seged) - 2))
            elozo = CByte(KiErtekeles(Sorszam - 1, 1))
            
            'seged = Utasitas(ert_hatarok.List(Sorszam))
            jelenlegi = CByte(KiErtekeles(Sorszam, 1))
            
            If jelenlegi < elozo Then
                    Kisebb = True
                Else
                    Kisebb = False
            End If
        Else
            Kisebb = False
    End If
End Function

Private Sub torol_Click()
On Error Resume Next
    ert_hatarok.RemoveItem (ert_hatarok.ListIndex)
End Sub
Public Function KiErtekeles(Sorszam As Integer, Melyik As Byte) As String
Dim seged As String
On Error GoTo Hiba

    Select Case Melyik
        Case 1
            seged = Utasitas(ert_hatarok.List(Sorszam))
            KiErtekeles = Mid(seged, 1, Len(seged) - 2)
        Case 2
            seged = Ertek(ert_hatarok.List(Sorszam))
            KiErtekeles = Utasitas(seged)
        Case 3
            seged = Ertek(ert_hatarok.List(Sorszam))
            KiErtekeles = Ertek(seged)
    End Select
Exit Function
Hiba:
    MsgBox Err.Description
End Function
Private Sub BetoltPonthatarok()
Dim sor As String
On Error GoTo Hiba
     Open Atalakit("$vt\ertekeles.ini") For Input As 5
        Do While Not EOF(5)
            Line Input #5, sor
            ert_hatarok.AddItem sor
        Loop
    Close 5
Exit Sub
Hiba:
    'KozosHibak Err.Number
    Close 5
End Sub
