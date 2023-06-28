VERSION 5.00
Begin VB.Form beallitasok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beállítások"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton sugo 
      Caption         =   "Súgó"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox alap_frm 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   1275
      TabIndex        =   29
      Top             =   0
      Width           =   1335
      Begin VB.Label beall_mnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Értékelés"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   660
      End
      Begin VB.Label beall_mnu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Általános"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   645
      End
   End
   Begin VB.Frame beall_frm 
      BorderStyle     =   0  'None
      Caption         =   "Értékelés"
      Height          =   2415
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox levonas 
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "10"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox pont 
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "10"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "52"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "60"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "75"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   4
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "91"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   28
         Top             =   1980
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "Kérdésenként százalék levonás:"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   9
         Left            =   2400
         TabIndex        =   26
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   25
         Top             =   900
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   24
         Top             =   540
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   23
         Top             =   180
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "Egy feladatra adható pont:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label cetli 
         Caption         =   "Elégséges alsó határa:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Közepes alsó határa:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Jó alsó határa:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Példás alsó határa:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
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
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox behuzas 
         Caption         =   "Behúzások engedélyezése"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox egyeni 
         Caption         =   "Projektfájlok egyéni beállításainak engedélyezése"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox enged 
         Caption         =   "Beállítások menüpont engedélyezése"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.TextBox jel 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "?"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox kerdesek 
         Caption         =   "Segítõ kérdések"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox tippek 
         Caption         =   "Gyorstippek"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.Label cetli 
         Caption         =   "Pótlójel:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
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
Dim seged
    Open Atalakit("$vt\vakterkep.ini") For Output As 3
        Print #3, "[Vaktérkép " & Vakterkep.Verzio & "]"
        Print #3, ""
        Print #3, "Beallitasok_Engedelyezese=" & enged.Value
        Print #3, "Egyeni_Beallitasok=" & egyeni.Value
        Print #3, "Ponthatarok=" & hatarok(1).Text & ","; hatarok(2).Text & ","; hatarok(3).Text & ","; hatarok(4).Text
        Print #3, "Helyettesito_Szoveg=" & jel.Text
        Print #3, "Pont=" & pont.Text
        Print #3, "Levonasok=" & levonas.Text
        Print #3, "Tippek_Engedelyezese=" & tippek.Value
        Print #3, "Kerdesek_Engedelyezese=" & kerdesek.Value
        Print #3, "Behuzasok_Engedelyezese=" & behuzas.Value
        Print #3, "Szerkesztes_engedelyezese=" & szerkesztes.Value
    Close 3
    If enged.Value = 0 Then
        MsgBox "Ön letiltotta a Beállítások menüt. Ahhoz, hogy újra el tudja érni a beállításokat, indítsa a programot a -beall kapcsolóval az alábbi módon: 'vakterkep.exe -beall'.", vbExclamation, "Beállítások"
    End If
    terkep.megnyitas ("$vt\vakterkep.ini")
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


Private Sub Form_Load()
Dim i As Integer
On Error Resume Next
    Me.Width = 4965
    Me.Height = 3630
    For i = 0 To beall_frm.Count - 1
        beall_frm(i).Move 1440, 0
    Next i
    beall_mnu_Click (0)
End Sub

Private Sub megse_Click()
    terkep.megnyitas ("$vt\vakterkep.ini")
    ok_Click
End Sub


Private Sub ok_Click()
    If terkep.Megnyitva <> "" Then
        'Call MsgBox("Az új beállítások érvénybelépéséhez újratöltöm a megnyitott projektet.", vbInformation, "Új beállítások érvényesítése...")
        terkep.Ujratolt (terkep.Megnyitva)
    End If
    Me.Hide
End Sub

Private Sub sugo_Click()
    HHSugo ("beall.htm")
End Sub
