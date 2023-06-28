VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form opciok 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vaktérkép beállításai"
   ClientHeight    =   3360
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5460
   ControlBox      =   0   'False
   Icon            =   "opciok.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame alt 
      Caption         =   "Általános beállítások:"
      Height          =   2775
      Left            =   1920
      TabIndex        =   20
      Top             =   0
      Width           =   3495
      Begin VB.CheckBox tippek 
         Caption         =   "Gyorstippek"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox segito 
         Caption         =   "Segítõ kérdések"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox jel 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "?"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox beeng 
         Caption         =   "Beállítások menüpont engedélyezése"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.CheckBox beall 
         Caption         =   "Projektfájlok egyéni beállításainak engedélyezése"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label cetli 
         Caption         =   "Pótlójel:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.CommandButton megse 
      Caption         =   "&Mégse"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Alk 
      Caption         =   "Menté&s"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame hat 
      Caption         =   "Értékelés:"
      Height          =   2775
      Left            =   1920
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox szazal 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "20"
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox pont 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "10"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   1
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "52"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   2
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "60"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   3
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "75"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Height          =   285
         Index           =   4
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "91"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   11
         Left            =   3120
         TabIndex        =   32
         Top             =   2220
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "Kérdésenként százalék levonás:"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   9
         Left            =   2640
         TabIndex        =   29
         Top             =   1380
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   28
         Top             =   1020
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   27
         Top             =   660
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "%"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   26
         Top             =   300
         Width           =   255
      End
      Begin VB.Label cetli 
         Caption         =   "Egy feladatra adható pont:"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label cetli 
         Caption         =   "Elégséges alsó határa:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Közepes alsó határa:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Jó alsó határa:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Példás alsó határa:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
   End
   Begin MSComctlLib.TreeView fa 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4895
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   14
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   13
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "opciok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Alk_Click()
Open eleres & "\vakterkep.ini" For Output As 3
    Print #3, "egyeni=" & beall.Value
    Print #3, "hatarok=" & hatarok(1) & "," & hatarok(2) & "," & hatarok(3) & "," & hatarok(4)
    Print #3, "beallitas=" & beeng.Value
    Print #3, "jel=" & opciok.jel
    Print #3, "pont=" & opciok.pont
    Print #3, "minusz=" & opciok.szazal
    Print #3, "segito=" & opciok.segito
    Print #3, "tippek=" & opciok.tippek
Close 3
End Sub

Private Sub fa_NodeClick(ByVal Node As MSComctlLib.Node)
alt.Visible = False
hat.Visible = False
Select Case Node.Key
    Case "alt"
        alt.Visible = True
    Case "szazal"
        hat.Visible = True
        
End Select
End Sub


Private Sub Form_Load()
fa.Nodes.Add , , "alt", "Általános"
fa.Nodes.Add , , "szazal", "Értékelés"
    
End Sub

Private Sub hatarok_LostFocus(Index As Integer)
On Error Resume Next
    For i = 1 To 3
        If IsNumeric(hatarok(i)) = False Or hatarok(i) < 0 Then
            MsgBox "Ide csak pozitív egész számot adhat meg!", vbInformation, i + 1 & "-s alsó határa:"
            hatarok(i) = hatarok(i) = hatarok(i + 1) - 1
            Exit Sub
        End If
        
        If hatarok(i) > hatarok(i + 1) Then
            MsgBox "A megadott százaléknak kisebbnek kell lennie a az utána következõnél!", vbInformation, i + 1 & "-s alsó határa"
            hatarok(i) = hatarok(i + 1) - 1
        End If
    Next i
End Sub

Private Sub megse_Click()
Call terkep.tolt(eleres & "\vakterkep.ini")
Me.Hide
End Sub

Private Sub ok_Click()
Me.Hide
End Sub


Private Sub pont_LostFocus()
On Error Resume Next
If IsNumeric(pont.Text) = False Then
     MsgBox "Ide csak szám kerülhet!", vbInformation, "Pontok:"
    pont.Text = 10
End If
End Sub

Private Sub szazal_LostFocus()
On Error Resume Next
If IsNumeric(szazal.Text) = False Or szazal.Text < 1 Or szazal.Text > 20 Then
    MsgBox "Ide csak 1 és 20 közé esõ egész szám kerülhet.", vbInformation, "Kérdésenkénti levonás:"
    szazal.Text = 20
End If
End Sub
