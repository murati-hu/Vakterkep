VERSION 5.00
Begin VB.Form nevjegy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vaktérkép névjegye - Muráti Ákos"
   ClientHeight    =   3615
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4515
   ClipControls    =   0   'False
   Icon            =   "nevjegy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2495.137
   ScaleMode       =   0  'User
   ScaleWidth      =   4239.818
   ShowInTaskbar   =   0   'False
   Begin VB.Image kep 
      Height          =   3000
      Left            =   0
      Picture         =   "nevjegy.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vaktérkép 2.0 Weboldalának megtekintése"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   3675
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail küldése"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   3360
      Width           =   1245
   End
End
Attribute VB_Name = "nevjegy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Me.Caption = "Vaktérkép névjegye - Muráti Ákos"
    Me.Icon = terkep.Icon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
    For i = 0 To url.Count - 1
        url(i).FontUnderline = False
        url(i).ForeColor = vbBlack
    Next i
End Sub

Private Sub kep_Click()
    Unload Me
End Sub

Private Sub url_Click(Index As Integer)
    Select Case Index
        Case 0
            Shell "explorer mailto:b0murako@gyakg.u-szeged.hu", vbMinimizedNoFocus
            'Shell "outlook -c IPM.Note /m murako@index.hu", vbNormalFocus
        Case 1
            Shell "explorer http://www.tar.hu/vakterkep2002"
    End Select
End Sub

Private Sub url_Mousemove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To url.Count - 1
        url(i).FontUnderline = False
        url(i).ForeColor = vbBlack
Next i
    url(Index).ForeColor = vbBlue
    url(Index).FontUnderline = True
End Sub
