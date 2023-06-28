VERSION 5.00
Begin VB.Form nevjegy 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vaktérkép - Muráti Ákos"
   ClientHeight    =   5175
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "nevjegy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3571.877
   ScaleMode       =   0  'User
   ScaleWidth      =   4507.447
   ShowInTaskbar   =   0   'False
   Begin VB.Image kep 
      Height          =   3000
      Left            =   120
      Picture         =   "nevjegy.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label cimke 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fordítás:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   4485
   End
   Begin VB.Label cimke 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eredeti magyar nyelvû változat."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   4320
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   112.686
      X2              =   4282.074
      Y1              =   2650.437
      Y2              =   2650.437
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   585
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   465
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vakterkep.ini.hu"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   3240
      Width           =   2430
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "muratiakos@hotmail.com"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   3480
      Width           =   2115
   End
End
Attribute VB_Name = "nevjegy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cimke_Click(Index As Integer)
    Me.Hide
End Sub

Private Sub Form_Click()
    'Unload Me
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Caption = Vakterkep.Verzio & "." & App.Revision & " - Muráti Ákos"
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
    Me.Hide
End Sub

Private Sub tamogato_Click(Index As Integer)
    Me.Hide
End Sub

Private Sub url_Click(Index As Integer)
    Select Case Index
        Case 0
            Shell "explorer mailto:muratiakos@hotmail.com", vbMinimizedNoFocus
            'Shell "outlook -c IPM.Note /m murako@index.hu", vbNormalFocus
        Case 1
            Shell "explorer http://www.vakterkep.ini.hu"
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
