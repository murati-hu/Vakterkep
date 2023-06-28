VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vaktérkép és Vaktérkép Szerkesztõ Névjegye"
   ClientHeight    =   4170
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4875
   ClipControls    =   0   'False
   Icon            =   "nevjegy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2878.208
   ScaleMode       =   0  'User
   ScaleWidth      =   4577.877
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox kep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   202
      Picture         =   "nevjegy.frx":8EDA
      ScaleHeight     =   3300
      ScaleWidth      =   4500
      TabIndex        =   1
      ToolTipText     =   "A kilépéshez kattints ide..."
      Top             =   120
      Width           =   4530
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muráti Ákos honlapja - www.extra.hu/murako"
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   3480
      Width           =   3195
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:  b0murako@gyakg.u-szeged.hu"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   3840
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    kep.Picture = sysmon.kep.Picture
    Me.Caption = App.Title & " névjegye"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
            Shell "explorer mailto:b0murako@gyakg.u-szeged.hu", vbHide
        Case 1
            Shell "explorer http://www.extra.hu/murako"
    End Select
End Sub

Private Sub url_Mousemove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To url.Count - 1
        url(i).FontUnderline = False
        url(i).ForeColor = vbBlack
Next i
    url(Index).ForeColor = vbBlue
    url(Index).FontUnderline = True
End Sub
