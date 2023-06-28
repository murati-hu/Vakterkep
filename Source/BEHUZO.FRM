VERSION 5.00
Begin VB.Form behuzo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Behúzó"
   ClientHeight    =   105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   105
   ScaleWidth      =   915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label szoveg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jel szöveg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   -120
      TabIndex        =   0
      Top             =   -195
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "behuzo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 27 Then Me.Hide
End Sub
Private Sub szoveg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    szoveg(Index).Drag
End Sub
