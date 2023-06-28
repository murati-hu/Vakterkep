VERSION 5.00
Begin VB.Form jelmagyarazat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   Caption         =   "Jelmagyarázat"
   ClientHeight    =   165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1485
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   165
   ScaleWidth      =   1485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Vakterkep2.jel jelm 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   -255
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   873
      _ExtentY        =   873
      KitoltesSzine   =   -2147483640
      KeretSzine      =   -2147483640
      HatterSzine     =   -2147483643
   End
   Begin VB.Label jelm_szoveg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jel szöveg"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   -250
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "jelmagyarazat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then Me.Hide
End Sub
