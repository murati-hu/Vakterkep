VERSION 5.00
Begin VB.Form jelol 
   Caption         =   "Jelmagyar�zat"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape pont 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   -160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C�mke"
      Height          =   195
      Index           =   0
      Left            =   360
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   -160
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "jelol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 10 Then
    terkep.jelm_Click
End If
End Sub
