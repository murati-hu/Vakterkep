VERSION 5.00
Begin VB.Form koszonet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Köszönet"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox lista 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "koszonet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Meg(Kit As String)
    lista.Text = lista.Text & Kit & vbCrLf
End Sub

Private Sub Form_Load()
    lista.Text = ""
    Meg "Beke Ferenc"
    Meg "C3 Kulturális és Kommunikációs Központ"
    Meg "Farkas Zoltán"
    Meg "Fodor Zsolt"
    Meg "Ravasz Viktória"
    Meg "Vida Csaba"
End Sub

Private Sub ok_Click()
    Me.Hide
End Sub
