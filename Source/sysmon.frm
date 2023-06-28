VERSION 5.00
Begin VB.Form sysmon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rendszer monitor"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox monitor 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "sysmon.frx":0000
      Top             =   0
      Width           =   5775
   End
   Begin VB.PictureBox kep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3330
      Left            =   480
      Picture         =   "sysmon.frx":0020
      ScaleHeight     =   3300
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4530
   End
End
Attribute VB_Name = "sysmon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
terkep.kuldo.Enabled = False
End Sub
