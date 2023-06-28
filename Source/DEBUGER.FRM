VERSION 5.00
Begin VB.Form debuger 
   Caption         =   "Rendszer ellenõrzõ"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox egyeb 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox szovegek 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox jelek 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   120
   End
   Begin VB.Label Label2 
      Caption         =   "jelek:"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "szövegegk::"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "debuger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dim kitol As Object
If terkep.Visible = True Then
    jelek.Text = terkep.jel.Count - 1 + terkep.megj.Count - 1
    szovegek.Text = terkep.jel_szoveg.Count - 1 + terkep.megj_szoveg.Count - 1
Else
    jelek.Text = szerkeszto.jel.Count - 1
    szovegek.Text = szerkeszto.jel_szoveg.Count - 1
End If


End Sub
