VERSION 5.00
Begin VB.Form teszter 
   Caption         =   "Hahó"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin szax.szoveg szoveg1 
      Height          =   210
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   765
      _ExtentX        =   1032
      _ExtentY        =   1429
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   3015
   End
End
Attribute VB_Name = "teszter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub szoveg1_Click()
    
End Sub
