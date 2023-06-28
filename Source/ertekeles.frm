VERSION 5.00
Begin VB.Form ertekeles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Értékelés"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   ControlBox      =   0   'False
   Icon            =   "ertekeles.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   15
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   14
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label neve 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elégtelen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label maxpont 
      BackStyle       =   0  'Transparent
      Caption         =   "/ 0"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label pontok 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "Pontok:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label jegy 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "Érdemjegy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label szazalek 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "Eredmény:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label hibak 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label helyes 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "Találatok száma:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label cetli 
      BackStyle       =   0  'Transparent
      Caption         =   "Hibák száma:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "ertekeles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then HHSugo ("ert.htm")
End Sub



Private Sub ok_Click()
Me.Hide
End Sub
