VERSION 5.00
Object = "{7B12C111-305A-4D91-915D-4A35973E9B52}#1.0#0"; "Szoveg_ax.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin szax.szoveg szoveg1 
      Height          =   1725
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      _extentx        =   3625
      _extenty        =   3043
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    szoveg1.BorderStyle = Check1.Value
End Sub

Private Sub Check2_Click()
    szoveg1.BackStyle = Check2.Value
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move X, Y
End Sub

Private Sub szoveg1_Click()
    MsgBox "klikk"
    
End Sub

Private Sub szoveg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    szoveg1.Drag
End Sub

Private Sub Text1_Change()
    szoveg1.Caption = Text1.Text
End Sub

Private Sub Text2_Change()
On Error Resume Next
    szoveg1.Forgatas = Text2.Text
End Sub
