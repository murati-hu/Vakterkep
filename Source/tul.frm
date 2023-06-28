VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tul 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tulajdonságok"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton megse 
      Cancel          =   -1  'True
      Caption         =   "&Mégse"
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   24
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame tertul 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   4920
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton szerkeszto 
         Caption         =   "Kép szerkesztése"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CommandButton talloz 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox kep 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox terkep 
         Height          =   285
         Left            =   840
         TabIndex        =   18
         Text            =   "Névtelen"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label duma 
         Caption         =   "Térkép háttere:"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label duma 
         Caption         =   "Cím:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame elem 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox jobbra 
         Caption         =   "Szöveg a bal oldalon"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox meret 
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "8"
         Top             =   2880
         Width           =   615
      End
      Begin VB.Frame mintafr 
         Caption         =   "Minta:"
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
         Begin VB.Shape pont 
            BorderColor     =   &H00000000&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Shape           =   3  'Circle
            Top             =   480
            Width           =   135
         End
         Begin VB.Label cimke 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minta"
            Height          =   195
            Left            =   250
            MousePointer    =   1  'Arrow
            TabIndex        =   15
            Top             =   450
            Width           =   390
         End
      End
      Begin VB.CommandButton szin 
         Caption         =   "Szín..."
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox alahuzott 
         Caption         =   "Aláhúzott"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox dolt 
         Caption         =   "Dõlt"
         Height          =   195
         Left            =   1800
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox felkover 
         Caption         =   "Félkövér"
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox py 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox px 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox tip 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox nev 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Text            =   "Névtelen"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label duma 
         Caption         =   "Méret:"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   23
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label duma 
         Caption         =   "Y="
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label duma 
         Caption         =   "Pozíció:  X="
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label duma 
         Caption         =   "Típus:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.Label duma 
         Caption         =   "Név:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Általános"
            Key             =   "elem"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "tul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub alahuzott_Click()
cimke.FontUnderline = alahuzott.Value
End Sub

Private Sub dolt_Click()
cimke.FontItalic = dolt.Value
End Sub






Private Sub felkover_Click()
cimke.FontBold = felkover.Value
End Sub

Private Sub Form_Load()
tip.AddItem "Város"
tip.AddItem "Terület"
Me.Move Me.Left, Me.Top, 4845, 4845
tertul.Move 240, 480
elem.Move 240, 480
terkep.Text = szerk.terkepneve
kep.Text = szerk.kepneve

If szerk.aktualis = 0 Then
        Me.Caption = "Térkép tulajdonságai"
        tertul.Visible = True
    Else
        Me.Caption = szerk.cimke(szerk.aktualis).Caption & " tulajdonságai"
        elem.Visible = True
        jobbra.Value = Abs(CInt(szerk.cimke(szerk.aktualis).Alignment))
        nev.Text = szerk.cimke(szerk.aktualis).Caption
        Select Case szerk.pont(szerk.aktualis).Shape
            Case 3
                tip.Text = tip.List(0)
            Case 1
                tip.Text = tip.List(1)
        End Select
        px.Text = szerk.pont(szerk.aktualis).Left + 67
        py.Text = szerk.pont(szerk.aktualis).Top + 67
        pont.Shape = szerk.pont(szerk.aktualis).Shape
        pont.BackColor = szerk.pont(szerk.aktualis).BackColor
        pont.FillColor = szerk.pont(szerk.aktualis).FillColor
        pont.BorderColor = szerk.pont(szerk.aktualis).BorderColor
        cimke.FontBold = szerk.cimke(szerk.aktualis).FontBold
        cimke.FontItalic = szerk.cimke(szerk.aktualis).FontItalic
        cimke.FontUnderline = szerk.cimke(szerk.aktualis).FontUnderline
        cimke.FontSize = szerk.cimke(szerk.aktualis).FontSize
        cimke.Caption = szerk.cimke(szerk.aktualis).Caption
        cimke.ForeColor = szerk.cimke(szerk.aktualis).ForeColor
        felkover.Value = Abs(CInt(cimke.FontBold))
        dolt.Value = Abs(CInt(cimke.FontItalic))
        alahuzott.Value = Abs(CInt(cimke.FontUnderline))
        meret.Text = szerk.cimke(szerk.aktualis).FontSize
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
szerk.aktualis = 0
tertul.Visible = False
elem.Visible = False
End Sub

Private Sub jobbra_Click()
    Select Case jobbra.Value
        Case 1
            pont.Move 1320, 480
            cimke.Left = pont.Left - (cimke.Width + 15)
        Case 0
            pont.Move 120, 480
            cimke.Left = pont.Left + pont.Width + 15
    End Select
End Sub

Private Sub megse_Click()
Unload Me
End Sub

Private Sub nev_Change()
cimke.Caption = nev.Text
End Sub

Private Sub ok_Click()
    szerk.kepneve = kep.Text
    szerk.terkepneve = terkep.Text
    szerk.Caption = szerk.terkepneve & " - Vaktérkép Szerkesztõ " & App.Major & "." & App.Minor
    Call szerk.igazit(szerk.aktualis, jobbra.Value)
    
    
    szerk.cimke(szerk.aktualis).Caption = nev.Text
    szerk.cimke(szerk.aktualis).ForeColor = cimke.ForeColor
    szerk.cimke(szerk.aktualis).FontBold = cimke.FontBold
    szerk.cimke(szerk.aktualis).FontItalic = cimke.FontItalic
    szerk.cimke(szerk.aktualis).FontUnderline = cimke.FontUnderline
    szerk.cimke(szerk.aktualis).FontSize = cimke.FontSize
    szerk.pont(szerk.aktualis).FillColor = pont.FillColor
    szerk.pont(szerk.aktualis).BackColor = pont.BackColor
    szerk.pont(szerk.aktualis).BorderColor = pont.BorderColor
    szerk.pont(szerk.aktualis).Shape = pont.Shape
    
    'szerk.pont(szerk.aktualis).Move
    szerk.szerkesztett = True
 Unload Me
End Sub

Private Sub szerkeszto_Click()
On Error GoTo hiba
Shell "pbrush.exe ", vbNormalFocus
Exit Sub
hiba:
    MsgBox "A paint nem nyitható meg!", vbCritical, "Hiba:"
End Sub

Private Sub szin_Click()
On Error GoTo megse
szerk.pb.ShowColor
pont.FillColor = szerk.pb.Color
pont.BackColor = szerk.pb.Color
pont.BorderColor = szerk.pb.Color
cimke.ForeColor = szerk.pb.Color
megse:
End Sub

Private Sub talloz_Click()
 szerk.picopen
End Sub

Private Sub meret_Change()
On Error Resume Next
cimke.FontSize = CInt(meret.Text)
End Sub


Private Sub tip_Click()
Select Case tip.ListIndex
    Case 0
        pont.Shape = 3
    Case 1
        pont.Shape = 1
End Select
  
End Sub
