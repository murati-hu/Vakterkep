VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tul 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tulajdonságok"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton sugo 
      Caption         =   "Sú&gó"
      Height          =   375
      Left            =   3240
      TabIndex        =   44
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame seg 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   240
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox megold 
         Height          =   285
         Left            =   840
         TabIndex        =   38
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox segitseg 
         Height          =   1935
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   960
         Width           =   4215
      End
      Begin VB.ComboBox szama 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label cetli 
         Caption         =   "Megoldás:"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   39
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label cetli 
         Caption         =   "Kérdés szövege:"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   37
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label cetli 
         Caption         =   "Kérdés száma:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton megse 
      Cancel          =   -1  'True
      Caption         =   "&Mégse"
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame tertul 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox ponthat 
         Caption         =   "Egyéni százalékhatárok erre a térképre"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox hatarok 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "91"
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "75"
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "60"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox hatarok 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "52"
         Top             =   1680
         Width           =   375
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
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox kep 
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox terkep 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Text            =   "Névtelen"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label cetli 
         Caption         =   "Példás alsó határa:"
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   30
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Jó alsó határa:"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   29
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Közepes alsó határa:"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label cetli 
         Caption         =   "Elégséges alsó határa:"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label duma 
         Caption         =   "Térkép háttere:"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label duma 
         Caption         =   "Cím:"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame elem 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CheckBox lathatatlan 
         Caption         =   "Láthatatlan jel"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Formátum másoló"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3000
         Width           =   3975
      End
      Begin VB.TextBox gyorstip 
         Height          =   615
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox jelmagy 
         Caption         =   "Jelmagyarázat"
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox jobbra 
         Caption         =   "Szöveg a bal oldalon"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox meret 
         Height          =   285
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "8"
         Top             =   2640
         Width           =   495
      End
      Begin VB.Frame mintafr 
         Caption         =   "Minta:"
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
         Begin VB.Shape pont 
            BorderColor     =   &H00000000&
            BorderStyle     =   0  'Transparent
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   120
            Shape           =   3  'Circle
            Top             =   360
            Width           =   135
         End
         Begin VB.Label cimke 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minta"
            Height          =   195
            Left            =   255
            MousePointer    =   1  'Arrow
            TabIndex        =   11
            Top             =   330
            Width           =   390
         End
      End
      Begin VB.CommandButton szin 
         Caption         =   "Szín..."
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox alahuzott 
         Caption         =   "Aláhúzott"
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CheckBox dolt 
         Caption         =   "Dõlt"
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox felkover 
         Caption         =   "Félkövér"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox tip 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox nev 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "Névtelen"
         Top             =   0
         Width           =   3255
      End
      Begin VB.Label duma 
         Caption         =   "Tipp:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   495
      End
      Begin VB.Label duma 
         Caption         =   "Méret:"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label duma 
         Caption         =   "Típus:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label duma 
         Caption         =   "Név:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComctlLib.TabStrip ful 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Általános"
            Key             =   "alt"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Kérdések"
            Key             =   "segit"
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
Option Explicit
Dim segito(1 To 5) As String, megoldas(1 To 5) As String


Private Sub alahuzott_Click()
cimke.FontUnderline = alahuzott.Value
End Sub

Private Sub Command1_Click()
szerk.koppint = True
tul.Visible = False
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

For i = 1 To 5
    szama.AddItem CStr(i)
Next i
szama.Text = szama.List(0)

Me.Move Me.Left, Me.Top, 4845, 4845
tertul.Move 240, 480
elem.Move 240, 480
seg.Move 240, 480

terkep.Text = szerk.terkepneve
kep.Text = szerk.kepneve

If szerk.aktualis = 0 Then
        Me.Caption = "Térkép tulajdonságai"
        tertul.Visible = True
        For i = 1 To 4
            tul.hatarok(i) = opciok.hatarok(i)
        Next i
        ponthat.Value = opciok.beall.Value
        ful.Tabs.Remove (2)
    Else
        segitseg.Text = segito(1)
        megold.Text = megoldas(1)
        
        Me.Caption = szerk.cimke(szerk.aktualis).Caption & " tulajdonságai"
        elem.Visible = True
        formatuma (szerk.aktualis)
        'jobbra.Value = Abs(CInt(szerk.cimke(szerk.aktualis).Alignment))
        nev.Text = szerk.cimke(szerk.aktualis).Caption
        'Select Case szerk.pont(szerk.aktualis).Shape
         '   Case 3
          '      tip.Text = tip.List(0)
           ' Case 1
            '    tip.Text = tip.List(1)
        'End Select
        jelmagy.Value = szerk.cimke(szerk.aktualis).BorderStyle
     
        'pont.Shape = szerk.pont(szerk.aktualis).Shape
        'pont.BackColor = szerk.pont(szerk.aktualis).BackColor
        'pont.FillColor = szerk.pont(szerk.aktualis).FillColor
        'pont.BorderColor = szerk.pont(szerk.aktualis).BorderColor
        'If szerk.cimke(szerk.aktualis).Alignment = 0 Then
        '    cimke.Left = pont.Left + pont.Width + 15
        '    cimke.Alignment = 0
        'Else
        '    cimke.Alignment = 1
        '    cimke.Left = pont.Left - (cimke.Width + 15)
        'End If
        'cimke.FontBold = szerk.cimke(szerk.aktualis).FontBold
        'cimke.FontItalic = szerk.cimke(szerk.aktualis).FontItalic
        'cimke.FontUnderline = szerk.cimke(szerk.aktualis).FontUnderline
        'cimke.FontSize = szerk.cimke(szerk.aktualis).FontSize
        cimke.Caption = szerk.cimke(szerk.aktualis).Caption
        'cimke.ForeColor = szerk.cimke(szerk.aktualis).ForeColor
        'felkover.Value = Abs(CInt(cimke.FontBold))
        'dolt.Value = Abs(CInt(cimke.FontItalic))
        'alahuzott.Value = Abs(CInt(cimke.FontUnderline))
        'meret.Text = szerk.cimke(szerk.aktualis).FontSize
        gyorstip.Text = szerk.cimke(szerk.aktualis).ToolTipText
End If
End Sub




Private Sub ful_Click()
tertul.Visible = False
elem.Visible = False
seg.Visible = False

Select Case ful.SelectedItem.Index
    Case 1
        If szerk.aktualis = 0 Then
            tertul.Visible = True
        Else
            elem.Visible = True
        End If
    Case 2
        seg.Visible = True
End Select
End Sub


Private Sub hatarok_LostFocus(Index As Integer)
On Error Resume Next
    For i = 1 To 3
        If IsNumeric(hatarok(i)) = False Or hatarok(i) < 0 Then
            MsgBox "Ide csak pozitív egész számot adhat meg!", vbInformation, i + 1 & "-s alsó határa:"
            hatarok(i) = hatarok(i) = hatarok(i + 1) - 1
            Exit Sub
        End If
        
        If hatarok(i) > hatarok(i + 1) Then
            MsgBox "A megadott százaléknak kisebbnek kell lennie a az utána következõnél!", vbInformation, i + 1 & "-s alsó határa"
            hatarok(i) = hatarok(i + 1) - 1
        End If
    Next i
End Sub

Private Sub jobbra_Click()
    Select Case jobbra.Value
        Case 1
            pont.Move 1320, 360
            cimke.Left = pont.Left - (cimke.Width + 15)
        Case 0
            pont.Move 120, 360
            cimke.Left = pont.Left + pont.Width + 15
    End Select
End Sub

Private Sub lathatatlan_Click()
  pont.Visible = Abs(lathatatlan.Value - 1)
End Sub

Private Sub megold_LostFocus()
elment
End Sub

Private Sub megse_Click()
Unload Me
End Sub

Private Sub nev_Change()
cimke.Caption = nev.Text
End Sub

Private Sub ok_Click()
    szerk.kepneve = kep.Text
    If terkep = "" Then
            szerk.terkepneve = "Névtelen"
        Else
            szerk.terkepneve = terkep.Text
    End If
    szerk.Caption = szerk.terkepneve & " - Vaktérkép Szerkesztõ " & App.Major & "." & App.Minor
    If szerk.aktualis <> 0 Then
    Call szerk.igazit(szerk.aktualis, jobbra.Value)
    
    szerk.cimke(szerk.aktualis).ToolTipText = gyorstip.Text
    szerk.cimke(szerk.aktualis).BorderStyle = jelmagy.Value
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
    szerk.pont(szerk.aktualis).Visible = pont.Visible
    
    
    'szerk.pont(szerk.aktualis).Move
    k = ""
    For i = 1 To 5
        If segito(i) = "" Then segito(i) = " "
        k = k & segito(i) & "||"
'        MsgBox k
    Next i
    szerk.segitsegford (k)
    
    
    k = ""
    For i = 1 To 5
        If megoldas(i) = "" Then megoldas(i) = " "
        k = k & megoldas(i) & "||"
        'MsgBox k
    Next i
    szerk.megoldasford (k)
    
    End If
    For i = 1 To 4
            opciok.hatarok(i) = tul.hatarok(i)
    Next i
    opciok.beall.Value = ponthat.Value
    szerk.szerkesztett = True
 Unload Me
 'Me.Hide
End Sub

Private Sub szerkeszto_Click()
On Error GoTo hiba
Shell "pbrush.exe ", vbNormalFocus
Exit Sub
hiba:
    MsgBox "A paint nem nyitható meg!", vbCritical, "Hiba:"
End Sub

Private Sub ponthat_Click()
For i = 1 To hatarok.Count
    hatarok(i).Enabled = ponthat.Value
Next i
End Sub





Private Sub segitseg_LostFocus()
elment
End Sub

Private Sub sugo_Click()
On Error GoTo hiba
If szerk.aktualis <> 0 Then
    Shell "hh.exe " & eleres & "\szerkeszto.chm::/page/elem.htm", vbNormalFocus
Else
    Shell "hh.exe " & eleres & "\szerkeszto.chm::/page/ptul.htm", vbNormalFocus
End If
Exit Sub
hiba:
    MsgBox "Az ön Windowsa nem képes kezelni a HTML Help fájlokat.", vbInformation, "Súgó nem tölthetõ be"


End Sub

Private Sub szama_Click()
If seg.Visible Then
    segitseg.Text = segito(szama.ListIndex + 1)
    megold.Text = megoldas(szama.ListIndex + 1)
End If
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
Public Sub segitobe(segitsegek As String)
On Error Resume Next

For i = 1 To 5
        segito(i) = ""
Next i

If Len(segitsegek) = 10 Then Exit Sub
Dim ker As Integer
    i = 1
    j = 1
       
        For ker = 1 To Len(segitsegek)
            If Mid(segitsegek, ker, 2) = "||" Then
                    segito(j) = Mid(segitsegek, i, ker - i)
                    i = ker + 2
                    j = j + 1
             End If
        Next
End Sub

Public Sub megoldasba(megoldasok As String)
On Error Resume Next

For i = 1 To 5
        megoldas(i) = ""
Next i

If Len(megoldasok) = 10 Then Exit Sub
Dim ker As Integer
    i = 1
    j = 1
       
        For ker = 1 To Len(megoldasok)
            If Mid(megoldasok, ker, 2) = "||" Then
                    megoldas(j) = Mid(megoldasok, i, ker - i)
                    i = ker + 2
                    j = j + 1
             End If
        Next
End Sub

Private Sub elment()
segito(szama.ListIndex + 1) = Trim(segitseg.Text)
megoldas(szama.ListIndex + 1) = Trim(megold.Text)

If (segito(szama.ListIndex + 1) = "" Or megoldas(szama.ListIndex + 1) = "") And szama.ListIndex <> 4 Then
    For i = szama.ListIndex + 2 To 5
            If (Trim(megoldas(i)) <> "" And Trim(segito(i)) <> "") Then j = 1
    Next i
    If j = 1 Then
        'MsgBox "Üresen hagyta a kérdés vagy a megoldás mezõt," & vbCrLf & _
               ' "úgy hogy ezt követõen még kérdések vannak." & vbCrLf & _
                '"Ez azt jelenti, hogy az üres kérdések vagy megoldások" & vbCrLf & _
                '"nem lesznek elmentve, így csak az utolsó helyesen" & vbCrLf & _
                '"megadott kérdés-válasz fog mûködni a vaktérképben.", vbCritical, "Üres kérdés vagy válasz mezõ"
    End If
End If
End Sub
Public Sub formatuma(id As Integer)
        jobbra.Value = Abs(CInt(szerk.cimke(id).Alignment))
        Select Case szerk.pont(id).Shape
            Case 3
                tip.Text = tip.List(0)
            Case 1
                tip.Text = tip.List(1)
        End Select
        
        pont.Shape = szerk.pont(id).Shape
        pont.BackColor = szerk.pont(id).BackColor
        pont.FillColor = szerk.pont(id).FillColor
        pont.BorderColor = szerk.pont(id).BorderColor
        pont.Visible = szerk.pont(id).Visible
        
        If szerk.cimke(id).Alignment = 0 Then
            cimke.Left = pont.Left + pont.Width + 15
            cimke.Alignment = 0
        Else
            cimke.Alignment = 1
            cimke.Left = pont.Left - (cimke.Width + 15)
        End If
        cimke.FontBold = szerk.cimke(id).FontBold
        cimke.FontItalic = szerk.cimke(id).FontItalic
        cimke.FontUnderline = szerk.cimke(id).FontUnderline
        cimke.FontSize = szerk.cimke(id).FontSize
        cimke.ForeColor = szerk.cimke(id).ForeColor
        
        lathatatlan.Value = Abs(Abs(CInt(pont.Visible)) - 1)
        felkover.Value = Abs(CInt(cimke.FontBold))
        dolt.Value = Abs(CInt(cimke.FontItalic))
        alahuzott.Value = Abs(CInt(cimke.FontUnderline))
        meret.Text = szerk.cimke(id).FontSize
End Sub
