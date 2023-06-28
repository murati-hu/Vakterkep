VERSION 5.00
Begin VB.Form tulajdonsagok 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tulajdonságai"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame tul_lap 
      Height          =   4215
      Index           =   1
      Left            =   960
      TabIndex        =   28
      Top             =   240
      Width           =   4335
      Begin VB.ComboBox vallas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox minta 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9E9E9&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   0
         Left            =   0
         ScaleHeight     =   1665
         ScaleWidth      =   4305
         TabIndex        =   59
         Top             =   2520
         Width           =   4335
         Begin Vakablak.jel jel 
            Height          =   135
            Left            =   1800
            TabIndex        =   62
            Top             =   720
            Width           =   135
            _ExtentX        =   873
            _ExtentY        =   873
            KitoltesSzine   =   -2147483640
            KeretSzine      =   -2147483640
            HatterSzine     =   -2147483643
         End
      End
      Begin VB.CheckBox kitolte 
         Caption         =   "Kitöltés:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox elrejt 
         Caption         =   "Láthatatlan alakzat"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   2160
         Width           =   4095
      End
      Begin VB.TextBox vastagsag 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   40
         Text            =   "1"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox keret 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox kitoltes 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton valaszt 
         Height          =   300
         Index           =   2
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton valaszt 
         Height          =   300
         Index           =   1
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton valaszt 
         Height          =   300
         Index           =   0
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton talloz 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3480
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox Alakzat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Vonal tájolása:"
         Height          =   195
         Index           =   12
         Left            =   1680
         TabIndex        =   64
         Top             =   1360
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4200
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Vastagsága:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   41
         Top             =   1340
         Width           =   885
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Keret:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   420
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Jel:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame tul_lap 
      Height          =   4215
      Index           =   3
      Left            =   7440
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox szama 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox segitseg 
         Height          =   2415
         Left            =   120
         OLEDragMode     =   1  'Automatic
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox megold 
         Height          =   525
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   3975
      End
      Begin VB.Label cimke 
         Caption         =   "Kérdés:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Megoldás:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   735
      End
   End
   Begin VB.Frame tul_lap 
      Height          =   4215
      Index           =   0
      Left            =   7560
      TabIndex        =   13
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton tipus 
         Caption         =   "Megjegyzés, felirat"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   49
         Top             =   2760
         Width           =   2415
      End
      Begin VB.OptionButton tipus 
         Caption         =   "Jelmagyarázat (külön ablakban)"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   48
         Top             =   2520
         Width           =   3255
      End
      Begin VB.OptionButton tipus 
         Caption         =   "Kikérdezendõ"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   47
         Top             =   2280
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton formatum_masolo 
         Caption         =   "Formátum másoló"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox Nev 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Text            =   "Névtelen"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox tipp 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   240
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label cimke 
         Caption         =   "Tipp szövege:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label cimke 
         Caption         =   "Név:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame oldal 
      Height          =   4215
      Left            =   4680
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox nagyito 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1038
            SubFormatType   =   5
         EndProperty
         Height          =   315
         ItemData        =   "tulajdonsagok.frx":0000
         Left            =   840
         List            =   "tulajdonsagok.frx":0002
         TabIndex        =   54
         Text            =   "Combo1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.PictureBox terulet 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   600
         ScaleHeight     =   2025
         ScaleWidth      =   3105
         TabIndex        =   52
         Top             =   1800
         Width           =   3135
         Begin VB.Shape kijelolo 
            BorderStyle     =   3  'Dot
            Height          =   2415
            Left            =   240
            Top             =   480
            Width           =   3855
         End
      End
      Begin VB.TextBox Cime 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox kep 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton talloz 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3840
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox eredeti 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   240
         ScaleHeight     =   2655
         ScaleWidth      =   3855
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Nagyítás:"
         Height          =   195
         Index           =   10
         Left            =   105
         TabIndex        =   55
         Top             =   960
         Width           =   675
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Cím:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   315
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Kép:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   330
      End
   End
   Begin VB.Frame tul_lap 
      Height          =   4215
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   4335
      Begin VB.PictureBox minta 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9E9E9&
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   1
         Left            =   0
         ScaleHeight     =   1665
         ScaleWidth      =   4305
         TabIndex        =   60
         Top             =   2520
         Width           =   4335
         Begin VB.Label jel_szoveg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "jel_szöveg"
            Height          =   195
            Left            =   960
            TabIndex        =   61
            Top             =   600
            Width           =   750
         End
      End
      Begin VB.CheckBox hatter 
         Caption         =   "Van háttérszine"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CommandButton valaszt 
         Caption         =   "Háttér"
         Height          =   255
         Index           =   4
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1560
         Width           =   975
      End
      Begin VB.CheckBox elrejt_szov 
         Caption         =   "Szöveg nem látszik"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox meret 
         Height          =   315
         Left            =   960
         TabIndex        =   46
         Text            =   "Combo1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton valaszt 
         Caption         =   "Betû Szín"
         Height          =   255
         Index           =   3
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox alahuzott 
         Caption         =   "Aláhúzott"
         Height          =   195
         Left            =   2160
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox athuzva 
         Caption         =   "Áthúzva"
         Height          =   195
         Left            =   2160
         TabIndex        =   26
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox betutipus 
         Height          =   315
         Left            =   960
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox dolt 
         Caption         =   "Dõlt"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox felkover 
         Caption         =   "Félkövér"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4080
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Betûtípus:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   720
      End
      Begin VB.Label cimke 
         Caption         =   "Méret:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton sugo 
      Caption         =   "Súgó"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton megse 
      Caption         =   "&Mégse"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton ful 
      BackColor       =   &H8000000D&
      Caption         =   "&Szöveg"
      Height          =   255
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton ful 
      BackColor       =   &H8000000D&
      Caption         =   "&Jel"
      Height          =   255
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton ful 
      Caption         =   "&Általános"
      Height          =   255
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton ful 
      BackColor       =   &H8000000D&
      Caption         =   "&Kérdések"
      Height          =   255
      Index           =   3
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "tulajdonsagok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const max_szelesseg = 3000
Const max_magassag = 2000

Dim id As Integer, px, py, lenntart As Boolean, kijM, kijSZ
Dim ures As KerdesValasz
Public Masolas As Boolean
Dim segito(1 To 10) As KerdesValasz


Private Sub alahuzott_Click()
    jel_szoveg.FontUnderline = alahuzott.Value
End Sub

Private Sub Alakzat_Click()
    talloz(1).Enabled = True
    valaszt(0).Enabled = True
    kitolte.Enabled = True
    kitoltes.Enabled = True
    valaszt(2).Enabled = True
    keret.Enabled = True
    valaszt(1).Enabled = True
    
    vastagsag.Enabled = True
    vallas.Enabled = True
    
    Select Case Alakzat.ListIndex
        Case 6, 7
            If Alakzat.ListIndex = 6 Then
                jel.KepElerese = Alakzat.List(6)
                
                valaszt(0).Enabled = False
                keret.Enabled = False
                valaszt(1).Enabled = False
                vastagsag.Enabled = False
                vallas.Enabled = False
            End If
            
            valaszt(0).Enabled = False
            kitolte.Enabled = False
            kitoltes.Enabled = False
            valaszt(2).Enabled = False
            
        Case Else
            'jel.jel = Alakzat.ListIndex
            talloz(1).Enabled = False
    End Select
    jel.jel = Alakzat.ListIndex
    Alakzat.ToolTipText = Alakzat.List(Alakzat.ListIndex)
End Sub

Private Sub athuzva_Click()
    jel_szoveg.FontStrikethru = athuzva.Value
End Sub

Private Sub betutipus_Click()
    jel_szoveg.FontName = betutipus.List(betutipus.ListIndex)
    Kozepre
End Sub

Private Sub Cime_Change()
    Me.Caption = Atalakit(KozosSzovegek(24), Cime.Text)
End Sub


Private Sub dolt_Click()
    jel_szoveg.FontItalic = dolt.Value
End Sub

Private Sub elrejt_szov_Click()
    jel_szoveg.Visible = Not CBool(elrejt_szov.Value)
    elrejt.Enabled = jel_szoveg.Visible
End Sub

Private Sub elrejt_Click()
    jel.Bekapcsolva = Not CBool(elrejt.Value)
    elrejt_szov.Enabled = jel.Bekapcsolva
End Sub

Private Sub felkover_Click()
    jel_szoveg.FontBold = felkover.Value
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case 27
            megse_Click
        Case 112
            sugo_Click
        'Case Else
        '    MsgBox KeyCode
    End Select
End Sub


Private Sub Form_Load()
    Dim i As Integer
    oldal.Move 120, 240
    
    For i = 0 To tul_lap.Count - 1
        tul_lap(i).Move 120, 240
        tul_lap(i).Visible = False
    Next i
    Me.Width = 4620
    'minta.Move 120, 2780
    
    For i = 0 To Screen.FontCount - 1
        betutipus.AddItem Screen.Fonts(i)
    Next i
    betutipus.ListIndex = 0
    
    
    vallas.AddItem "_"
    vallas.AddItem "|"
    vallas.AddItem "\"
    vallas.AddItem "/"
    vallas.ListIndex = 1
    
    
    meret.AddItem "8"
    meret.AddItem "10"
    meret.AddItem "11"
    meret.AddItem "12"
    meret.AddItem "14"
    meret.AddItem "16"
    meret.AddItem "20"
    meret.AddItem "22"
    meret.AddItem "24"
    meret.AddItem "32"
    meret.AddItem "72"
    
    nagyito.AddItem "10%"
    For i = 1 To 6
        nagyito.AddItem i * 25 & "%"
    Next i
    nagyito.AddItem "200%"
    nagyito.AddItem "400%"
    nagyito.Text = "100%"
    
    UjraNyelvel
    terulet_DblClick
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'MsgBox "most"
End Sub

Private Sub formatum_masolo_Click()
    Masolas = True
    Me.Hide
End Sub

Private Sub ful_Click(Index As Integer)
On Error Resume Next
    Dim i As Integer
    For i = 0 To ful.Count
        tul_lap(i).Visible = False
        ful(i).BackColor = &HC0C0C0
    Next i
    If id = 0 Then
            oldal.Visible = True
        Else
            oldal.Visible = False
            tul_lap(Index).Visible = True
    End If
    
    minta(0).Visible = False
    minta(1).Visible = False
    minta(Index - 1).Visible = True
    
    ful(Index).BackColor = vbButtonFace
    megse.SetFocus
End Sub



Private Sub hatter_Click()
        jel_szoveg.BackStyle = V(hatter.Value)
End Sub



Private Sub keret_Click()
    If keret.ListIndex <> 1 Then vastagsag.Text = 1
    jel.KeretTipus = keret.ListIndex
End Sub

Private Sub kitolte_Click()
    jel.Atlatszo = Not CBool(kitolte.Value)
End Sub

Private Sub kitoltes_Click()
     jel.KitoltesTipus = kitoltes.ListIndex
End Sub


Private Sub megold_Change()
    segito(szama.ListIndex + 1).Valasz = megold
End Sub

Private Sub megse_Click()
Dim i As Integer
    For i = 1 To 10
        segito(i) = ures
    Next i
    segitseg.Text = ""
    megold.Text = ""
    szama.Text = szama.List(0)
    If id = 0 Then
            'Unload tulajdonsagok
            Me.Hide
        Else
            Me.Hide
    End If
End Sub

Public Sub Mutat(Melyiket As Integer)
On Error Resume Next
    id = Melyiket
    If Melyiket = 0 Then
            With szerkeszto
                Cime.Text = .Cime
                Cime_Change
                kep.Text = .Kephelye
                nagyito.Text = .nagyitas * 100 & "%"
                If .Kephelye <> "" Then kijeloles .x1, .y1, .szel, .mag
            End With
            
            ful(1).Visible = False
            ful(2).Visible = False
            ful(3).Visible = False
            
        Else
            ful(1).Visible = True
            ful(2).Visible = True
            ful(3).Visible = True
            szama_Click
            Nev.Text = szerkeszto.jel_szoveg(Melyiket).Caption
            Nev_Change
            tipp.Text = szerkeszto.jel_szoveg(Melyiket).ToolTipText
            Formatuma (Melyiket)
    End If
    ful_Click (0)
    felkover_Click
    Me.Show vbModal
End Sub


Private Sub meret_Change()
    meret_Click
End Sub

Private Sub meret_Click()
On Error Resume Next
    jel_szoveg.FontSize = meret.Text
    Kozepre
End Sub


Private Sub nagyito_Change()
    If Right(nagyito.Text, 1) <> "%" And IsNumeric(nagyito.Text) Then
        nagyito.Text = nagyito.Text & "%"
    End If
    If Format(nagyito.Text, "####.##") > 4 Then
        nagyito.Text = "400%"
    End If
End Sub

Private Sub Nev_Change()
    jel_szoveg.Caption = Nev.Text
    Me.Caption = Atalakit(KozosSzovegek(24), Nev.Text)
End Sub

Private Sub ok_Click()
Dim i As Integer, nk
If id = 0 Then
    On Error Resume Next
    If kep.Text = "" Then
        megse_Click
        Exit Sub
    End If
'megse.Enabled = True
    With szerkeszto
        .Cime = Cime.Text
        .Caption = .Cime & " - " & Vakterkep.Verzio & " " & KozosSzovegek(7)
        .Kephelye = kep.Text
        
        .nagyitas = Format(nagyito.Text, "####.##")
        .terulet.Height = kijM * .nagyitas
        .terulet.Width = kijSZ * .nagyitas
        
        .x1 = kijelolo.Left * eredeti.Width / terulet.Width
        .y1 = kijelolo.Top * eredeti.Height / terulet.Height
        .szel = kijelolo.Width * eredeti.Width / terulet.Width
        .mag = kijelolo.Height * eredeti.Height / terulet.Height
        
        .terulet.Cls
        .terulet.PaintPicture eredeti.Picture, 0, 0, .terulet.Width, .terulet.Height, .x1, .y1, .szel, .mag
        
        .Form_Resize
    End With
Else
    With szerkeszto
        For i = 0 To 2
            If tipus(i).Value Then .tipus id, i + 1
        Next i
        .latszik id, jel.Bekapcsolva, jel_szoveg.Visible
        With .jel_szoveg(id)
            .Caption = Nev.Text
            .ToolTipText = tipp.Text
            .BackStyle = jel_szoveg.BackStyle
            .BackColor = jel_szoveg.BackColor
            
            .Font = jel_szoveg.Font
            .FontBold = jel_szoveg.FontBold
            .FontItalic = jel_szoveg.FontItalic
            .FontSize = jel_szoveg.FontSize
            .FontStrikethru = jel_szoveg.FontStrikethru
            .FontUnderline = jel_szoveg.FontUnderline
            .ForeColor = jel_szoveg.ForeColor
            .Visible = jel_szoveg.Visible
        End With
        'szerkeszto.Cimkexy id, jel_szoveg.Left - jel.Left, jel_szoveg.Top - jel.Top
        
        For i = 1 To 10
            .Kave id, i, segito(i).Kerdes, segito(i).Valasz
        Next i
        
        With .jel(id)
            .Atlatszo = jel.Atlatszo
            .ToolTipText = tipp.Text
            .HatterSzine = jel.HatterSzine
            .Height = jel.Height
            .jel = jel.jel
            If jel.jel = 6 Then
                .KepElerese = jel.KepElerese
            End If
            .KeretSzine = jel.KeretSzine
            .KeretTipus = jel.KeretTipus
            .KeretVastagsaga = jel.KeretVastagsaga
            .KitoltesTipus = jel.KitoltesTipus
            .KitoltesSzine = jel.KitoltesSzine
            .Width = jel.Width
            .VonalAllas = jel.VonalAllas
            .Visible = jel.Bekapcsolva
        End With
    End With
End If
    szerkeszto.mentett = False
    szerkeszto.MentesAktiv
    megse_Click
End Sub
Public Sub Formatuma(Index As Integer)
Dim i As Integer
    Masolas = False
    With szerkeszto
    'Szöveg Formázó értékek
    felkover.Value = V(.jel_szoveg(Index).FontBold)
    felkover_Click
    
    alahuzott.Value = V(.jel_szoveg(Index).FontUnderline)
    alahuzott_Click
    
    dolt.Value = V(.jel_szoveg(Index).FontItalic)
    dolt_Click
    
    athuzva.Value = V(.jel_szoveg(Index).FontStrikethru)
    athuzva_Click
    
    elrejt_szov.Value = V(Not .jel_szoveg(Index).Visible)
    elrejt_szov_Click
    
    hatter.Value = V(.jel_szoveg(Index).BackStyle)
    hatter_Click
    
    
    For i = 0 To betutipus.ListCount - 1
        If betutipus.List(i) = .jel_szoveg(Index).FontName Then
            betutipus.Text = betutipus.List(i)
            GoTo megva
        End If
    Next i
megva:

    meret.Text = .jel_szoveg(Index).FontSize
    meret_Click
    jel_szoveg.FontSize = meret.Text
    
    'betutipus.Text = .jel_szoveg(Index).FontName
    valaszt(3).BackColor = .jel_szoveg(Index).ForeColor
    valaszt(4).BackColor = .jel_szoveg(Index).BackColor
    
    'Jel értékei
        'alap
    jel.Width = .jel(Index).Width
    jel.Height = .jel(Index).Height
    elrejt.Value = V(Not .jel(Index).Visible)
    
    'Vonal
    jel.VonalAllas = .jel(Index).VonalAllas
    vallas.ListIndex = jel.VonalAllas
    
    'alakzat
    valaszt(0).BackColor = .jel(Index).HatterSzine
    Alakzat.RemoveItem 6
    If .jel(Index).jel = 6 Then
            Alakzat.AddItem .jel(Index).KepElerese, 6
        Else
            Alakzat.AddItem KozosSzovegek(32), 6
    End If
    Alakzat.ListIndex = .jel(Index).jel
    'kitoltes
    kitoltes.ListIndex = .jel(Index).KitoltesTipus
    valaszt(2).BackColor = .jel(Index).KitoltesSzine
    kitolte.Value = V(Not .jel(Index).Atlatszo)
    
    'keret
    keret.ListIndex = .jel(Index).KeretTipus
    vastagsag.Text = .jel(Index).KeretVastagsaga
    valaszt(1).BackColor = .jel(Index).KeretSzine
    End With
    Kozepre
    Szinez
End Sub


Private Sub segitseg_change()
    segito(szama.ListIndex + 1).Kerdes = segitseg.Text
End Sub

Private Sub sugo_Click()
    If id = 0 Then
        HHSugo ("ptul.htm")
    Else
        HHSugo ("elem.htm")
    End If
End Sub

Private Sub szama_Click()
    segitseg.Text = segito(szama.ListIndex + 1).Kerdes
    megold.Text = segito(szama.ListIndex + 1).Valasz
End Sub


Private Sub talloz_Click(Index As Integer)
Dim w, h
On Error GoTo megse
    With szerkeszto
        .pb.CancelError = True
        .pb.DialogTitle = KozosSzovegek(25)
        .pb.Filter = KozosSzovegek(26) & "|*.bmp;*.gif;*.jpg;*.jpe;*.jpeg"
        .pb.FileName = kep.Text
        .pb.ShowOpen
        If Index = 0 Then
                kep.Text = .pb.FileName
                nagyito.Text = "100%"
                kijeloles 0, 0, 0, 0
                terulet_DblClick
            
            Else
                'Alakzat.RemoveItem (6)
                'Alakzat.AddItem (.pb.FileName), 6
                Alakzat.List(6) = .pb.FileName
                Alakzat.Text = Alakzat.List(6)
        End If
    End With
megse:
End Sub



Public Sub terulet_DblClick()
    kijelolo.Move 0, 0, terulet.Width, terulet.Height
    kijM = eredeti.Height
    kijSZ = eredeti.Width
End Sub


Private Sub terulet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lenntart = True
    px = X
    py = Y
End Sub

Private Sub terulet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lenntart Then Exit Sub
    If px > X Then
            kijelolo.Left = X
            kijelolo.Width = px - X
        Else
            kijelolo.Left = px
            kijelolo.Width = X - px
    End If
    
    If py > Y Then
            kijelolo.Top = Y
            kijelolo.Height = py - Y
        Else
            kijelolo.Top = py
            kijelolo.Height = Y - py
    End If
    
End Sub

Private Sub terulet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lenntart = False
    kijM = kijelolo.Height * eredeti.Height / terulet.Height
    kijSZ = kijelolo.Width * eredeti.Width / terulet.Width
End Sub



Private Sub valaszt_Click(Index As Integer)
On Error GoTo megse
    With szerkeszto
        .pb.Color = valaszt(Index).BackColor
        .pb.ShowColor
        valaszt(Index).BackColor = .pb.Color
    End With
megse:
    Szinez
End Sub



Private Sub vallas_Click()
    jel.VonalAllas = vallas.ListIndex
End Sub

Private Sub vastagsag_Change()
On Error Resume Next
    jel.KeretVastagsaga = vastagsag.Text
End Sub
Public Sub Kave(Hanyadik As Integer, Kerdes As String, Valasz As String)
    segito(Hanyadik).Kerdes = Kerdes
    segito(Hanyadik).Valasz = Valasz
End Sub
Private Sub Szinez()
    jel.HatterSzine = valaszt(0).BackColor
    jel.KeretSzine = valaszt(1).BackColor
    jel.KitoltesSzine = valaszt(2).BackColor
    jel_szoveg.ForeColor = valaszt(3).BackColor
    If V(hatter.Value) Then jel_szoveg.BackColor = valaszt(4).BackColor
End Sub
Private Sub kijeloles(X, Y, sz, m)
Dim w, h
On Error Resume Next
                
                eredeti.Picture = LoadPicture(kep.Text)
                eredeti.Refresh
                terulet.Cls
                
                w = eredeti.Width * max_magassag / eredeti.Height
                h = eredeti.Height * max_szelesseg / eredeti.Width
                
                If eredeti.Width > eredeti.Height Then
                        terulet.Width = max_szelesseg
                        terulet.Height = h
                        terulet.PaintPicture eredeti.Picture, 0, 0, max_szelesseg, h
                    Else
                        terulet.Width = w
                        terulet.Height = max_magassag
                        terulet.PaintPicture eredeti.Picture, 0, 0, w, max_magassag
                End If
                
                terulet.Move (oldal.Width - terulet.Width) / 2, (oldal.Height - 1250 - terulet.Height) / 2 + 1250

                
                If sz = 0 Then
                    sz = eredeti.Width
                    m = eredeti.Height
                End If
                
                terulet.Refresh
                'On Error GoTo megse
                'kijelolo.Move X / (eredeti.Width / terulet.Width), Y / (eredeti.Height / terulet.Height), sz / (eredeti.Width / terulet.Width), m / (eredeti.Height / terulet.Height)
                terulet_MouseDown 1, 0, X / (eredeti.Width / terulet.Width), Y / (eredeti.Height / terulet.Height)
                terulet_MouseMove 1, 0, (X + sz) / (eredeti.Width / terulet.Width), (Y + m) / (eredeti.Height / terulet.Height)
                terulet_MouseUp 1, 0, (X + sz) / (eredeti.Width / terulet.Width), (Y + m) / (eredeti.Height / terulet.Height)
                
End Sub
Public Sub tipusa(tip As Byte)
On Error Resume Next
    tipus(tip - 1).Value = True
End Sub
Private Sub Kozepre()
    jel.Move jel.Left, (minta(0).Height - jel.Height) / 2
    jel.Move (minta(0).Width - jel.Width) / 2
    jel_szoveg.Move jel_szoveg.Left, (minta(1).Height - jel_szoveg.Height) / 2
    jel_szoveg.Move (minta(1).Width - jel_szoveg.Width) / 2
End Sub
Public Sub UjraNyelvel()
Dim i As Integer
    'Alakzatok listája
    Alakzat.Clear
    Alakzat.AddItem KozosSzovegek(47)
    Alakzat.AddItem KozosSzovegek(27)
    Alakzat.AddItem KozosSzovegek(28)
    Alakzat.AddItem KozosSzovegek(29)
    Alakzat.AddItem KozosSzovegek(30)
    Alakzat.AddItem KozosSzovegek(31)
    Alakzat.AddItem KozosSzovegek(32)
    Alakzat.AddItem KozosSzovegek(33)
    Alakzat.ListIndex = 0
    
    'Kerettípusok
    keret.Clear
    keret.AddItem KozosSzovegek(34)
    keret.AddItem "_______________________"
    keret.AddItem "__ __ __ __ __ __ __ __ __"
    keret.AddItem ". . . . . . . . . . . . . . . . . . . . . . . ."
    keret.AddItem "__ . __ . __ . __ . __ . __ . __ ."
    keret.AddItem ".. __ .. __ .. __ .. __ .. __ .. __"
    keret.ListIndex = 0
    
    'Kitöltéstípusok
    kitoltes.Clear
    kitoltes.AddItem KozosSzovegek(35)
    kitoltes.AddItem KozosSzovegek(34)
    kitoltes.AddItem "============="
    kitoltes.AddItem "|  |  |  |  |  |  |  |  |  |  |"
    kitoltes.AddItem "\ \ \ \ \ \ \ \ \ \ \"
    kitoltes.AddItem "/ / / / / / / / / / /"
    kitoltes.AddItem "############"
    kitoltes.AddItem "XXXXXXXXXXXX"
    'kitoltes.AddItem "Átlátszó"
    kitoltes.ListIndex = 0
    
    'Kérdésszámok
    szama.Clear
    For i = 1 To 10
        szama.AddItem Atalakit(KozosSzovegek(23), CStr(i))
    Next i
    szama.ListIndex = 0
End Sub
