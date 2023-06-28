VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form szerkeszto 
   BackColor       =   &H8000000C&
   Caption         =   "Vaktérkép Szerkesztõ"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6540
   Icon            =   "szerk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog pb 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton gomb 
      Caption         =   "T"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   4320
      Width           =   255
   End
   Begin VB.HScrollBar jb 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   4320
      Width           =   5055
   End
   Begin VB.VScrollBar fl 
      Height          =   3855
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox terulet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   720
      ScaleHeight     =   3825
      ScaleWidth      =   5025
      TabIndex        =   3
      Top             =   480
      Width           =   5055
      Begin VB.PictureBox ba 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox bf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   0
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox jf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   840
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox ja 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   840
         ScaleHeight     =   60
         ScaleWidth      =   60
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   90
      End
      Begin Vakablak.jel jel 
         Height          =   200
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   344
         KitoltesSzine   =   -2147483640
         KeretSzine      =   -2147483640
         HatterSzine     =   -2147483643
      End
      Begin VB.Shape keret 
         Height          =   375
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label jel_szoveg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jel szöveg"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Menu file 
      Caption         =   "&Projekt"
      Begin VB.Menu uj_mnu 
         Caption         =   "&Új projekt"
         Shortcut        =   ^U
      End
      Begin VB.Menu megnyit_mnu 
         Caption         =   "&Megnyitás"
         Shortcut        =   ^M
      End
      Begin VB.Menu v0 
         Caption         =   "-"
      End
      Begin VB.Menu ment_mnu 
         Caption         =   "Menté&s"
         Shortcut        =   ^S
      End
      Begin VB.Menu ment_mint_mnu 
         Caption         =   "Mentés m&ásként"
         Shortcut        =   ^A
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu olda_mnu 
         Caption         =   "Projekt &tulajdonságai"
         Shortcut        =   ^T
      End
      Begin VB.Menu beall_mnu 
         Caption         =   "&Beállítások"
         Shortcut        =   ^B
      End
      Begin VB.Menu megtekint_mnu 
         Caption         =   "Megt&ekintés..."
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu v2 
         Caption         =   "-"
      End
      Begin VB.Menu kilepes_mnu 
         Caption         =   "&Kilépés"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu szerkesztes_mnu 
      Caption         =   "Szerkesztés"
      Visible         =   0   'False
      Begin VB.Menu nev_mnu 
         Caption         =   "Névtelen"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu v6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu uj_elem_mnu 
         Caption         =   "Új elem"
         Enabled         =   0   'False
      End
      Begin VB.Menu v3 
         Caption         =   "-"
      End
      Begin VB.Menu torles_mnu 
         Caption         =   "Töröl"
      End
      Begin VB.Menu v4 
         Caption         =   "-"
      End
      Begin VB.Menu tipus_mnu 
         Caption         =   "Típusa"
         Begin VB.Menu tip_mnu 
            Caption         =   "Kikérdezendõ"
            Index           =   0
         End
         Begin VB.Menu tip_mnu 
            Caption         =   "Jelmagyarázat"
            Index           =   1
         End
         Begin VB.Menu tip_mnu 
            Caption         =   "Megjegyzés"
            Index           =   2
         End
      End
      Begin VB.Menu igazitas_mnu 
         Caption         =   "Szöveg igazítás"
         Begin VB.Menu szoveg_igazit 
            Caption         =   "Alá"
            Index           =   0
         End
         Begin VB.Menu szoveg_igazit 
            Caption         =   "Fölé"
            Index           =   1
         End
         Begin VB.Menu szoveg_igazit 
            Caption         =   "Középre"
            Index           =   2
         End
         Begin VB.Menu szoveg_igazit 
            Caption         =   "Jobbra"
            Index           =   3
         End
         Begin VB.Menu szoveg_igazit 
            Caption         =   "Balra"
            Index           =   4
         End
      End
      Begin VB.Menu meretez_mnu 
         Caption         =   "Jel méretezése"
      End
      Begin VB.Menu tulajdonsag_mnu 
         Caption         =   "Tulajdonságok"
      End
   End
   Begin VB.Menu sugo_mnu 
      Caption         =   "&Súgó"
      Begin VB.Menu sugo_mnup 
         Caption         =   "Segítség"
         Shortcut        =   {F1}
      End
      Begin VB.Menu v5 
         Caption         =   "-"
      End
      Begin VB.Menu koszonet_mnu 
         Caption         =   "Köszönet"
      End
      Begin VB.Menu nevjegy_mnu 
         Caption         =   "&Névjegy"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "szerkeszto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Globális konstansként nem definiált Egyéni tulajdonság - jelm: &H80000018
Option Explicit
Public Cime As String, Kephelye As String, obj As Object, tabulalo As Integer, mentett As Boolean
Public x1 As Double, y1 As Double, szel As Double, mag As Double, nagyitas As Double
Dim elemek(1 To 1024) As elem
Dim ures As elem, ux As Single, uy As Single
Dim mentettFajl As String, meretez As Boolean

Private Sub beall_mnu_Click()
    beallitasok.Show vbModal
End Sub

Private Sub fl_Change()
    terulet.Top = fl.Value
End Sub

Private Sub Form_DblClick()
    megnyit_mnu_Click
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    jel(tabulalo).Visible = True
    jel_szoveg(tabulalo).Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'MsgBox KeyCode
    terulet.SetFocus
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 39, 40
                    If jel.Count = 1 Then
                        tabulalo = 0
                        Exit Sub
                    End If
                    tabulalo = tabulalo + 1
                    If tabulalo > jel.Count - 1 Then tabulalo = 1
                    megjelol (tabulalo)
        
                Case 37, 38
                    If jel.Count = 1 Then
                        tabulalo = 0
                        Exit Sub
                    End If
                    tabulalo = tabulalo - 1
                    If tabulalo < 1 Then tabulalo = jel.Count - 1
                    megjelol (tabulalo)
        
                Case 13
                    If meretez Then
                            meretez = False
                            megjelol (tabulalo)
                        Else
                            jel_DblClick (tabulalo)
                    End If
                Case 46
                    torles_mnu_Click
                Case 27
                    If meretez Then
                            meretez = False
                            megjelol (tabulalo)
                        Else
                            tabulalo = 0
                            megjelol (tabulalo)
                    End If
            End Select
        Case 1
            
    End Select
            
End Sub


Private Sub Form_Load()
    'magyar_nyelv
    'megnyitas ("$vt\vakablak.ini")
    'Load tulajdonsagok
    torol
    'tulajdonsagok.Hide
End Sub

Public Sub Form_Resize()
Pozicional
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not menti Then
        Cancel = 1
        Exit Sub
    End If
    Megsemmisit
    End
End Sub

Private Sub foszerk_mnu_Click()
    If tabulalo = 0 Then
            terulet_MouseUp 2, 0, 0, 0
        Else
            jel_MouseDown tabulalo, 2, 0, 0, 0
    End If
End Sub

Private Sub gomb_Click()
    tulajdonsag_mnu_Click
End Sub

Private Sub jb_Change()
    terulet.Left = jb.Value
End Sub


Private Sub jel_DblClick(Index As Integer)
    If tulajdonsagok.Masolas Then
            tulajdonsagok.Formatuma (Index)
            tulajdonsagok.Show vbModal
            Exit Sub
    End If
    If meretez Then
            meretez = False
            jel_MouseDown tabulalo, 1, 0, 0, 0
            Exit Sub
    End If
    tulajdonsag_mnu_Click
End Sub

Private Sub jel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    terulet_DragDrop Source, jel(Index).Left + X, jel(Index).Top + Y
End Sub

Private Sub jel_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If TypeOf Source Is PictureBox Then
        terulet_DragDrop Source, jel(Index).Left + X, jel(Index).Top + Y
    End If
End Sub

Private Sub jel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not tulajdonsagok.Masolas Then
        tabulalo = Index
        megjelol (tabulalo)
        nev_mnu.Caption = jel_szoveg(tabulalo).Caption
        If Button = 2 Then
                uj_elem_mnu.Enabled = False
                torles_mnu.Enabled = True
                igazitas_mnu.Enabled = True
                meretez_mnu.Enabled = True
                tipus_mnu.Enabled = True
                Dim i As Integer
                For i = 0 To 2
                    tip_mnu(i).Checked = False
                Next i
                tip_mnu(elemek(Index).Kovetkezo - 1).Checked = True
                PopupMenu szerkesztes_mnu ', 0, terulet.Left + jel(tabulalo).Left + X, terulet.Top + jel(tabulalo).Top + Y
            Else
                If Shift = 1 Then
                    ux = X
                    uy = Y
                    jel_szoveg(Index).Visible = False
                    jel(Index).Visible = False
                    jel(Index).Drag
                End If
        End If
    Else
        tulajdonsagok.Formatuma (Index)
        tulajdonsagok.Show vbModal
End If
End Sub

Private Sub jel_szoveg_Click(Index As Integer)
    jel_MouseDown Index, 1, 0, 0, 0
End Sub

Private Sub jel_szoveg_DblClick(Index As Integer)
    jel_DblClick (Index)
End Sub

Private Sub jel_szoveg_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    terulet_DragDrop Source, jel_szoveg(Index).Left + X, jel_szoveg(Index).Top + Y
End Sub


Private Sub jel_szoveg_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If TypeOf Source Is PictureBox Then
        terulet_DragDrop Source, jel_szoveg(Index).Left + X, jel_szoveg(Index).Top + Y
    End If
End Sub

Private Sub jel_szoveg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Shift = 1 And Not tulajdonsagok.Masolas Then
            tabulalo = Index
            ux = X
            uy = Y
            'jel(Index).Visible = False
            jel_szoveg(Index).Visible = False
            jel_szoveg(Index).Drag
        Else
            jel_MouseDown Index, Button, Shift, 0, 0
    End If
End Sub

Private Sub kilepes_mnu_Click()
    Unload Me
End Sub

Private Sub koszonet_mnu_Click()
    koszonet.Caption = Szulo.koszonet_mnu.Caption
    koszonet.Show vbModal
End Sub

Private Sub megnyit_mnu_Click()
On Error GoTo megse
    If Not menti Then Exit Sub
ujra:
        pb.CancelError = True
        pb.DialogTitle = KozosSzovegek(3)
        pb.Filter = KozosSzovegek(4) & "(*.vtk)|*.vtk"
        pb.FileName = "*.vtk"
        pb.ShowOpen
        torol
        megnyitas (pb.FileName)
megse:
End Sub

Private Sub megtekint_mnu_Click()
    If Not menti Then Exit Sub
    Shell Vakterkep.Konyvtar & App.EXEName & ".exe " & mentettFajl, vbNormalFocus
End Sub

Private Sub ment_mint_mnu_Click()
On Error GoTo megse
ujra:
        pb.CancelError = True
        pb.DialogTitle = KozosSzovegek(5)
        pb.Filter = KozosSzovegek(4) & "(*.vtk)|*.vtk"
        pb.FileName = Cime & ".vtk"
        pb.ShowSave
        
        If Not mentes(pb.FileName) Then
            'MsgBox "A fájl mentése sikertelen volt, kérem adjon meg egy másik nevet!", vbExclamation, "Mentési hiba"
            GoTo ujra
        End If
megse:
End Sub

Private Sub ment_mnu_Click()
On Error GoTo megse
If mentettFajl = "" Then
ujra:
        pb.CancelError = True
        pb.DialogTitle = KozosSzovegek(6)
        pb.Filter = KozosSzovegek(4) & "(*.vtk)|*.vtk"
        pb.FileName = Cime & ".vtk"
        pb.ShowSave
        mentettFajl = pb.FileName
End If
    'pb.FileName = mentettFajl
        If Not mentes(mentettFajl) Then
            'MsgBox "A fájl mentése sikertelen volt, kérem adjon meg egy másik nevet!", vbExclamation, "Mentési hiba"
            GoTo ujra
        End If
        mentett = True
        ment_mnu.Enabled = False
megse:
End Sub



Private Sub meretez_mnu_Click()
    meretez = True
    megjelol (tabulalo)
End Sub

Private Sub nevjegy_mnu_Click()
    nevjegy.Show vbModal
End Sub

Private Sub olda_mnu_Click()
    tulajdonsagok.Mutat (0)
End Sub


Private Sub sugo_mnup_Click()
    HHSugo ("kezdo.htm")
End Sub




Private Sub szoveg_igazit_Click(Index As Integer)
    With jel_szoveg(tabulalo)
        Select Case Index
            Case 0 'Alá
                .Left = jel(tabulalo).Left + ((jel(tabulalo).Width - .Width) / 2)
                .Top = jel(tabulalo).Top + jel(tabulalo).Height
            Case 1 'Fölé
                .Left = jel(tabulalo).Left + ((jel(tabulalo).Width - .Width) / 2)
                .Top = jel(tabulalo).Top - .Height
            Case 2 'Középre
                .Left = jel(tabulalo).Left + ((jel(tabulalo).Width - .Width) / 2)
                .Top = jel(tabulalo).Top + ((jel(tabulalo).Height - .Height) / 2)
            Case 3 'jobbra
                .Left = jel(tabulalo).Left + jel(tabulalo).Width
                .Top = jel(tabulalo).Top + ((jel(tabulalo).Height - .Height) / 2)
            Case 4 'balra
                .Left = jel(tabulalo).Left - .Width
                .Top = jel(tabulalo).Top + ((jel(tabulalo).Height - .Height) / 2)
        End Select
        
        Cimkexy tabulalo, .Left - jel(tabulalo).Left, .Top - jel(tabulalo).Top
        

    End With
End Sub

Private Sub terulet_DblClick()
    If Not tulajdonsagok.Masolas Then
            tulajdonsagok.Mutat (0)
        Else
            tulajdonsagok.Masolas = False
            tulajdonsagok.Show vbModal
    End If
End Sub

Private Sub terulet_DragDrop(Source As Control, X As Single, Y As Single)
'MsgBox Source.Name
    If TypeOf Source Is jel Then
        jel(tabulalo).Left = X - ux
        jel(tabulalo).Top = Y - uy
    
        Cimkexy tabulalo, CSng(elemek(tabulalo).Bal), CSng(elemek(tabulalo).Felso)
        If meretez Then jel_MouseDown tabulalo, 1, 0, 0, 0
    End If
    If TypeOf Source Is Label Then
        jel_szoveg(tabulalo).Left = X - ux
        jel_szoveg(tabulalo).Top = Y - uy
        
        'jel(tabulalo).Left = jel_szoveg(tabulalo).Left - elemek(tabulalo).Bal
        'jel(tabulalo).Top = jel_szoveg(tabulalo).Top - elemek(tabulalo).Felso
        'elemek(tabulalo).Bal = jel_szoveg(tabulalo).Left - jel(tabulalo).Left
        'elemek(tabulalo).Felso = jel_szoveg(tabulalo).Top - jel(tabulalo).Top
        Cimkexy tabulalo, jel_szoveg(tabulalo).Left - jel(tabulalo).Left, jel_szoveg(tabulalo).Top - jel(tabulalo).Top
    End If
    If TypeOf Source Is PictureBox Then
        On Error Resume Next
        Source.Left = X - ux
        Source.Top = Y - uy
        With jel(tabulalo)
            Select Case Source.Name
                Case "bf"
                    .Move bf.Left, bf.Top, jf.Left - bf.Left + jf.Width, ba.Top - bf.Top + ba.Height
                    'Passzint (tabulalo)
                Case "ba"
                    .Move ba.Left, .Top, .Left + .Width - ba.Left, ba.Top + ba.Height - .Top
                    'Passzint (tabulalo)
                Case "jf"
                    .Move .Left, jf.Top, jf.Left + jf.Width - .Left, ja.Top + ja.Height - jf.Top
                    'Passzint (tabulalo)
                Case "ja"
                    .Move .Left, jf.Top, ja.Left + ja.Width - .Left, ja.Top + ja.Height - .Top
                    'Passzint (tabulalo)
            End Select
        End With
        Passzint (tabulalo)
    End If
jel(tabulalo).Visible = CBool(Mid(elemek(tabulalo).tipp, 1, 1))
jel_szoveg(tabulalo).Visible = CBool(Mid(elemek(tabulalo).tipp, 2, 1))
mentett = False
MentesAktiv
End Sub

Private Sub terulet_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'MsgBox Source.Name
    jel_szoveg(tabulalo).Visible = False
    If TypeOf Source Is jel Then
        jel(tabulalo).Visible = False
    End If
    If TypeOf Source Is PictureBox Then
        terulet_DragDrop Source, X, Y
    End If
End Sub

Private Sub terulet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not tulajdonsagok.Masolas Then
    tabulalo = 0
    megjelol (tabulalo)
    nev_mnu.Caption = Cime
    If Button = 2 Then
        ux = X
        uy = Y
        uj_elem_mnu.Enabled = Len(Kephelye) > 0
        torles_mnu.Enabled = False
        tipus_mnu.Enabled = False
        igazitas_mnu.Enabled = False
        meretez_mnu.Enabled = False
        PopupMenu szerkesztes_mnu
    End If
    Else
        tulajdonsagok.Masolas = False
        tulajdonsagok.Show vbModal
End If
End Sub

Private Sub tip_mnu_Click(Index As Integer)
    'Dim i As Integer
    'For i = 0 To 2
    '    tip_mnu(i).Checked = False
    'Next i
    elemek(tabulalo).Kovetkezo = Index + 1
    mentett = False
    MentesAktiv
End Sub

Private Sub torles_mnu_Click()
    If tabulalo = 0 Then Exit Sub
    If MsgBox("Biztosan törölni akarja a(z) " & jel_szoveg(tabulalo).Caption & " nevû elemet?", vbQuestion + vbYesNo, "Törlés megerõsítése") = vbNo Then Exit Sub

        With jel_szoveg(tabulalo)
            .Left = jel_szoveg(jel_szoveg.Count - 1).Left
            .Top = jel_szoveg(jel_szoveg.Count - 1).Top
            .Caption = jel_szoveg(jel_szoveg.Count - 1).Caption
            .ToolTipText = jel_szoveg(jel_szoveg.Count - 1).ToolTipText
            .BackStyle = jel_szoveg(jel_szoveg.Count - 1).BackStyle
            .BackColor = jel_szoveg(jel_szoveg.Count - 1).BackColor
            
            .Font = jel_szoveg(jel_szoveg.Count - 1).Font
            .FontBold = jel_szoveg(jel_szoveg.Count - 1).FontBold
            .FontItalic = jel_szoveg(jel_szoveg.Count - 1).FontItalic
            .FontSize = jel_szoveg(jel_szoveg.Count - 1).FontSize
            .FontStrikethru = jel_szoveg(jel_szoveg.Count - 1).FontStrikethru
            .FontUnderline = jel_szoveg(jel_szoveg.Count - 1).FontUnderline
            .ForeColor = jel_szoveg(jel_szoveg.Count - 1).ForeColor
        End With
        
        With jel(tabulalo)
            .Left = jel(jel.Count - 1).Left
            .Top = jel(jel.Count - 1).Top
            .ToolTipText = jel(jel.Count - 1).ToolTipText
            .HatterSzine = jel(jel.Count - 1).HatterSzine
            .Height = jel(jel.Count - 1).Height
            .jel = jel(jel.Count - 1).jel
            If jel(jel.Count - 1).jel = 6 Then
                .KepElerese = jel(jel.Count - 1).KepElerese
            End If
            .KeretSzine = jel(jel.Count - 1).KeretSzine
            .KeretTipus = jel(jel.Count - 1).KeretTipus
            .KeretVastagsaga = jel(jel.Count - 1).KeretVastagsaga
            .KitoltesTipus = jel(jel.Count - 1).KitoltesTipus
            .KitoltesSzine = jel(jel.Count - 1).KitoltesSzine
            .Width = jel(jel.Count - 1).Width
            .Visible = jel(jel.Count - 1).Visible
        End With
        elemek(tabulalo) = elemek(jel.Count - 1)
    
    elemek(jel.Count - 1) = ures
    Unload jel(jel.Count - 1)
    Unload jel_szoveg(jel_szoveg.Count - 1)
    mentett = False
    Form_KeyDown 37, 0
    
End Sub

Private Sub tulajdonsag_mnu_Click()
Dim i As Integer
    With tulajdonsagok
    If tabulalo <> 0 Then
        tulajdonsagok.tipusa (elemek(tabulalo).Kovetkezo)
        
        For i = 1 To 10
            .Kave i, elemek(tabulalo).kerdesek(i).Kerdes, elemek(tabulalo).kerdesek(i).Valasz
        Next i
        
    End If
    .Mutat (tabulalo)
    End With
End Sub

Private Sub uj_elem_mnu_Click()
    Load jel(jel.Count)
    With jel(jel.Count - 1)
        .Left = .BalKozep(ux)
        .Top = .FelsoKozep(uy)
        .Visible = True
    End With
    
    Load jel_szoveg(jel_szoveg.Count)
    With jel_szoveg(jel_szoveg.Count - 1)
        .Caption = "Elem" & jel.Count - 1
        Cimkexy jel_szoveg.Count - 1, jel(jel.Count - 1).Width, (jel(jel.Count - 1).Height - jel_szoveg(jel.Count - 1).Height) / 2
        .Visible = True
    End With
    tabulalo = jel.Count - 1
    elemek(tabulalo).Kovetkezo = 1
    elemek(tabulalo).tipp = "11"
    megjelol (tabulalo)
    mentett = False
    MentesAktiv
End Sub

Private Sub uj_mnu_Click()
    If Not menti Then Exit Sub
    torol
    terulet.Width = nevjegy.kep.Width
    terulet.Height = nevjegy.kep.Height
    terulet.Picture = Nothing
    'Unload tulajdonsagok
    tulajdonsagok.eredeti.Picture = Nothing
    tulajdonsagok.terulet.Picture = Nothing
    tulajdonsagok.terulet.Height = 2055
    tulajdonsagok.terulet.Width = 3135
    tulajdonsagok.terulet_DblClick
    
    Form_Resize
    'tulajdonsagok.megse.Enabled = False
    tulajdonsagok.Mutat (0)
End Sub
Public Sub torol()
    Dim i As Integer
    For i = 1 To jel.Count - 1
        Unload jel(i)
        Unload jel_szoveg(i)
        elemek(i) = ures
    Next i
    tabulalo = 0
    Me.Caption = Vakterkep.Verzio & " " & KozosSzovegek(7)
    ment_mnu.Enabled = False
    ment_mint_mnu.Enabled = False
    megtekint_mnu.Enabled = False
    Cime = KozosSzovegek(8)
    Kephelye = ""
    nagyitas = 1
    x1 = 0
    y1 = 0
    szel = 0
    mag = 0
    mentett = False
    mentettFajl = ""
End Sub
Public Sub Cimkexy(Index As Integer, Bal As Single, Felso As Single)
    elemek(Index).Bal = Bal
    elemek(Index).Felso = Felso
    
    jel_szoveg(Index).Left = jel(Index).Left + Bal
    jel_szoveg(Index).Top = jel(Index).Top + Felso
End Sub
Public Sub Kave(Index As Integer, Hanyadik As Integer, Kerdes As String, Valasz As String)
    elemek(Index).kerdesek(Hanyadik).Kerdes = Kerdes
    elemek(Index).kerdesek(Hanyadik).Valasz = Valasz
End Sub
Public Function mentes(Fajlnev As String) As Boolean
Dim i As Integer, j As Integer, f As String, seged As Variant
Dim Mappa As String, Fajl As String, emappa As String

Mappa = Konyvtara(Fajlnev)
emappa = Mid(CsakANeve(Fajlnev), 1, Len(CsakANeve(Fajlnev)) - 4) & "\"
Fajl = CsakANeve(Fajlnev)
On Error GoTo Hiba
Open Fajlnev For Output As 2
    Print #2, "; " & Vakterkep.Verzio & "." & App.Revision
    Print #2, ";Muráti Ákos 2001-2004 - Minden jog fenntartva."
    Print #2, ""
    Print #2, "cim=" & Cime
    Print #2, "kep=";
        seged = ""
        If RelativEleres(Konyvtara(Fajlnev), Kephelye) <> "" Then seged = "\" & RelativEleres(Konyvtara(Fajlnev), Kephelye)
        If RelativEleres(Vakterkep.Konyvtar, Kephelye) <> "" Then seged = "$vt\" & RelativEleres(Vakterkep.Konyvtar, Kephelye)
        If seged <> "" Then
                Print #2, seged
            Else
                On Error Resume Next
                If MsgBox(KozosSzovegek(9), vbYesNo + vbQuestion, "Alapkép mentése") = vbYes Then
                        MkDir Mappa & emappa
                        FileCopy Kephelye, Mappa & emappa & CsakANeve(Kephelye)
                        Print #2, "\" & emappa & CsakANeve(Kephelye)
                    Else
                        Print #2, Kephelye
                End If
        End If
    If Not (x1 = 0 And y1 = 0 And szel = terulet.Width And mag = terulet.Height) Then
        Print #2, "kijelol=" & x1 & ";" & y1 & ";" & szel & ";" & mag
    End If
    If Not (nagyitas = 1) Then
        Print #2, "nagyitas=" & nagyitas
    End If
    
    For i = 1 To jel_szoveg.Count - 1
            Select Case elemek(i).Kovetkezo
                Case 3
                    Print #2, "<megjegyzes";
                Case 2
                    Print #2, "<jelmagyarazat";
                Case 1
                    Print #2, "<elem";
            End Select
        With jel(i)
            Print #2, "=" & jel_szoveg(i).Caption
            Kiir "xy=" & .Left & "," & .Top
            If Not .Visible Then Kiir "lathatatlan-jel"
            If .ToolTipText <> "" Then Kiir "tipp=" & .ToolTipText
            Kiir "meret=" & .Width & "," & .Height
            If .jel <> 0 Then Kiir "jel=" & .jel
            If .jel = 6 Then
                    If MsgBox(Atalakit(KozosSzovegek(10), jel_szoveg(i).Caption), vbQuestion + vbYesNo, KozosSzovegek(11)) = vbNo Then
                            If RelativEleres(Konyvtara(Fajlnev), .KepElerese) <> "" Then
                                    Kiir "ikon=" & "\" & RelativEleres(Konyvtara(Fajlnev), .KepElerese)
                                Else
                                    Kiir "ikon=" & .KepElerese
                            End If
                        Else
                            On Error Resume Next
                            MkDir Mappa & emappa
                            FileCopy .KepElerese, Mappa & emappa & CsakANeve(.KepElerese)
                            Kiir "ikon=" & "\" & emappa & CsakANeve(.KepElerese)
                    End If
                Else
                    If .HatterSzine <> -2147483643 Then Kiir "hatter=" & .HatterSzine 'jel(0).HatterSzine
                    If .KeretTipus <> 1 Then Kiir "keret-tipus=" & .KeretTipus 'jel(0).KeretTipus
                    If .KeretVastagsaga <> 1 Then Kiir "keret-vastagsag=" & .KeretVastagsaga 'jel(0).KeretVastagsaga
                    If .KeretSzine <> -2147483640 Then Kiir "keret-szin=" & .KeretSzine  'jel(0).KeretSzine
                    If .Atlatszo Then
                            Kiir "atlatszo"
                        Else
                            If .KitoltesTipus <> 1 Then Kiir "kitoltes-tipus=" & .KitoltesTipus 'jel(0).KitoltesTipus
                            If .KitoltesSzine <> -2147483640 Then Kiir "kitoltes-szin=" & .KitoltesSzine  'jel(0).KitoltesSzine
                    End If
                    If .VonalAllas <> 1 Then Kiir "vonal=" & .VonalAllas
            End If
        End With
        
        With jel_szoveg(i)
            Kiir "szovegXY=" & elemek(i).Bal & "," & elemek(i).Felso
            If Not .Visible Then Kiir "lathatatlan-szoveg"
            If .FontName <> szerkeszto.FontName Then Kiir "betu-tipus=" & .FontName  'jel_szoveg(0).FontName
            If .FontSize <> szerkeszto.FontSize Then Kiir "betu-meret=" & .FontSize
            If .ForeColor <> -2147483640 Then Kiir "betu-szin=" & .ForeColor
            If .BackStyle = 1 And .BackColor <> -2147483643 Then Kiir "betu-hatter=" & .BackColor   'jel_szoveg(0).BackColor
                f = ""
                If .FontBold Then f = f & "f"
                If .FontItalic Then f = f & "d"
                If .FontUnderline Then f = f & "a"
                If .FontStrikethru Then f = f & "k"
                If f <> "" Then Kiir "formazas=" & f
        End With
        
        For j = 1 To 10
            If elemek(i).kerdesek(j).Kerdes <> "" Or elemek(i).kerdesek(j).Valasz <> "" Then
                Kiir "kerdes=" & elemek(i).kerdesek(j).Kerdes & "|" & elemek(i).kerdesek(j).Valasz
            End If
        Next j
        Print #2, "!>"
    Next i
Close 2
    
    mentes = True
    megtekint_mnu.Enabled = True
    Exit Function
    
Hiba:
    KozosHibak Err.Number
    mentes = False
    Close 2
End Function
Private Sub Kiir(Mit As String)
    Print #2, Chr(9) & Mit
End Sub
Public Function megnyitas(Fajlnev As String) As Boolean
Dim sor As String, i As Integer, j As Integer, ker As Integer, megvan As Boolean
Dim id As Integer, KID As Integer
Dim kulcsszo As String, parameter As String, kep As String
id = 0
Fajlnev = Atalakit(Fajlnev)
On Error GoTo Hiba
    Open Fajlnev For Input As 1
        Do While Not EOF(1)
            Line Input #1, sor
            On Error GoTo Hiba
            kulcsszo = ""
            parameter = ""
            If Mid(sor, 1, 1) = ";" Or Mid(sor, 1, 1) = "#" Or Mid(sor, 1, 1) = "/" Or Mid(sor, 1, 1) = "[" Or sor = "" Then GoTo kihagy
            kulcsszo = LCase(Utasitas(sor))
            parameter = Ertek(sor)
            
            'Parancs formázgatása
            If kulcsszo = "" Then kulcsszo = sor
            kulcsszo = Korulmetel(kulcsszo)
            kulcsszo = Trim(kulcsszo)
            parameter = Trim(parameter)
            
            'MsgBox kulcsszo & " := " & parameter
        
            
            Select Case kulcsszo
' ######################## Projektekkel összefüggõ beállítások #################################
                Case "cim", "cime"
                    Cime = parameter
                    Me.Caption = parameter & " - " & Vakterkep.Verzio & " Szerkesztõ mód"
                Case "terkep", "kep"
                    On Error GoTo kephiba
                        parameter = Atalakit(parameter, Konyvtara(Fajlnev))
                        Kephelye = parameter
                        terulet.Picture = LoadPicture(parameter)
                        x1 = 0
                        y1 = 0
                        szel = terulet.Width
                        mag = terulet.Height
                        Form_Resize
                    
                Case "kijelol"
                    On Error GoTo kephiba
                    x1 = Kicsontoz(parameter, ";", 0)
                    y1 = Kicsontoz(parameter, ";", 1)
                    szel = Kicsontoz(parameter, ";", 2)
                    mag = Kicsontoz(parameter, ";", 3)
                    
                    terulet.Width = szel
                    terulet.Height = mag
                    
                    terulet.Cls
                    atmeretez (Kephelye)
                    Form_Resize
                    
                Case "nagyitas"
                    On Error GoTo kephiba
                        nagyitas = parameter
                        terulet.Width = terulet.Width * parameter
                        terulet.Height = terulet.Height * parameter
                    
                        terulet.Cls
                        atmeretez (Kephelye)
                        Form_Resize
                Case "!>"
                    lezaras
                ' ############# OBJEKTUMOK
                    
                Case "<elem", "<megjegyzes", "<jelmagyarazat"
                    lezaras
                    id = id + 1
                    Load jel(id)
                    Set obj = jel(id)
                    obj.Height = 135
                    obj.Width = 135
                    obj.Visible = True
                    
                    Load jel_szoveg(id)
                    jel_szoveg(id).Caption = parameter
                    elemek(id).Kovetkezo = 1
                    
                    Select Case kulcsszo
                        Case "<megjegyzes"
                            elemek(id).Kovetkezo = 3
                        Case "<jelmagyarazat"
                            elemek(id).Kovetkezo = 2
                    End Select
                    
                    jel_szoveg(id).Visible = True
                    'alapértelmezések:
                    elemek(id).Bal = obj.Width
                    elemek(id).Felso = 0
                    KID = 0
                    
                '######## KOZOS TULAJDONSAGOK
                Case "pozicio", "xy", "koordianatak"
                    On Error Resume Next
                    'If obj.Name <> "jelm" Then
                        obj.Left = CSng(Atalakit(Kicsontoz(parameter, ",", 0)))
                        obj.Top = CSng(Atalakit(Kicsontoz(parameter, ",", 1)))
                    'End If
                Case "meret", "meretek"
                    On Error Resume Next
                    'If aze Then
                        obj.Width = CLng(Atalakit(Kicsontoz(parameter, ",", 0)))
                        obj.Height = CLng(Atalakit(Kicsontoz(parameter, ",", 1)))
                    'End If
                Case "tipp"
                    obj.ToolTipText = parameter
                    'If aze Then
                        jel_szoveg(id).ToolTipText = parameter
                    'End If
                
                Case "betu-tipus"
                    On Error Resume Next
                    'If aze Then
                            jel_szoveg(id).FontName = parameter
                    'End If
                    
                Case "betu-meret"
                    On Error Resume Next
                    parameter = CDbl(parameter)
                    'If aze Then
                            jel_szoveg(id).FontSize = parameter
                    'End If
                    
                Case "betu-szin"
                    On Error Resume Next
                    parameter = CLng(parameter)
                    'If aze Then
                            jel_szoveg(id).ForeColor = parameter
                    'End If
                
                Case "betu-hatter"
                    On Error Resume Next
                    parameter = CLng(parameter)
                        jel_szoveg(id).BackStyle = 1
                        jel_szoveg(id).BackColor = parameter
                    
                Case "formazas"
                    parameter = LCase(parameter)
                    'If aze Then
                            jel_szoveg(id).FontBold = VanEBenne(parameter, "f")
                            jel_szoveg(id).FontItalic = VanEBenne(parameter, "d")
                            jel_szoveg(id).FontUnderline = VanEBenne(parameter, "a")
                            jel_szoveg(id).FontStrikethru = VanEBenne(parameter, "k")
                    'End If
                        
                ' ########## ELEM EGYEDITULAJDONSÁGAI
                Case "kerdes"
                    'If aze Then
                        KID = KID + 1
                        elemek(id).kerdesek(KID).Kerdes = Kicsontoz(parameter, "|", 0)
                        elemek(id).kerdesek(KID).Valasz = Kicsontoz(parameter, "|", 1)
                    'End If
                Case "jel", "alakzat"
                    'If aze Then
                        obj.jel = CByte(parameter) Mod 8
                    'End If
                
                Case "ikon", "jelkep", "szimbolum"
                    On Error Resume Next
                    'If aze Then
                        obj.jel = 6
                        obj.KepElerese = Atalakit(parameter, Konyvtara(Fajlnev))
                    'End If
                    
                Case "szovegxy", "cimkexy", "cimke"
                    On Error Resume Next
                    'If aze Then
                            elemek(id).Bal = Kicsontoz(parameter, ",", 0)
                            elemek(id).Felso = Kicsontoz(parameter, ",", 1)
                            igazit (id)
                    'End If
                Case "hatter"
                    On Error Resume Next
                    'If aze Then
                        obj.HatterSzine = parameter
                    'End If
                Case "atlatszo"
                    'If aze Then
                        obj.Atlatszo = True
                    'End If
                Case "kitoltes-szin"
                    On Error Resume Next
                    'If aze Then
                        obj.KitoltesSzine = parameter
                    'End If
                Case "kitoltes-tipus"
                    On Error Resume Next
                    'If aze Then
                        obj.KitoltesTipus = parameter
                    'End If
                Case "keret-szin"
                    On Error Resume Next
                    'If aze Then
                        obj.KeretSzine = parameter
                    'End If
                Case "keret-tipus"
                    On Error Resume Next
                    'If aze Then
                        obj.KeretTipus = parameter
                    'End If
                Case "keret-vastagsag"
                    On Error Resume Next
                    'If aze Then
                        obj.KeretVastagsaga = parameter
                    'End If
                Case "vonal"
                        obj.VonalAllas = parameter
                Case "lathatatlan-jel"
                    If jel_szoveg(id).Visible Then obj.Visible = False
                Case "lathatatlan-szoveg"
                    If obj.Visible Then jel_szoveg(id).Visible = False
                Case "nyelv"
                    If parameter <> "" Then
                        beallitasok.nyelv(1).Value = True
                        NyelvAlkalmazasa (parameter)
                        Close 1
                        Exit Function
                    End If
            End Select
kihagy:
        Loop
    
    If aze Then lezaras
    Close 1
    Form_KeyDown 27, 0
    mentett = True
    mentettFajl = Fajlnev
    ment_mint_mnu.Enabled = True
    megtekint_mnu.Enabled = True
    
Exit Function
Hiba:
    'Dim uzenet As String
    Select Case Err.Number
        Case 52
            UzenetAblak Atalakit(KozosSzovegek(12), Fajlnev)
            'Close 1
            mentett = True
            uj_mnu_Click
        Case 53
            If Fajlnev <> Vakterkep.Konyvtar & "vakablak.ini" Then
                UzenetAblak KozosSzovegek(13)
                'Close 1
                mentett = True
                uj_mnu_Click
            End If
        Case 5
            'Hibás területkijelölés
        Case Else
            If MsgBox(KozosSzovegek(14), vbQuestion + vbYesNo, KozosSzovegek(15)) = vbYes Then
                UzenetAblak Err.Number & ": " & Err.Description
                On Error Resume Next
                GoTo kihagy
            Else
                'Close 1
                'Resume
                mentett = True
                uj_mnu_Click
            End If
        End Select
        'If uzenet <> "" Then MsgBox uzenet, vbExclamation, "Megnyitási hiba"
    Close 1
    Exit Function

kephiba:
    If MsgBox(Atalakit(KozosSzovegek(16), vbCrLf), vbCritical + vbYesNo, KozosSzovegek(17)) = vbYes Then
            terulet.Picture = Nothing
            terulet.Cls
            'Kephelye = ""
            GoTo kihagy
        Else
            Close 1
            mentett = True
            uj_mnu_Click
    End If
    
End Function
Private Sub lezaras()
On Error Resume Next
    obj.ZOrder 1
    If jel_szoveg.Count > 1 Then
        igazit (jel_szoveg.Count - 1)
        jel_szoveg(jel_szoveg.Count - 1).ZOrder (0)
    End If
    latszik jel_szoveg.Count - 1, jel(jel.Count - 1).Visible, jel_szoveg(jel_szoveg.Count - 1).Visible
   obj = Nothing
End Sub
Private Sub igazit(Index As Integer)
    Cimkexy Index, CSng(elemek(Index).Bal), CSng(elemek(Index).Felso)
End Sub
Public Sub tipus(Index As Integer, tipusa As Byte)
    elemek(Index).Kovetkezo = tipusa
End Sub
Public Sub latszik(Index As Integer, JelLathato As Boolean, JelSzovegLathato As Boolean)
    elemek(Index).tipp = Abs(JelLathato) & Abs(JelSzovegLathato)
End Sub
Private Function menti() As Boolean
Dim valaszt
If Not mentett And Kephelye <> "" Then
    valaszt = MsgBox(Atalakit(KozosSzovegek(18), Cime), vbYesNoCancel + vbExclamation, KozosSzovegek(19))
    Select Case valaszt
        Case vbYes
            menti = True
            ment_mnu_Click
        Case vbCancel
            menti = False
        Case vbNo
            menti = True
    End Select
Else
    menti = True
End If
End Function
Private Sub atmeretez(kep As String)
On Error Resume Next
    terulet.PaintPicture LoadPicture(kep), 0, 0, terulet.Width, terulet.Height, x1, y1, szel, mag
End Sub
Public Sub MentesAktiv()
If szerkeszto.Kephelye <> "" Then
    szerkeszto.ment_mint_mnu.Enabled = True
    szerkeszto.ment_mnu.Enabled = True
End If
End Sub
Public Sub Passzint(id As Integer)
With jel(id)
    bf.Move .Left, .Top
    ba.Move .Left, .Top + .Height - ja.Height
    jf.Move .Left + .Width - jf.Width, .Top
    ja.Move .Left + .Width - jf.Width, .Top + .Height - ja.Height
    keret.Move .Left, .Top, .Width, .Height
End With
    If id = 0 Then
            meretez = False
        Else
            Cimkexy id, jel_szoveg(id).Left - jel(id).Left, jel_szoveg(id).Top - jel(id).Top
    End If
    
    ba.Visible = meretez
    bf.Visible = meretez
    ja.Visible = meretez
    jf.Visible = meretez
    keret.Visible = meretez
End Sub
Private Sub ja_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ux = X
        uy = Y
        ja.Drag
    End If
End Sub
Private Sub jf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ux = X
        uy = Y
        jf.Drag
    End If
End Sub
Private Sub bf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ux = X
        uy = Y
        bf.Drag
    End If
End Sub
Private Sub ba_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ux = X
        uy = Y
        ba.Drag
    End If
End Sub

