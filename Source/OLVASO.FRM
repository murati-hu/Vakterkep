VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form terkep 
   BackColor       =   &H8000000C&
   Caption         =   "Vakablak"
   ClientHeight    =   4485
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5850
   Icon            =   "olvaso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar fl 
      Height          =   4095
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar jb 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   5295
   End
   Begin VB.CommandButton gomb 
      Caption         =   "J"
      Height          =   255
      Left            =   5280
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog pb 
      Left            =   1800
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox terulet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   4785
      TabIndex        =   3
      Top             =   240
      Width           =   4815
      Begin VB.TextBox szoveg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "Bevitel"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin Vakablak.jel jel 
         Height          =   135
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   873
         _ExtentY        =   873
         KitoltesSzine   =   -2147483640
         KeretSzine      =   -2147483640
         HatterSzine     =   -2147483643
      End
      Begin Vakablak.jel megj 
         Height          =   135
         Index           =   0
         Left            =   2880
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   873
         _ExtentY        =   873
         KitoltesSzine   =   -2147483640
         KeretSzine      =   -2147483640
         HatterSzine     =   -2147483643
      End
      Begin VB.Label jel_szoveg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Jel sz�veg"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label megj_szoveg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Jel sz�veg"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Menu fajl_mnu 
      Caption         =   "&F�jl"
      Begin VB.Menu ujra_mnu 
         Caption         =   "�jra kezd"
         Enabled         =   0   'False
         Shortcut        =   ^U
      End
      Begin VB.Menu megnyit_mnu 
         Caption         =   "Megnyit�s"
         Shortcut        =   ^M
      End
      Begin VB.Menu nyomtat 
         Caption         =   "Nyomtat�s..."
         Shortcut        =   ^P
      End
      Begin VB.Menu v7 
         Caption         =   "-"
      End
      Begin VB.Menu szerkesztes_mnu 
         Caption         =   "Szerkeszt�s..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu kilep_mnu 
         Caption         =   "Kil�p�s"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu eszkozok_mnu 
      Caption         =   "&Eszk�z�k"
      Begin VB.Menu jelm_mnu 
         Caption         =   "Jelmagyar�zat"
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
      Begin VB.Menu behuzo_mnu 
         Caption         =   "Beh�z�"
         Enabled         =   0   'False
         Shortcut        =   ^H
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu ertekel_mnu 
         Caption         =   "�rt�kel�s"
         Shortcut        =   ^E
      End
      Begin VB.Menu beall_mnu 
         Caption         =   "Be�ll�t�sok..."
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu sugo_mnu 
      Caption         =   "&S�g�"
      Begin VB.Menu help_mnu 
         Caption         =   "Seg�ts�g"
         Shortcut        =   {F1}
      End
      Begin VB.Menu v2 
         Caption         =   "-"
      End
      Begin VB.Menu koszonet_mnu 
         Caption         =   "K�sz�net..."
      End
      Begin VB.Menu nevjegy_mnu 
         Caption         =   "N�vjegy"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "terkep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public obj As Object, Cime As String, tabulalo As Integer
Dim elemek(1 To 1024) As elem
Dim ures As elem, proba As Byte
Public Megnyitva As String
Private Sub beall_mnu_Click()
    szoveg.Visible = False
    beallitasok.Show vbModal
End Sub

Private Sub behuzo_mnu_Click()
   If behuzo.Visible Then
        behuzo.Hide
    Else
        behuzo.Show , Me
        Me.SetFocus
    End If
End Sub

Private Sub ertekel_mnu_Click()
Dim jo As Integer, i As Integer, osszeg As Double

szoveg.Visible = False

For i = 1 To (jel.Count - 1)
    If elemek(i).pont = 0 Then
        If LegegyszerubbAlak(elemek(i).Cime) = LegegyszerubbAlak(jel_szoveg(i).Caption) Then
                Ellenorzott i, True
            
            If proba = 1 Or beallitasok.kerdesek.Value = 0 Then
                elemek(i).pont = beallitasok.pont.Text
            Else
                elemek(i).pont = beallitasok.pont.Text * ((100 - ((elemek(i).Kovetkezo) * beallitasok.levonas.Text)) / 100)
            End If
            jo = jo + 1
        Else
            Ellenorzott i, False
        End If
    Else
        jo = jo + 1
    End If
Next i

osszeg = 0
For i = 1 To jel.Count - 1
     osszeg = osszeg + elemek(i).pont
Next i


ertekeles.helyes.Caption = jo
ertekeles.hibak.Caption = (jel.Count - 1) - jo
ertekeles.pontok.Caption = osszeg
ertekeles.maxpont.Caption = (jel.Count - 1) * beallitasok.pont.Text

ertekeles.szazalek.Caption = Format(CDbl(ertekeles.pontok.Caption) / CDbl(ertekeles.maxpont.Caption) * 100, "##,##")
If ertekeles.szazalek.Caption = "" Then ertekeles.szazalek.Caption = 0

'Ki�rt�kel�si szab�lyok
jo = 0
For i = 0 To beallitasok.ert_hatarok.ListCount - 1
    If jo = 0 And CByte(ertekeles.szazalek.Caption) <= CByte(beallitasok.KiErtekeles(i, 1)) Then
        ertekeles.jegy.Caption = beallitasok.KiErtekeles(i, 3)
        ertekeles.neve.Caption = beallitasok.KiErtekeles(i, 2)
        jo = 1
    End If
Next i

'Ha nem tal�ltunk semmit
If jo = 0 Then
    ertekeles.jegy.Caption = beallitasok.KiErtekeles(0, 3)
    ertekeles.neve.Caption = beallitasok.KiErtekeles(0, 2)
End If

'Men� �t�r�sa
If proba = 5 Then
        ertekel_mnu.Caption = KozosSzovegek(20)
        ertekeles.Caption = KozosSzovegek(20)
    Else
        ertekel_mnu.Caption = proba + 1 & ". " & ertekeles.ertekeles
        ertekeles.Caption = proba & ". " & ertekeles.ertekeles
        proba = proba + 1
End If
ertekeles.Show vbModal
End Sub



Private Sub fl_Change()
    terulet.Top = fl.Value
End Sub

Private Sub Form_DblClick()
    megnyit_mnu_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If jel.Count = 1 Then Exit Sub
If szoveg.Visible Then
    Select Case KeyCode
        Case 13
            jel_szoveg(tabulalo).Caption = szoveg.Text
            szoveg.Visible = False
            igazit (tabulalo)
            terulet.SetFocus
        Case 27
            szoveg.Visible = False
            terulet.SetFocus
    End Select
Else
    If Not beallitasok.Visible Then terulet.SetFocus
    Select Case KeyCode
     Case 123
        Dim i As Integer
        For i = 1 To jel.Count - 1
            jel_szoveg(i).Caption = elemek(i).Cime
        Next i
     Case 39, 40
        tabulalo = tabulalo + 1
        If tabulalo > jel.Count - 1 Then tabulalo = 1
        megjelol (tabulalo)
        
    Case 37, 38
        tabulalo = tabulalo - 1
        If tabulalo < 1 Then tabulalo = jel.Count - 1
        megjelol (tabulalo)
        
    Case 13
        If jel(tabulalo).Bekapcsolva Then
            Beiras (tabulalo)
        End If
    End Select
End If
End Sub


Private Sub Form_Load()
    'MsgBox "'" & Lefoloz("Hello                        Bell�!   ", " ") & "'"
    'magyar_nyelv
    torol
End Sub

Private Sub Form_Resize()
Pozicional
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Megsemmisit
    End
End Sub

Private Sub gomb_Click()
    jelm_mnu_Click
End Sub

Private Sub help_mnu_Click()
    HHSugo ("kezdo.htm")
End Sub

Private Sub jb_Change()
    terulet.Left = jb.Value
End Sub

Private Sub jel_Click(Index As Integer)
    If Not szoveg.Visible Then
            tabulalo = Index
            Form_KeyDown 13, 0
        Else
            Form_KeyDown 27, 0
    End If
End Sub

Private Sub jel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    jel_szoveg_DragDrop Index, Source, X, Y
End Sub

Private Sub jel_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    jel_szoveg_DragOver Index, Source, X, Y, State
End Sub

Private Sub jel_Hiba(Index As Integer, hibakod As Byte)
    MsgBox "Az egyik szimb�lum hib�s vagy nem tal�lhat�!", vbInformation, "Hib�s szimb�lum"
    obj.jel = 1
End Sub

Private Sub jel_szoveg_Click(Index As Integer)
    jel_Click (Index)
End Sub

Private Sub jel_szoveg_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If IrhatE(Index) Then
        jel_szoveg(Index).Caption = Source.Caption
        Me.SetFocus
    End If
End Sub

Private Sub jel_szoveg_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    megjelol (Index)
End Sub

Private Sub jelm_mnu_Click()
    If jelmagyarazat.Visible Then
        jelmagyarazat.Hide
    Else
        jelmagyarazat.Show , Me
        Me.SetFocus
    End If
End Sub

Private Sub kilep_mnu_Click()
    Unload Me
End Sub

Private Sub koszonet_mnu_Click()
    koszonet.Caption = Szulo.koszonet_mnu.Caption
    koszonet.Show vbModal
End Sub

Private Sub megnyit_mnu_Click()
On Error GoTo megse
    pb.CancelError = True
    pb.DialogTitle = KozosSzovegek(3)
    pb.Filter = KozosSzovegek(4) & "(*.vtk)|*.vtk"
    pb.FileName = "*.vtk"
    pb.ShowOpen
    torol
    megnyitas (pb.FileName)
megse:
End Sub

Public Sub megnyitas(Fajlnev As String)
Dim sor As String, i As Integer, j As Integer, ker As Integer, megvan As Boolean
Dim id As Integer, JID As Integer, KID As Integer, FID As Integer
Dim kulcsszo As String, parameter As String
Dim kep As String, X, Y, szel, mag
id = 0
JID = 0
KID = 0
FID = 0
Fajlnev = Atalakit(Fajlnev)
On Error GoTo Hiba
    Open Fajlnev For Input As 1
        terulet.Visible = False
        Do While Not EOF(1)
            Line Input #1, sor
            On Error GoTo Hiba
            kulcsszo = ""
            parameter = ""
            If Mid(sor, 1, 1) = ";" Or Mid(sor, 1, 1) = "#" Or Mid(sor, 1, 1) = "/" Or Mid(sor, 1, 1) = "[" Or sor = "" Then GoTo kihagy
            kulcsszo = LCase(Utasitas(sor))
            parameter = Ertek(sor)
            
            'Parancs form�zgat�sa
            If kulcsszo = "" Then kulcsszo = sor
            kulcsszo = Korulmetel(kulcsszo)
            parameter = Korulmetel(parameter)
            
            'MsgBox kulcsszo & " := " & parameter
        
            
            Select Case kulcsszo
' ######################## Projektekkel �sszef�gg� be�ll�t�sok #################################
                Case "cim", "cime"
                    Cime = parameter
                    Me.Caption = parameter & " - " & Vakterkep.Verzio
                Case "terkep", "kep"
                    On Error GoTo kephiba
                        parameter = Atalakit(parameter, Konyvtara(Fajlnev))
                        kep = parameter
                        
                        terulet.Picture = LoadPicture(parameter)
                        nyomtat.Enabled = True
                        X = 0
                        Y = 0
                        szel = terulet.Width
                        mag = terulet.Height
                        
                        Form_Resize
                Case "kijelol"
                    X = Kicsontoz(parameter, ";", 0)
                    Y = Kicsontoz(parameter, ";", 1)
                    szel = Kicsontoz(parameter, ";", 2)
                    mag = Kicsontoz(parameter, ";", 3)
                    
                    If Y <> 0 And X <> 0 And szel <> terulet.ScaleWidth Or mag <> terulet.ScaleHeight Then nyomtat.Enabled = False
                    terulet.Width = szel
                    terulet.Height = mag
                    
                    terulet.Cls
                    terulet.PaintPicture LoadPicture(kep), 0, 0, terulet.Width, terulet.Height, X, Y, szel, mag
                     'nyomtat.Enabled = False
                     'terulet.PaintPicture
                     'terulet.Picture = terulet
                    Form_Resize
                    
                Case "nagyitas"
                    On Error Resume Next
                        terulet.Width = terulet.Width * parameter
                        terulet.Height = terulet.Height * parameter
                    terulet.Cls
                    terulet.PaintPicture LoadPicture(kep), 0, 0, terulet.Width, terulet.Height, X, Y, szel, mag
                    Form_Resize
                Case "!>"
                    lezaras
                ' ############# OBJEKTUMOK
                    
                Case "<elem"
                    lezaras
                    id = id + 1
                    Load jel(id)
                    Set obj = jel(id)
                    obj.Height = 135
                    obj.Width = 135
                    obj.Visible = True
                    
                    Load jel_szoveg(id)
                    jel_szoveg(id).Caption = beallitasok.jel.Text
                    jel_szoveg(id).Visible = True
                    elemek(id).Cime = parameter
                    
                    
                
                    'alap�rtelmez�sek:
                    elemek(id).Kovetkezo = 1
                    elemek(id).Bal = obj.Width
                    elemek(id).Felso = 0
                    KID = 0
                    
                Case "<megjegyzes"
                    lezaras
                    FID = FID + 1
                    Load megj(FID)
                    Set obj = megj(FID)
                    obj.Height = 135
                    obj.Width = 135
                    obj.Visible = True
                    
                    Load megj_szoveg(FID)
                    megj_szoveg(FID).Caption = parameter
                    megj_szoveg(FID).Visible = True
                    
                Case "<jelmagyarazat"
                    lezaras
                    JID = JID + 1
                    Load jelmagyarazat.jelm(JID)
                    Set obj = jelmagyarazat.jelm(JID)
                    obj.Height = 135
                    obj.Width = 135
                    obj.Top = Nagyobb(jelmagyarazat.jelm(JID - 1).Top + jelmagyarazat.jelm(JID - 1).Height, jelmagyarazat.jelm_szoveg(JID - 1).Top + jelmagyarazat.jelm_szoveg(JID - 1).Height) + 100
                    obj.Visible = True
                    
                    Load jelmagyarazat.jelm_szoveg(JID)
                    jelmagyarazat.jelm_szoveg(JID).Caption = parameter
                    
                '######## KOZOS TULAJDONSAGOK
                Case "pozicio", "xy", "koordianatak"
                    On Error Resume Next
                    If obj.Name <> "jelm" Then
                        obj.Left = CSng(Atalakit(Kicsontoz(parameter, ",", 0)))
                        obj.Top = CSng(Atalakit(Kicsontoz(parameter, ",", 1)))
                    End If
                Case "meret", "meretek"
                    On Error Resume Next
                    If aze Then
                        obj.Width = CLng(Atalakit(Kicsontoz(parameter, ",", 0)))
                        obj.Height = CLng(Atalakit(Kicsontoz(parameter, ",", 1)))
                    End If
                Case "tipp"
                    obj.ToolTipText = parameter
                    If obj.Name = "jel" Then
                        jel_szoveg(id).ToolTipText = parameter
                    End If
                
                Case "betu-tipus"
                    On Error Resume Next
                    Select Case obj.Name
                        Case "jel"
                            jel_szoveg(id).FontName = parameter
                        Case "megj"
                            megj_szoveg(FID).FontName = parameter
                        Case "jelm"
                            jelmagyarazat.jelm_szoveg(JID).FontName = parameter
                    End Select
                    
                Case "betu-meret"
                    On Error Resume Next
                    parameter = CDbl(parameter)
                    Select Case obj.Name
                        Case "jel"
                            jel_szoveg(id).FontSize = parameter
                        Case "megj"
                            megj_szoveg(FID).FontSize = parameter
                        Case "jelm"
                            jelmagyarazat.jelm_szoveg(JID).FontSize = parameter
                    End Select
                    
                Case "betu-szin"
                    On Error Resume Next
                    parameter = CLng(parameter)
                    Select Case obj.Name
                        Case "jel"
                            jel_szoveg(id).ForeColor = parameter
                        Case "megj"
                            megj_szoveg(FID).ForeColor = parameter
                        Case "jelm"
                            jelmagyarazat.jelm_szoveg(JID).ForeColor = parameter
                    End Select
                    
                Case "betu-hatter"
                    On Error Resume Next
                    parameter = CLng(parameter)
                    Select Case obj.Name
                        Case "megj"
                            megj_szoveg(FID).BackStyle = 1
                            megj_szoveg(FID).BackColor = parameter
                        Case "jelm"
                            jelmagyarazat.jelm_szoveg(JID).BackStyle = 1
                            jelmagyarazat.jelm_szoveg(JID).BackColor = parameter
                    End Select
                Case "formazas"
                    parameter = LCase(parameter)
                    Select Case obj.Name
                        Case "jel"
                            jel_szoveg(id).FontBold = VanEBenne(parameter, "f")
                            jel_szoveg(id).FontItalic = VanEBenne(parameter, "d")
                            jel_szoveg(id).FontUnderline = VanEBenne(parameter, "a")
                            jel_szoveg(id).FontStrikethru = VanEBenne(parameter, "k")
                        Case "megj"
                            megj_szoveg(FID).FontBold = VanEBenne(parameter, "f")
                            megj_szoveg(FID).FontItalic = VanEBenne(parameter, "d")
                            megj_szoveg(FID).FontUnderline = VanEBenne(parameter, "a")
                            megj_szoveg(FID).FontStrikethru = VanEBenne(parameter, "k")
                        Case "jelm"
                            jelmagyarazat.jelm_szoveg(JID).FontBold = VanEBenne(parameter, "f")
                            jelmagyarazat.jelm_szoveg(JID).FontItalic = VanEBenne(parameter, "d")
                            jelmagyarazat.jelm_szoveg(JID).FontUnderline = VanEBenne(parameter, "a")
                            jelmagyarazat.jelm_szoveg(JID).FontStrikethru = VanEBenne(parameter, "k")
                    End Select
                        
                ' ########## ELEM EGYEDI TULAJDONS�GAI
                Case "kerdes"
                    If obj.Name = "jel" And KID < 10 Then
                        KID = KID + 1
                        elemek(id).kerdesek(KID).Kerdes = Kicsontoz(parameter, "|", 0)
                        elemek(id).kerdesek(KID).Valasz = Kicsontoz(parameter, "|", 1)
                    End If
                Case "jel", "alakzat"
                    If aze Then
                        obj.jel = CByte(parameter) Mod 8
                    End If
                
                Case "ikon", "jelkep", "szimbolum"
                    On Error Resume Next
                    If aze Then
                        obj.jel = 6
                        obj.KepElerese = Atalakit(parameter, Konyvtara(Fajlnev))
                    End If
                    
                Case "szovegxy", "cimkexy", "cimke"
                    On Error Resume Next
                    Select Case obj.Name
                        Case "jel"
                            elemek(id).Bal = Kicsontoz(parameter, ",", 0)
                            elemek(id).Felso = Kicsontoz(parameter, ",", 1)
                            igazit (id)
                        Case "megj"
                            megj_szoveg(FID).Left = megj(FID).Left + Kicsontoz(parameter, ",", 0)
                            megj_szoveg(FID).Top = megj(FID).Top + Kicsontoz(parameter, ",", 1)
                    End Select
                Case "hatter"
                    On Error Resume Next
                    If aze Then
                        obj.HatterSzine = parameter
                    End If
                Case "atlatszo"
                    If aze Then
                        obj.Atlatszo = True
                    End If
                Case "kitoltes-szin"
                    On Error Resume Next
                    If aze Then
                        obj.KitoltesSzine = parameter
                    End If
                Case "kitoltes-tipus"
                    On Error Resume Next
                    If aze Then
                        obj.KitoltesTipus = parameter
                    End If
                Case "keret-szin"
                    On Error Resume Next
                    If aze Then
                        obj.KeretSzine = parameter
                    End If
                Case "keret-tipus"
                    On Error Resume Next
                    If aze Then
                        obj.KeretTipus = parameter
                    End If
                Case "keret-vastagsag"
                    On Error Resume Next
                    If aze Then
                        obj.KeretVastagsaga = parameter
                    End If
                Case "vonal"
                        obj.VonalAllas = parameter
                Case "lathatatlan-jel"
                    'If obj.Name <> "jelm" Then  'And megj_szoveg(FID).Visible Then
                        obj.Visible = False
                    'End If
                Case "lathatatlan-szoveg"
                    If obj.Name = "megj" Then megj_szoveg(FID).Visible = False
' ######################## Be�ll�t�sokkal kapcsolatos �rt�kek ###################################
                'Case "hatarok", "ponthatarok"
                '    If beallitasok.egyeni.Value = 1 Or Fajlnev = Atalakit("$vt\vakablak.ini") Then
                '        For i = 1 To 4
                '            beallitasok.hatarok(i) = Kicsontoz(parameter, ",", i - 1)
                '        Next i
                '    End If
                Case "beallitas", "beallitasok_engedelyezese"
                    beallitasok.enged.Value = parameter
                    beall_mnu.Enabled = Abs(CInt(parameter))
            
                Case "egyeni", "egyeni_beallitasok"
                    beallitasok.egyeni.Value = parameter
                Case "jel", "helyettesito_szoveg"
                    beallitasok.jel = parameter
                Case "pont"
                    beallitasok.pont = parameter
                Case "segito", "kerdesek_engedelyezese"
                    beallitasok.kerdesek = parameter
                Case "tippek", "tippek_engedelyezese"
                    beallitasok.tippek = parameter
                Case "minusz", "levonasok"
                    beallitasok.levonas = parameter
                Case "behuzasok_engedelyezese"
                    beallitasok.behuzas = parameter
                Case "szerkesztes_engedelyezese"
                    beallitasok.szerkesztes = parameter
                Case "nyelv"
                    If parameter <> "0" Then
                        beallitasok.nyelv(1).Value = True
                        NyelvAlkalmazasa (parameter)
                    End If
            End Select
kihagy:
        Loop
        Close 1
    
    'M�g le nem z�rt objektumok lez�r�sa
    If aze Then lezaras
    terulet.Visible = True
    
    If Fajlnev <> Atalakit("$vt\vakablak.ini") Then
            'van kik�rdezend� elem
            If jel.Count > 1 Then
                ujra_mnu.Enabled = True
                ertekel_mnu.Enabled = True
        
                'Ha enged�lyezve van, akkor beh�z� aktiv�l�sa
                If beallitasok.behuzas.Value = 1 Then
                        behuzo_mnu.Enabled = True
                    Else
                        behuzo_mnu.Enabled = False
                End If
                
                'beh�z� elemeinek �sszekever�se
                For i = 1 To jel.Count - 1 Step 2
                    Ujbehuzo (i) 'p�ratlanok
                Next i
            
                For i = 2 To jel.Count - 1 Step 2
                    Ujbehuzo (i) 'p�rosak
                Next i
            
                For i = 1 To behuzo.szoveg.Count - 1
                    If behuzo.szoveg(i).Width + 200 > behuzo.Width - 120 Then
                        behuzo.Width = (behuzo.szoveg(i).Width + 300) + 120 '* (behuzo.Width / behuzo.ScaleWidth)
                    End If
                Next i
                behuzo.Height = (behuzo.szoveg(behuzo.szoveg.Count - 1).Top + behuzo.szoveg(behuzo.szoveg.Count - 1).Height + 610) '* 1.5 '(behuzo.Height / (behuzo.ScaleHeight))
            End If
    
    
            'jelmagyar�zat �s ablakm�ret�nek be�ll�t�sa
            If jelmagyarazat.jelm.Count > 1 Then
                jelm_mnu.Enabled = True
                jelmagyarazat.Height = (Nagyobb(jelmagyarazat.jelm(jelmagyarazat.jelm.Count - 1).Top + jelmagyarazat.jelm(jelmagyarazat.jelm.Count - 1).Height, jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Top + jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Height)) + 610 '* 1.5 '* (jelmagyarazat.Height / jelmagyarazat.ScaleHeight)
                
                For i = 1 To JID
                    If jelmagyarazat.Width < (200 + jelmagyarazat.jelm_szoveg(i).Left + jelmagyarazat.jelm_szoveg(i).Width) Then '* (jelmagyarazat.ScaleWidth / jelmagyarazat.Width) Then
                        jelmagyarazat.Width = (200 + jelmagyarazat.jelm_szoveg(i).Left + jelmagyarazat.jelm_szoveg(i).Width) + 120 '* (jelmagyarazat.ScaleWidth / jelmagyarazat.Width)
                    End If
                Next i
            End If
            
            Form_KeyDown 40, 0
            Megnyitva = Fajlnev
            szerkesztes_mnu.Enabled = CBool(beallitasok.szerkesztes.Value)
    End If
Exit Sub
Hiba:
    'Dim uzenet As String
    Select Case Err.Number
        Case 52
            UzenetAblak Atalakit(KozosSzovegek(12), Fajlnev)
            'Exit Sub
        Case 53
            If Fajlnev <> Vakterkep.Konyvtar & "vakablak.ini" Then
                UzenetAblak KozosSzovegek(21)
            End If
            'Exit Sub
        Case Else
            UzenetAblak KozosSzovegek(22)
            UzenetAblak Err.Number & ": " & Err.Description, KozosSzovegek(15)
            torol
    End Select
    'If uzenet <> "" Then MsgBox uzenet, vbExclamation, "Megnyit�si hiba"
    Close 1
Exit Sub

'k�phib�k kezel�se
kephiba:
    UzenetAblak Atalakit(KozosSzovegek(16), vbCrLf), KozosSzovegek(17)
    UzenetAblak Err.Description, Err.Number & " - " & kep
    Close 1
    torol
End Sub

Private Sub torol()
    Dim i As Integer, j As Integer
    ujra_mnu_Click
    For i = 1 To jel.Count - 1
        Unload jel(i)
        Unload jel_szoveg(i)
        Unload behuzo.szoveg(i)
        elemek(i) = ures
    Next i
    For i = 1 To megj.Count - 1
        Unload megj(i)
        Unload megj_szoveg(i)
    Next i
    For i = 1 To jelmagyarazat.jelm.Count - 1
        Unload jelmagyarazat.jelm(i)
        Unload jelmagyarazat.jelm_szoveg(i)
    Next i
    
    jelm_mnu.Enabled = False
    behuzo_mnu.Enabled = False
    jelmagyarazat.Visible = False
    behuzo.Visible = False
    ertekel_mnu.Enabled = False
    nyomtat.Enabled = False
    Me.Caption = Vakterkep.Verzio
    terulet.Picture = nevjegy.kep.Picture
    terulet.Visible = True
    
    jelmagyarazat.Height = 510
    jelmagyarazat.Width = 1560
    Cime = ""
    Megnyitva = ""
    tabulalo = 0
    
    Form_Resize
End Sub
Private Sub lezaras()
On Error Resume Next
    obj.ZOrder 1
    Select Case obj.Name
        Case "jel"
            If jel_szoveg.Count > 1 Then
                igazit (jel_szoveg.Count - 1)
                jel_szoveg(jel_szoveg.Count - 1).Visible = True
                jel_szoveg(jel_szoveg.Count - 1).ZOrder (0)
            End If
        Case "jelm"
            If jelmagyarazat.jelm_szoveg.Count > 1 Then
                'obj.Visible = True
                jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Visible = True
                'jelmagyarazat.Height = Nagyobb(jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Top + jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Height, jelmagyarazat.jelm(jelmagyarazat.jelm.Count - 1).Top + jelmagyarazat.jelm(jelmagyarazat.jelm.Count - 1).Height) + 200 * (jelmagyarazat.ScaleHeight / jelmagyarazat.Height) + ((jelmagyarazat.jelm.Count - 1) * 100)
                jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Left = obj.Left + obj.Width + 100
                jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Top = obj.Top
                jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).ZOrder (0)
            End If
        Case "megj"
            If megj.Count > 1 Then
                megj_szoveg(megj_szoveg.Count - 1).ZOrder (0)
            End If
    End Select
   obj = Nothing
End Sub


Private Sub nevjegy_mnu_Click()
    nevjegy.Show vbModal
End Sub


Private Sub Beiras(Index As Integer)
    tabulalo = Index
    megjelol (Index)
If IrhatE(Index) Then
    szoveg.Visible = True
    szoveg.Move jel_szoveg(Index).Left, jel_szoveg(Index).Top, szoveg.Width, jel_szoveg(Index).Height
    szoveg.Text = jel_szoveg(Index).Caption
    If szoveg.Text = beallitasok.jel.Text Then
        szoveg.SelStart = 0
        szoveg.SelLength = Len(jel_szoveg(Index).Caption)
    End If
    szoveg.SetFocus
End If
End Sub

Private Sub nyomtat_Click()
    'nyomtatas.Load
    'nyomtatas.Show vbModal
    
    nyomtatas.NyomtatasiKep
End Sub

Private Sub szerkesztes_mnu_Click()
    Shell Vakterkep.Konyvtar & App.EXEName & ".exe -sz=" & Megnyitva, vbNormalFocus
    kilep_mnu_Click
End Sub

Private Sub szoveg_Change()
On Error Resume Next
    jel_szoveg(0).FontSize = szoveg.FontSize
    jel_szoveg(0).Caption = szoveg.Text
    szoveg.Width = jel_szoveg(0).Width + 200
    If szoveg.Width < 500 Then szoveg.Width = 500
End Sub

Private Sub szoveg_DragDrop(Source As Control, X As Single, Y As Single)
    szoveg.Visible = False
    jel_szoveg_DragDrop tabulalo, Source, X, Y
End Sub

Private Sub terulet_Click()
    Form_KeyDown 27, 0
End Sub

Private Sub ujra_mnu_Click()
Dim i As Integer
    proba = 1
    Form_KeyDown 27, 0
    ertekel_mnu.Caption = ertekeles.ertekeles
    For i = 1 To jel.Count - 1
        elemek(i).pont = 0
        elemek(i).Kovetkezo = 1
        elemek(i).Kitalalta = False
        jel(i).Bekapcsolva = True
        jel_szoveg(i).Enabled = True
        jel_szoveg(i).BackStyle = 0
        jel_szoveg(i).Caption = beallitasok.jel.Text
        igazit (i)
    Next i
End Sub
Private Sub igazit(Index As Integer)
'On Error Resume Next
    jel_szoveg(Index).Left = jel(Index).Left + Atalakit(elemek(Index).Bal, Index)
    jel_szoveg(Index).Top = jel(Index).Top + Atalakit(elemek(Index).Felso, Index)
End Sub

Private Sub Ujbehuzo(Elemhez As Integer)
    With behuzo
        Load .szoveg(.szoveg.Count)
        .szoveg(.szoveg.Count - 1).Top = .szoveg(.szoveg.Count - 2).Top + .szoveg(.szoveg.Count - 2).Height + 100
        .szoveg(.szoveg.Count - 1).Left = 100
        .szoveg(.szoveg.Count - 1).Caption = elemek(Elemhez).Cime
        .szoveg(.szoveg.Count - 1).Visible = True
    End With
End Sub
Public Sub Ujratolt(Fajlt As String)
    torol
    megnyitas (Fajlt)
End Sub
Private Function IrhatE(Index As Integer) As Boolean
    If proba > 1 And beallitasok.kerdesek.Value = 1 And elemek(Index).kerdesek(elemek(Index).Kovetkezo).Kerdes <> "" And Not elemek(Index).Kitalalta Then
        Dim ev As String
        ev = InputBox(elemek(Index).kerdesek(elemek(Index).Kovetkezo).Kerdes, Atalakit(KozosSzovegek(23), CStr(elemek(Index).Kovetkezo)), "")
        If LegegyszerubbAlak(ev) <> LegegyszerubbAlak(elemek(Index).kerdesek(elemek(Index).Kovetkezo).Valasz) Then
            If ev <> "" Then
                elemek(Index).Kovetkezo = elemek(Index).Kovetkezo + 1
            End If
            IrhatE = False
        Else
            elemek(Index).Kitalalta = True
            IrhatE = True
        End If
    Else
        IrhatE = True
    End If
End Function
Private Sub Ellenorzott(Index As Integer, Joe As Boolean)
    If beallitasok.megoldas(0).Value Then
        jel_szoveg(Index).BackStyle = 1
        If Joe Then
                jel_szoveg(Index).BackColor = vbGreen
            Else
                jel_szoveg(Index).BackColor = &HC0C0FF
        End If
    Else
        jel_szoveg(Index).Visible = Not Joe
        jel(Index).Visible = Not Joe
    End If
    
    jel_szoveg(Index).Enabled = Not Joe
    jel(Index).Bekapcsolva = Not Joe
End Sub
