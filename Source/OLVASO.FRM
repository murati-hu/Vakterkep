VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form terkep 
   BackColor       =   &H8000000C&
   Caption         =   "Vaktérkép"
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
      Begin Vakterkep2.jel jel 
         Height          =   135
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
         _extentx        =   873
         _extenty        =   873
         kitoltesszine   =   -2147483640
         keretszine      =   -2147483640
         hatterszine     =   -2147483643
      End
      Begin Vakterkep2.jel megj 
         Height          =   135
         Index           =   0
         Left            =   2880
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   135
         _extentx        =   873
         _extenty        =   873
         kitoltesszine   =   -2147483640
         keretszine      =   -2147483640
         hatterszine     =   -2147483643
      End
      Begin VB.Label jel_szoveg 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Jel szöveg"
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
         Caption         =   "Jel szöveg"
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
      Caption         =   "&Fájl"
      Begin VB.Menu ujra_mnu 
         Caption         =   "Újra kezd"
         Enabled         =   0   'False
         Shortcut        =   ^U
      End
      Begin VB.Menu megnyit_mnu 
         Caption         =   "Megnyitás"
         Shortcut        =   ^M
      End
      Begin VB.Menu nyomtat 
         Caption         =   "Nyomtatás..."
         Shortcut        =   ^P
      End
      Begin VB.Menu v7 
         Caption         =   "-"
      End
      Begin VB.Menu szerkesztes_mnu 
         Caption         =   "Szerkesztés..."
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu v1 
         Caption         =   "-"
      End
      Begin VB.Menu kilep_mnu 
         Caption         =   "Kilépés"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu eszkozok_mnu 
      Caption         =   "&Eszközök"
      Begin VB.Menu jelm_mnu 
         Caption         =   "Jelmagyarázat"
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
      Begin VB.Menu behuzo_mnu 
         Caption         =   "Behúzó"
         Enabled         =   0   'False
         Shortcut        =   ^H
      End
      Begin VB.Menu v6 
         Caption         =   "-"
      End
      Begin VB.Menu ertekel_mnu 
         Caption         =   "Értékelés"
         Shortcut        =   ^E
      End
      Begin VB.Menu beall_mnu 
         Caption         =   "Beállítások..."
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu sugo_mnu 
      Caption         =   "&Súgó"
      Begin VB.Menu help_mnu 
         Caption         =   "Súgó"
         Shortcut        =   {F1}
      End
      Begin VB.Menu v2 
         Caption         =   "-"
      End
      Begin VB.Menu nevjegy_mnu 
         Caption         =   "Névjegy"
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

For i = 1 To (jel.Count - 1)
    If elemek(i).pont = 0 Then
        If Trim(LCase(elemek(i).Cime)) = Trim(LCase(jel_szoveg(i).Caption)) Then
            jel_szoveg(i).BackStyle = 1
            jel_szoveg(i).BackColor = vbGreen
            jel_szoveg(i).Enabled = False
            jel(i).Bekapcsolva = False
            
            If proba = 1 Or beallitasok.kerdesek.Value = 0 Then
                elemek(i).pont = beallitasok.pont.Text
            Else
                elemek(i).pont = beallitasok.pont.Text * ((100 - ((elemek(i).Kovetkezo) * beallitasok.levonas.Text)) / 100)
            End If
            jo = jo + 1
        Else
            jel_szoveg(i).BackStyle = 1
            jel_szoveg(i).BackColor = &HC0C0FF
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

Select Case CByte(ertekeles.szazalek)
    Case 0 To beallitasok.hatarok(1) - 1
        ertekeles.jegy = 1
        ertekeles.neve = "Elégtelen"
       
    Case beallitasok.hatarok(1) To beallitasok.hatarok(2) - 1
        ertekeles.jegy = 2
        ertekeles.neve = "Elégséges"
        
    Case beallitasok.hatarok(2) To beallitasok.hatarok(3) - 1
        ertekeles.jegy = 3
        ertekeles.neve = "Közepes"
        
    Case beallitasok.hatarok(3) To beallitasok.hatarok(4) - 1
        ertekeles.jegy = 4
        ertekeles.neve = "Jó"
        
    Case beallitasok.hatarok(4) To 100
        ertekeles.jegy = 5
        ertekeles.neve = "Példás"
End Select

If proba = 5 Then
        ertekel_mnu.Caption = "Végsõ eredmény"
        ertekeles.Caption = "Végsõ eredmény"
    Else
        ertekel_mnu.Caption = proba + 1 & ". Értékelés"
        ertekeles.Caption = proba & ". Értékelés"
        proba = proba + 1
End If
ertekeles.Show vbModal
End Sub

Private Sub fl_Change()
    terulet.Top = fl.Value
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
    MsgBox "Az egyik szimbólum hibás vagy nem található!", vbInformation, "Hibás szimbólum"
    obj.jel = 1
End Sub

Private Sub jel_szoveg_Click(Index As Integer)
    jel_Click (Index)
End Sub

Private Sub jel_szoveg_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    jel_szoveg(Index).Caption = Source.Caption
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






Private Sub megnyit_mnu_Click()
On Error GoTo megse
    pb.CancelError = True
    pb.DialogTitle = "Vaktérkép megnyitása ..."
    pb.Filter = "Vaktérkép fájlok (*.vtk)|*.vtk"
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
Fajlnev = Atalakit(Fajlnev, "")
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
            
            'Parancs formázgatása
            If kulcsszo = "" Then kulcsszo = sor
            kulcsszo = Korulmetel(kulcsszo)
            parameter = Korulmetel(parameter)
            
            'MsgBox kulcsszo & " := " & parameter
        
            
            Select Case kulcsszo
' ######################## Projektekkel összefüggõ beállítások #################################
                Case "cim", "cime"
                    Cime = parameter
                    Me.Caption = parameter & " - " & "Vaktérkép " & Vakterkep.Verzio
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
                    
                    terulet.Width = szel
                    terulet.Height = mag
                    
                    terulet.Cls
                    terulet.PaintPicture LoadPicture(kep), 0, 0, terulet.Width, terulet.Height, X, Y, szel, mag
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
                    
                    
                
                    'alapértelmezések:
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
                        obj.Left = CSng(Atalakit(Kicsontoz(parameter, ",", 0), ""))
                        obj.Top = CSng(Atalakit(Kicsontoz(parameter, ",", 1), ""))
                    End If
                Case "meret", "meretek"
                    On Error Resume Next
                    If aze Then
                        obj.Width = CLng(Atalakit(Kicsontoz(parameter, ",", 0), ""))
                        obj.Height = CLng(Atalakit(Kicsontoz(parameter, ",", 1), ""))
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
                        'Case "jel"
                            'jel_szoveg(id).BackStyle = 1
                            'jel_szoveg(id).BackColor = parameter
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
                        
                ' ########## ELEM EGYEDITULAJDONSÁGAI
                Case "kerdes"
                    If obj.Name = "jel" And KID < 10 Then
                        KID = KID + 1
                        elemek(id).kerdesek(KID).Kerdes = Kicsontoz(parameter, "|", 0)
                        elemek(id).kerdesek(KID).Valasz = Kicsontoz(parameter, "|", 1)
                    End If
                Case "jel", "alakzat"
                    If aze Then
                        obj.jel = CByte(parameter) Mod 7
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
                Case "lathatatlan-jel"
                    'If obj.Name <> "jelm" Then  'And megj_szoveg(FID).Visible Then
                        obj.Visible = False
                    'End If
                Case "lathatatlan-szoveg"
                    If obj.Name = "megj" Then megj_szoveg(FID).Visible = False
' ######################## Beállításokkal kapcsolatos értékek ###################################
                Case "hatarok", "ponthatarok"
                    If beallitasok.egyeni.Value = 1 Or Fajlnev = Atalakit("$vt\vakterkep.ini", "") Then
                        For i = 1 To 4
                            beallitasok.hatarok(i) = Kicsontoz(parameter, ",", i - 1)
                        Next i
                    End If
                Case "beallitas", "beallitasok_engedelyezese"
                    beallitasok.enged.Value = parameter
                    beall_mnu.Visible = Abs(CInt(parameter))
            
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
                'Case "nyelv"
                '    If parameter <> "0" Then
                '        beallitasok.nyelv_chk.Value = 1
                '        For i = 0 To beallitasok.nyelvek.ListCount - 1
                '            If parameter = beallitasok.nyelvek.List(i) Then
                '                nyelv (parameter)
                '                beallitasok.nyelvek.Selected(i) = True
                '            End If
                '        Next i
                '    Else
                '        beallitasok.nyelv_chk.Value = 0
                '    End If

            End Select
kihagy:
        Loop
        Close 1
    
    If aze Then lezaras
    
    If jel.Count > 1 Then ertekel_mnu.Enabled = True
    
    If jelmagyarazat.jelm.Count > 1 Then
        jelm_mnu.Enabled = True
        jelmagyarazat.Height = (Nagyobb(jelmagyarazat.jelm(jelmagyarazat.jelm.Count - 1).Top + jelmagyarazat.jelm(jelmagyarazat.jelm.Count - 1).Height, jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Top + jelmagyarazat.jelm_szoveg(jelmagyarazat.jelm_szoveg.Count - 1).Height)) + 610 '* 1.5 '* (jelmagyarazat.Height / jelmagyarazat.ScaleHeight)
        For i = 1 To JID
            If jelmagyarazat.Width < (200 + jelmagyarazat.jelm_szoveg(i).Left + jelmagyarazat.jelm_szoveg(i).Width) Then '* (jelmagyarazat.ScaleWidth / jelmagyarazat.Width) Then
                jelmagyarazat.Width = (200 + jelmagyarazat.jelm_szoveg(i).Left + jelmagyarazat.jelm_szoveg(i).Width) + 120 '* (jelmagyarazat.ScaleWidth / jelmagyarazat.Width)
            End If
        Next i
        'else
    End If
    
    If jel.Count > 1 And beallitasok.behuzas.Value = 1 Then
            behuzo_mnu.Enabled = True
        Else
            behuzo_mnu.Enabled = False
    End If
    
    
    terulet.Visible = True
    If jel.Count > 1 Then ujra_mnu.Enabled = True
    If Fajlnev <> Atalakit("$vt\vakterkep.ini", "") Then
            For i = 1 To jel.Count - 1 Step 2
                Ujbehuzo (i)
            Next i
            
            For i = 2 To jel.Count - 1 Step 2
                Ujbehuzo (i)
            Next i
            
            For i = 1 To behuzo.szoveg.Count - 1
                If behuzo.szoveg(i).Width + 200 > behuzo.Width - 120 Then
                    behuzo.Width = (behuzo.szoveg(i).Width + 300) + 120 '* (behuzo.Width / behuzo.ScaleWidth)
                End If
            Next i
            
            behuzo.Height = (behuzo.szoveg(behuzo.szoveg.Count - 1).Top + behuzo.szoveg(behuzo.szoveg.Count - 1).Height + 610) '* 1.5 '(behuzo.Height / (behuzo.ScaleHeight))

            Form_KeyDown 40, 0
            Megnyitva = Fajlnev
            szerkesztes_mnu.Enabled = CBool(beallitasok.szerkesztes.Value)
    End If
Exit Sub
Hiba:
    Select Case Err.Number
        Case 52
            MsgBox "A '" & Fajlnev & "' fájl nem tölthetõ be.", vbInformation, "Hábás fájl adott meg!"
            'Exit Sub
        Case 53
            If Fajlnev <> Vakterkep.Konyvtar & "vakterkep.ini" Then
                MsgBox "A megadott fájl nem található!", vbCritical, "A megadott fájl nem található!"
            End If
            'Exit Sub
        Case Else
            MsgBox "A megadott projekt hibás bejegyzéseket tartalmaz, ami bizonytalanná teheti a program futását. A töltés e hiba miatt megszakad", vbCritical, "Ismeretlen hiba"
            MsgBox "A hiba oka: " & Err.Description, vbInformation, "Hiba(" & Err.Number & ")"
        End Select
    
    Close 1
    Exit Sub
    
kephiba:
    MsgBox "A megadott kép hibás, ismeretlen tömörítésû vagy nem található a megadott helyen." & vbCrLf & _
           "A projekt betöltése e hiba miatt megszakad.", vbCritical, "Képbetöltési hiba"
    MsgBox Err.Description & " - " & Err.Number, vbCritical, kep
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
    Me.Caption = "Vaktérkép " & Vakterkep.Verzio
    terulet.Picture = nevjegy.kep.Picture
    terulet.Visible = True
    
    jelmagyarazat.Height = 510
    jelmagyarazat.Width = 1560
    Cime = ""
    Megnyitva = ""
    tabulalo = 0
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
    
    If proba > 1 And beallitasok.kerdesek.Value = 1 And elemek(Index).kerdesek(elemek(Index).Kovetkezo).Kerdes <> "" And Not elemek(Index).Kitalalta Then
        Dim ev As String
        ev = InputBox(elemek(Index).kerdesek(elemek(Index).Kovetkezo).Kerdes, elemek(Index).Kovetkezo & ". kérdés", "")
        If Trim(LCase(ev)) <> Trim(LCase(elemek(Index).kerdesek(elemek(Index).Kovetkezo).Valasz)) Then
            If ev <> "" Then
                elemek(Index).Kovetkezo = elemek(Index).Kovetkezo + 1
            End If
            Exit Sub
        Else
            elemek(Index).Kitalalta = True
            GoTo irhat
        End If
    Else
irhat:
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
    nyomtatas.Show vbModal
    'Printer.ScaleMode = 1
    'terkep.Refresh
    'terkep.PrintForm
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

Private Sub terulet_Click()
    Form_KeyDown 27, 0
End Sub

Private Sub ujra_mnu_Click()
Dim i As Integer
    proba = 1
    Form_KeyDown 27, 0
    ertekel_mnu.Caption = "Értékelés"
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
