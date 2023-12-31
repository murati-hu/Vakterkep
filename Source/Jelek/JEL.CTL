VERSION 5.00
Begin VB.UserControl jel 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   ControlContainer=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   615
   ScaleWidth      =   1725
   ToolboxBitmap   =   "jel.ctx":0000
   Begin VB.Line jel_ln 
      X1              =   720
      X2              =   1680
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Image jel_img 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   2
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1038
         SubFormatType   =   9
      EndProperty
      Height          =   540
      Left            =   840
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Shape jel_shp 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "jel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const egyseg = 8
Const Alapmeret = 500

Enum Jelek
    Negyszog
    Negyzet
    Ellipszis
    Kor
    KerekitettNegyszog
    KerekitettNegyzet
    'vonal
    kep
    vonal
    'Szoveg
End Enum

Enum Vonalak
    Nincs
    Teljes
    Szaggatott
    Pontozott
    PontosanSzagatott
    DuplaPontosanSzaggatott
End Enum

Enum kitoltes
    Teljes
    hatter
    VizszintesVonalak
    FuggolegesVonalak
    BalrolSregVonalak
    JobbrolSregVonalak
    Halos
    SregHalos
    'Atlatszo
End Enum


Private sJel As Jelek 'bels� jelek
Private sKepFajl As String ', sSzoveg As String 'Egy�b jelek adatai
Private sVonalAllas As Byte

Public Event Click()
Public Event Hiba(hibakod As Byte)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()
Public Event KeyPress(KeyAscii As Integer)


Private Sub jel_img_Click()
    RaiseEvent Click
End Sub

Private Sub jel_img_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub jel_img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'Private Sub jel_lbl_Click()
'    RaiseEvent Click
'End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub



Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    jel_img.Move 0, 0, Alapmeret, Alapmeret
    jel_shp.Move 0, 0, Alapmeret, Alapmeret
    'jel_lbl.Move 0, 0, Alapmeret, Alapmeret
    UserControl.Width = Alapmeret
    UserControl.Height = Alapmeret
End Sub

'Jelek kezel�se
Public Property Get jel() As Jelek
    If jel_shp.Visible Then jel = jel_shp.Shape
    If jel_img.Visible Then jel = kep
    If jel_ln.Visible Then jel = vonal
    'If jel_lbl.Visible Then Jel = Szoveg
End Property

Public Property Let jel(ByVal UjJel As Jelek)
jel_shp.Visible = False
jel_img.Visible = False
jel_ln.Visible = False

    Select Case UjJel
        Case 0 To 5
            jel_shp.Visible = True
            jel_shp.Shape = UjJel
        Case 6
            jel_img.Visible = True
        Case 7
            jel_ln.Visible = True
        'Case 7
        '    jel_lbl.Visible = True
        '    jel_lbl.Caption = Szoveg
    End Select
    UserControl_Resize
End Property

'Alakzatok kitolt�se
Public Property Get KitoltesTipus() As kitoltes
    'If jel_shp.BackStyle = 0 Then
     '       KitoltesTipus = Atlatszo
      '  Else
            KitoltesTipus = jel_shp.FillStyle
    'End If
End Property
Public Property Let KitoltesTipus(ByVal uj As kitoltes)
    'If uj = Atlatszo Then
     '       jel_shp.BackStyle = 0
      '  Else
            jel_shp.FillStyle = uj
            'jel_shp.BackStyle = 1
    'End If
End Property

'Kitolt�s sz�ne
Public Property Get KitoltesSzine() As ColorConstants
    KitoltesSzine = jel_shp.FillColor
End Property
Public Property Let KitoltesSzine(ByVal uj As ColorConstants)
    jel_shp.FillColor = uj
End Property

'Kit�lt�s kezel�se
Public Property Get Atlatszo() As Boolean
    Atlatszo = Not CBool(jel_shp.BackStyle)
End Property
Public Property Let Atlatszo(uj As Boolean)
    jel_shp.BackStyle = Abs(Not uj)
    UserControl_Resize
End Property


'Keret t�pusa
Public Property Get KeretTipus() As Vonalak
    KeretTipus = jel_shp.BorderStyle
End Property
Public Property Let KeretTipus(ByVal uj As Vonalak)
    jel_shp.BorderStyle = uj
    jel_ln.BorderStyle = uj
End Property

'Keret sz�ne
Public Property Get KeretSzine() As ColorConstants
    KeretSzine = jel_shp.BorderColor
End Property
Public Property Let KeretSzine(ByVal uj As ColorConstants)
    jel_shp.BorderColor = uj
    jel_ln.BorderColor = uj
End Property

'Keret vasatgs�ga
Public Property Get KeretVastagsaga() As Integer
    KeretVastagsaga = jel_shp.BorderWidth
End Property
Public Property Let KeretVastagsaga(ByVal uj As Integer)
    If uj > 0 And uj < 101 Then
        jel_shp.BorderWidth = uj
        jel_ln.BorderWidth = uj
        UserControl_Resize
    End If
End Property

'H�tt�r sz�ne
Public Property Get HatterSzine() As ColorConstants
    HatterSzine = jel_shp.BackColor
End Property
Public Property Let HatterSzine(ByVal uj As ColorConstants)
    jel_shp.BackColor = uj
End Property

'K�p ny�jt�sa
Public Property Get KepNyujtasa() As Boolean
    KepNyujtasa = jel_img.Stretch
End Property
Public Property Let KepNyujtasa(ByVal uj As Boolean)
    jel_img.Stretch = uj
    UserControl_Resize
End Property

'K�p helye
Public Property Get KepElerese() As String
    KepElerese = sKepFajl
End Property
Public Property Let KepElerese(ByVal uj As String)
On Error GoTo Toltesihiba
    sKepFajl = uj
    If uj <> "" Then
            jel_img.Picture = LoadPicture(sKepFajl)
            jel = kep
        Else
            jel_img.Picture = Nothing
    End If
    Exit Property
Toltesihiba:
    jel_img.Picture = Nothing
    RaiseEvent Hiba(1)
End Property

' K�p magass�g �s sz�less�g �rt�kei
Public Property Get KepMagassag() As Single
    KepMagassag = jel_img.Height
End Property
Public Property Get KepSzelesseg() As Single
    KepSzelesseg = jel_img.Width
End Property


Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        KepElerese = .ReadProperty("KepElerese", "")
        jel = .ReadProperty("Jel", Negyszog)
        Bekapcsolva = .ReadProperty("Bekapcsolva", True)
        Atlatszo = .ReadProperty("Atlatszo", False)
        KitoltesTipus = .ReadProperty("KitoltesTipus", kitoltes.hatter)
        KitoltesSzine = .ReadProperty("KitoltesSzine", vbWhite)
        KeretTipus = .ReadProperty("KeretTipus", Vonalak.Teljes)
        KeretSzine = .ReadProperty("KeretSzine", vbBlack)
        KeretVastagsaga = .ReadProperty("KeretVastagsaga", 1)
        HatterSzine = .ReadProperty("HatterSzine", vbWhite)
        VonalAllas = .ReadProperty("VonalAllas", 1)
    End With
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    jel_shp.Move jel_shp.BorderWidth * egyseg, jel_shp.BorderWidth * egyseg, UserControl.ScaleWidth - (jel_shp.BorderWidth * 2 * egyseg), UserControl.ScaleHeight - (jel_shp.BorderWidth * 2 * egyseg)
    jel_img.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    With jel_ln
    Dim torzit As Single
    torzit = 20 * KeretVastagsaga
    .x1 = 0 + torzit
    Select Case sVonalAllas
        Case 0
            .y1 = 0 + torzit
            .X2 = UserControl.ScaleWidth - torzit
            .Y2 = 0 + torzit
        Case 1
            .y1 = 0 + torzit
            .X2 = 0 + torzit
            .Y2 = UserControl.ScaleHeight - torzit
        Case 2
            .y1 = 0 + torzit
            .X2 = UserControl.ScaleWidth - torzit
            .Y2 = UserControl.ScaleHeight - torzit
        Case 3
            .y1 = UserControl.ScaleHeight - torzit
            .X2 = UserControl.ScaleWidth - torzit
            .Y2 = 0 + torzit
    End Select
    End With
End Sub

Public Function BalKozep(Bal As Single)
    BalKozep = Bal - UserControl.Width / 2
End Function
Public Function FelsoKozep(Felso As Single)
    FelsoKozep = Felso - UserControl.Height / 2
End Function
Public Property Get Bekapcsolva() As Boolean
    Bekapcsolva = UserControl.Enabled
End Property
Public Property Let Bekapcsolva(uj As Boolean)
    UserControl.Enabled = uj
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "KepElerese", KepElerese, ""
        .WriteProperty "Jel", jel, Negyszog
        .WriteProperty "Bekapcsolva", Bekapcsolva, True
        .WriteProperty "Atlatszo", Atlatszo, False
        .WriteProperty "KitoltesTipus", KitoltesTipus, kitoltes.hatter
        .WriteProperty "KitoltesSzine", KitoltesSzine, vbWhite
        .WriteProperty "KeretTipus", KeretTipus, Vonalak.Teljes
        .WriteProperty "KeretSzine", KeretSzine, vbBlack
        .WriteProperty "KeretVastagsaga", KeretVastagsaga, 1
        .WriteProperty "HatterSzine", HatterSzine, vbWhite
        .WriteProperty "VonalAllas", VonalAllas, 1
    End With
End Sub
Public Property Get VonalAllas() As Byte
    VonalAllas = sVonalAllas
End Property
Public Property Let VonalAllas(uj As Byte)
    uj = uj Mod 4
    
    sVonalAllas = uj
    UserControl_Resize
End Property
