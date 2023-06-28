VERSION 5.00
Begin VB.UserControl szoveg 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PaletteMode     =   4  'None
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Windowless      =   -1  'True
   Begin VB.Shape keret 
      FillColor       =   &H0000FFFF&
      Height          =   1815
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line alahuzas 
      BorderStyle     =   0  'Transparent
      X1              =   600
      X2              =   3360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label karakter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "szoveg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const RaHagyas = 20
Private alfa As Double
Private p_betukoz As Double


Public Event Click()
'Public Event Hiba(hibakod As Byte)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick()
Public Event KeyPress(KeyAscii As Integer)
'Private lathato As Boolean

Private Sub felso_Click()
    RaiseEvent Click
End Sub

Private Sub karakter_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub karakter_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Private Sub karakter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'Tulajdonságok
'Betûtípusok - Neve
Private Sub UserControl_Initialize()
    ForeColor = vbBlack
    BackColor = vbButtonFace
    BackStyle = 0
    
    Betukoz = -10
    Caption = UserControl.Name
    Forgatas = 90
    Kirajzol
End Sub
Public Property Get Caption() As String
    Caption = ""
    Dim i As Long
    For i = 0 To karakter.Count - 1
        Caption = Caption & karakter(i).Caption
    Next i
End Property
Public Property Let Caption(Uj As String)
    Dim i As Long
    
    For i = 1 To karakter.Count - 1
        Unload karakter(i)
    Next i
    
    karakter(0).Caption = Mid(Uj, 1, 1)
    For i = 2 To Len(Uj)
        Load karakter(i - 1)
        karakter(i - 1).Caption = Mid(Uj, i, 1)
        karakter(i - 1).Move 0, 0
        karakter(i - 1).Visible = True
    Next i
    Kirajzol
End Property
Public Property Get Forgatas() As Double
    Forgatas = alfa
End Property
Public Property Let Forgatas(Uj As Double)
    alfa = Uj
    Kirajzol
End Property
Private Sub Kirajzol()
    Dim i As Long
    Dim szog As Double, sugar As Double
    Select Case alfa
            Case 0
                For i = karakter.Count - 1 To 0 Step -1
                    karakter(i).Left = 0
                    karakter(i).Top = (karakter(0).Height + Betukoz) * (karakter.Count - 1 - i)
                Next i
                MeretreSzab
                For i = 0 To karakter.Count - 1
                    karakter(i).Left = (UserControl.Width - karakter(i).Width) / 2
                Next i
                
                MeretreSzab
                alahuzas.X1 = karakter(0).Left
                alahuzas.Y1 = karakter(0).Top
        
                alahuzas.X2 = karakter(karakter.Count - 1).Left
                alahuzas.Y2 = karakter(karakter.Count - 1).Top + karakter(karakter.Count - 1).Height
            Case 90
                'On Error Resume Next
                karakter(0).Move 0, 0
                For i = 1 To karakter.Count - 1
                    karakter(i).Left = karakter(i - 1).Left + karakter(i - 1).Width + Betukoz
                    karakter(i).Top = 0
                Next i
                MeretreSzab
                alahuzas.X1 = karakter(0).Left
                alahuzas.Y1 = karakter(0).Top + karakter(0).Height
        
                alahuzas.X2 = karakter(karakter.Count - 1).Left + karakter(karakter.Count - 1).Width
                alahuzas.Y2 = karakter(karakter.Count - 1).Top + karakter(karakter.Count - 1).Height
            Case 180
                For i = 0 To karakter.Count - 1
                    karakter(i).Left = 0
                    karakter(i).Top = (karakter(0).Height + Betukoz) * i
                Next i
                MeretreSzab
                For i = 0 To karakter.Count - 1
                    karakter(i).Left = (UserControl.Width - karakter(i).Width) / 2
                Next i
                MeretreSzab
                alahuzas.X1 = karakter(0).Left
                alahuzas.Y1 = karakter(0).Top
        
                alahuzas.X2 = karakter(karakter.Count - 1).Left
                alahuzas.Y2 = karakter(karakter.Count - 1).Top + karakter(karakter.Count - 1).Height
            Case 90 To 180
                'Dim szog As Double
                szog = Radianba(alfa - 90)
                With karakter
                    'sugar = karakter(.Count - 1).Left + karakter(.Count - 1).Width
                    sugar = 0
                    karakter(0).Move 0, 0
                    
                    For i = 1 To .Count - 1
                        'karakter(i).Top = karakter(i - 1).Top + ((Sin(szog) * karakter(i - 1).Width) / Cos(szog)) + (Sin(szog) * Betukoz)
                        'karakter(i).Left = karakter(i - 1).Left + karakter(i - 1).Width + (Cos(szog) * Betukoz)
                        
                        sugar = sugar + karakter(i - 1).Width + Betukoz
                        karakter(i).Top = Sin(szog) * sugar
                        karakter(i).Left = Cos(szog) * sugar
                    Next i
                End With
                MeretreSzab
                
                alahuzas.X1 = karakter(0).Left
                alahuzas.Y1 = karakter(0).Top + karakter(0).Height
    
                If alfa < 120 Then
                    alahuzas.X2 = karakter(karakter.Count - 1).Left + karakter(karakter.Count - 1).Width
                Else
                    alahuzas.X2 = karakter(karakter.Count - 1).Left
                End If
                alahuzas.Y2 = karakter(karakter.Count - 1).Top + karakter(karakter.Count - 1).Height
            Case 0 To 90
                
                szog = Radianba(alfa + 90)
                With karakter
                    'sugar = karakter(.Count - 1).Left + karakter(.Count - 1).Width
                    sugar = 0
                    karakter(0).Move 0, 0
                    
                    For i = 1 To .Count - 1
                        sugar = sugar + karakter(i - 1).Width + Betukoz
                        karakter(i).Top = -1 * Sin(szog) * sugar 'karakter(i - 1).Top + ((Sin(szog) * karakter(i - 1).Width) / Cos(szog)) - (Sin(szog) * Betukoz)
                        karakter(i).Left = -Cos(szog) * sugar 'karakter(i - 1).Left + karakter(i - 1).Width - (Cos(szog) * Betukoz)
                    Next i
                'eltolás
                        
                        For i = 0 To .Count - 1
                            karakter(i).Top = karakter(i).Top - karakter(.Count - 1).Top
                        Next i
                End With
                MeretreSzab
                If alfa < 70 Then
                    alahuzas.X1 = karakter(0).Left + karakter(0).Width
                Else
                    alahuzas.X1 = karakter(0).Left
                End If
                    alahuzas.Y1 = karakter(0).Top + karakter(0).Height
        
                    alahuzas.X2 = karakter(karakter.Count - 1).Left + karakter(karakter.Count - 1).Width
                    alahuzas.Y2 = karakter(karakter.Count - 1).Top + karakter(karakter.Count - 1).Height
        End Select
        MeretreSzab
       'UserControl.Print "hello"
       'keret.Print "HEllo"
End Sub
Public Property Get Betukoz() As Double
    Betukoz = p_betukoz
End Property
Public Property Let Betukoz(Uj As Double)
    p_betukoz = Uj
    Kirajzol
End Property
Private Function Legszelesebb()
    Dim i As Long
    Legszelesebb = 0
    For i = 0 To karakter.Count - 1
        If karakter(i).Left + karakter(i).Width > Legszelesebb Then
            Legszelesebb = karakter(i).Left + karakter(i).Width
        End If
    Next i
    Legszelesebb = Legszelesebb + RaHagyas
End Function
Private Function Legmagasabb()
    Dim i As Long
    Legmagasabb = 0
    For i = 0 To karakter.Count - 1
        If karakter(i).Top + karakter(i).Height > Legmagasabb Then
            Legmagasabb = karakter(i).Top + karakter(i).Height
        End If
    Next i
    Legmagasabb = Legmagasabb + RaHagyas
End Function
Private Function Radianba(Fok As Double)
    Radianba = (Fok * 3.1415) / 180
End Function
Private Sub MeretreSzab()
    UserControl.Width = Legszelesebb
    UserControl.Height = Legmagasabb
End Sub
Public Property Get FontBold() As Boolean
    FontBold = karakter(0).FontBold
End Property
Public Property Let FontBold(Uj As Boolean)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).FontBold = Uj
    Next i
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = karakter(0).FontItalic
End Property
Public Property Let FontItalic(Uj As Boolean)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).FontItalic = Uj
    Next i
End Property

Public Property Get FontName() As String
    FontName = karakter(0).FontName
End Property
Public Property Let FontName(Uj As String)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).FontName = Uj
    Next i
End Property
Public Property Get FontSize() As Single
    FontSize = karakter(0).FontSize
End Property
Public Property Let FontSize(Uj As Single)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).FontSize = Uj
    Next i
    'alahuzas.BorderWidth = karakter(0).FontSize \ 6
    
    Kirajzol
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = karakter(0).FontStrikethru
End Property
Public Property Let FontStrikethru(Uj As Boolean)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).FontStrikethru = Uj
    Next i
End Property


Public Property Get FontUnderline() As Integer
    FontUnderline = alahuzas.BorderStyle
End Property
Public Property Let FontUnderline(Uj As Integer)
    alahuzas.BorderStyle = Uj
End Property
Public Property Get ForeColor() As ColorConstants
    ForeColor = karakter(0).ForeColor
End Property
Public Property Let ForeColor(Uj As ColorConstants)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).ForeColor = Uj
    Next i
    
    alahuzas.BorderColor = Uj
End Property
Public Property Get BackColor() As ColorConstants
    BackColor = karakter(0).BackColor
End Property

Public Property Let BackColor(Uj As ColorConstants)
    Dim i As Long
    
    For i = 0 To karakter.Count - 1
        karakter(i).BackColor = Uj
    Next i
    
    UserControl.BackColor = Uj
End Property
Public Property Get BackStyle() As Byte
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(Uj As Byte)
    'Dim i As Long
    
    'For i = 0 To karakter.Count - 1
    '    karakter(i).BackStyle = Uj
    'Next i
    
    UserControl.BackStyle = Uj
    'UserControl.Refresh
End Property

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Public Property Get BorderStyle() As Byte
    BorderStyle = Abs(CInt(keret.Visible)) 'karakter(0).BorderStyle UserControl.BorderStyle
End Property

Public Property Let BorderStyle(Uj As Byte)
    'Dim i As Long
    
    'For i = 0 To karakter.Count - 1
    '    karakter(i).BorderStyle = Uj
    'Next i
    'UserControl.BorderStyle = Uj
    'UserControl.Refresh
    keret.Visible = CBool(Uj)
End Property

Private Function Helyesbit(Mit As Variant, Hany As Byte) As Variant
    Helyesbit = Mit Mod Hany
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    keret.Move 0, 0
    keret.Width = UserControl.Width
    keret.Height = UserControl.Height
    
    'felso.Move 0, 0
    'felso.Width = keret.Width
    'felso.Height = keret.Height
    
End Sub


