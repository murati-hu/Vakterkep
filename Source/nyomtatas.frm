VERSION 5.00
Begin VB.Form nyomtatas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vaktérkép nyomtatása"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   Icon            =   "nyomtatas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame meret 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
      Begin VB.TextBox szelesseg 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox magassag 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Magasság:"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   780
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Szélesség:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   765
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CommandButton megse 
      Caption         =   "&Bezár"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Nyomtat"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton sugo 
      Caption         =   "Súgó"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox nyomtatando 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "nyomtatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arany As Double 'Az eredti képméret és a minta közötti arány
Const nyomtatoarany = 0.01789 'nyomtató képpontjai és a Twipek közötti arány

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
On Error Resume Next
    Me.Icon = terkep.Icon
    Me.Caption = terkep.Cime & " nyomtása"
    If terkep.terulet.Width > terkep.terulet.Height Then
                arany = terkep.terulet.Width / nyomtatando.Width
                nyomtatando.Height = terkep.terulet.Height / arany
        Else
                arany = terkep.terulet.Height / nyomtatando.Height
                nyomtatando.Width = terkep.terulet.Width / arany
    End If
    Vazol ("kepre")
    szelesseg.Text = nyomtatando.ScaleWidth * nyomtatoarany
    
    Me.Height = nyomtatando.Top + nyomtatando.Height + 610 + 660
    Me.Width = 2 * nyomtatando.Left + nyomtatando.Width + 120
    
    Form_Resize
    If megse.Left < meret.Left + meret.Width Then
        Me.Width = Me.Width + meret.Left + meret.Width - megse.Left + 100
        nyomtatando.Left = (Me.ScaleWidth - nyomtatando.Width) / 2
        Form_Resize
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    sugo.Move 100, nyomtatando.Top + nyomtatando.Height + 200
    meret.Move sugo.Left + sugo.Width + 100, sugo.Top - ((meret.Height - sugo.Height) / 2)
    ok.Move Me.ScaleWidth - 100 - ok.Width, sugo.Top
    megse.Move ok.Left - 200 - megse.Width, sugo.Top
End Sub

Private Sub magassag_Change()
On Error Resume Next
    szelesseg.Text = (((magassag.Text / nyomtatoarany) / nyomtatando.ScaleHeight) * nyomtatando.ScaleWidth) * nyomtatoarany
End Sub

Private Sub megse_Click()
    Unload Me
End Sub

Private Sub ok_Click()
    Vazol ("nyomtatora")
End Sub
Private Sub Vazol(mire As String)
    Dim i As Integer, j As Integer
    Dim hova As Object, nagyit As Double, db As Integer, sz As String
    On Error GoTo Hiba
    
    With terkep
        If mire = "kepre" Then
                Set hova = Me.nyomtatando
                nagyit = 1
            Else
                Set hova = Printer
                nagyit = (CDbl(szelesseg.Text) / nyomtatoarany) / nyomtatando.ScaleWidth
        End If
    
        'Kép megalkotása:
        db = 0
        
        hova.PaintPicture .terulet.Picture, 0, 0, nyomtatando.Width * nagyit, nyomtatando.Height * nagyit
    
        For i = 1 To .jel.Count - 1
            db = db + 1
            sz = .jel_szoveg(i).Caption
            If .jel(i).jel <> 6 Then
                For j = 1 To Int(.jel(i).Width / (2.5 * arany) * nagyit)
                    hova.Circle ((.jel(i).Left + .jel(i).Width / 2) / arany * nagyit, (.jel(i).Top + .jel(i).Height / 2) / arany * nagyit), j, vbBlack
                Next j
            Else
                hova.PaintPicture LoadPicture(.jel(i).KepElerese), .jel(i).Left / arany * nagyit, .jel(i).Top / arany * nagyit, .jel(i).Width / arany * nagyit, .jel(i).Height / arany * nagyit
            End If
                
            
            'Kikérdezendõ elem sorszáma
            hova.CurrentX = (.jel_szoveg(i).Left + .jel_szoveg(i).Width / 2) / arany * nagyit
            hova.CurrentY = (.jel_szoveg(i).Top + .jel_szoveg(i).Height / 2) / arany * nagyit
            hova.FontSize = .jel_szoveg(i).FontSize / arany * nagyit
            hova.FontBold = True
            hova.Print db
        Next i
    
        For i = 1 To .megj.Count - 1
            'Megjegyzések szövegeinek kirajzolása
            If .megj_szoveg(i).Visible Then
                hova.ForeColor = .megj_szoveg(i).ForeColor
                hova.CurrentX = .megj_szoveg(i).Left / arany * nagyit
                hova.CurrentY = .megj_szoveg(i).Top / arany * nagyit
                hova.FontName = .megj_szoveg(i).FontName
                hova.FontSize = .megj_szoveg(i).FontSize / arany * nagyit
                hova.FontBold = .megj_szoveg(i).FontBold
                hova.FontItalic = .megj_szoveg(i).FontItalic
                hova.FontUnderline = .megj_szoveg(i).FontUnderline
                hova.FontStrikethru = .megj_szoveg(i).FontStrikethru
                hova.Print .megj_szoveg(i).Caption
            End If
            
            'megjegyzés kirajolása
            If .megj(i).Visible Then
                If .megj(i).jel <> 6 Then
                    For j = 1 To Int(.megj(i).Width / (2.5 * arany) * nagyit)
                        hova.Circle ((.megj(i).Left + .megj(i).Width / 2) / arany * nagyit, (.megj(i).Top + .megj(i).Height / 2) / arany * nagyit), j, .megj_szoveg(i).ForeColor
                    Next j
                Else
                    hova.PaintPicture LoadPicture(.megj(i).KepElerese), .megj(i).Left / arany * nagyit, .megj(i).Top / arany * nagyit, .megj(i).Width / arany * nagyit, .megj(i).Height / arany * nagyit
                End If
            End If
        Next i
    End With
    
    If mire = "nyomtatora" Then Printer.EndDoc 'Ha nyomtatóra küldtük, akkor lezárni a csatornát
    Exit Sub
Hiba:
    Dim hibauzenet
    Select Case Err.Number
        Case 482
            hibauzenet = "nincs nyomtató telepítve."
        Case 380, 13
            hibauzenet = "hibás nyomtatási méretet adott meg."
        Case 6
            hibauzenet = "túl nagy nyomtatási méretet adott meg."
        Case Else
            hibauzenet = Err.Description
    End Select
    MsgBox "A projekt nyomtatása meghiúsult, mert " & hibauzenet, vbCritical, "Nyomtatási hiba (" & Err.Number & ")"
End Sub

Private Sub sugo_Click()
    HHSugo ("nyomtat.htm")
End Sub

Private Sub szelesseg_Change()
On Error Resume Next
    magassag = (((szelesseg.Text / nyomtatoarany) / nyomtatando.ScaleWidth) * nyomtatando.ScaleHeight) * nyomtatoarany
End Sub
