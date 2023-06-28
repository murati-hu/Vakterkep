VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form csomagolo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Projekt csomagoló"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "csomagolo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox log 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1560
      Width           =   4455
   End
   Begin VB.CommandButton ment 
      Caption         =   "#"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox cel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton csomagol 
      Caption         =   "Projekt csomagolása"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
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
      Height          =   290
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox forras 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog pb 
      Left            =   4320
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      Caption         =   "Csomag:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      Caption         =   "Projekt:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "csomagolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Private Sub cel_Change()
    If cel.Text <> "" Then
        csomagol.Enabled = True
    End If
End Sub

Public Sub csomagol_Click()
On Error GoTo hiba
    Osszepakol forras.Text, cel.Text
Exit Sub
hiba:
    loggol "HIBA:" & Err.Description
    loggol "A csomagolás megszakadt."
End Sub

Private Sub Form_DblClick()
    'ShellExecute Me.hwnd, "open", "cmd.exe", "", "c:\windows\system32\", SW_SHOWNORMAL
     ShellExecute Me.hwnd, "run", "cmd.exe", "", "c:\", SW_SHOWNORMAL
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tomorito = Nothing
    End
End Sub

Private Sub forras_Change()
    If forras.Text = "" Then
        ment.Enabled = False
        csomagol.Enabled = False
    Else
        ment.Enabled = True
    End If
End Sub

Private Sub ment_Click()
On Error GoTo megse
        pb.DialogTitle = "Csomag mentése"
        pb.Filter = "ZIP Arhívum(*.zip)|*.zip"
        pb.FileName = ""
        pb.ShowSave
        cel.Text = pb.FileName
megse:
End Sub

Private Sub talloz_Click()
On Error GoTo megse
        pb.DialogTitle = "Projekt megnyitása..."
        pb.Filter = "Vakablak projektek(*.vtk)|*.vtk"
        pb.FileName = "*.vtk"
        pb.ShowOpen
        forras.Text = pb.FileName
megse:
End Sub
Private Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lRet As Long
    
    sBuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sBuffer)
    
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    GetTempPathName = sBuffer
    
End Function

Public Function utasitas(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    utasitas = Mid(Adatsor, 1, i - 1)
        End If
        i = i + 1
    Loop
    If Not megvan Then utasitas = Adatsor
End Function
Public Function Ertek(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    Ertek = Mid(Adatsor, i + 1, Len(Adatsor) - i)
        End If
        i = i + 1
    Loop
    If Not megvan Then Ertek = ""
End Function
Public Function Korulmetel(Szoveg As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Szoveg) And Not megvan
        If Mid(Szoveg, i, 1) <> Chr(9) And Mid(Szoveg, i, 1) <> " " Then
            megvan = True
            Szoveg = Mid(Szoveg, i, Len(Szoveg) - i + 1)
        End If
        i = i + 1
    Loop
    megvan = False
    i = Len(Szoveg)
    Do While i >= 1 And Not megvan
        If Mid(Szoveg, i, 1) <> Chr(9) And Mid(Szoveg, i, 1) <> " " Then
            megvan = True
            Szoveg = Mid(Szoveg, 1, i)
        End If
        i = i - 1
    Loop
    Korulmetel = Szoveg
End Function
Public Function Atalakit(Adat As String, Optional egyeb As Variant)
    Dim i As Integer, uj As String
    uj = ""
    i = 1
    Do While i <= Len(Adat)
        If Mid(Adat, i, 3) = "$vt" Then
            uj = uj & Mid(Vakterkep.Konyvtar, 1, Len(Vakterkep.Konyvtar) - 1)
            i = i + 3
        End If
        
        If Mid(Adat, i, 1) = "\" And i = 1 Then
            uj = uj & egyeb
            i = i + 1
        End If
        
        If Mid(Adat, i, 3) = "$sz" Then
            uj = uj & egyeb
            i = i + 3
        End If
        
        uj = uj & Mid(Adat, i, 1)
        i = i + 1
    Loop
    Atalakit = uj
End Function
Public Function Konyvtara(Fajlnev As String)
    Dim i As Integer, j As Integer
    j = 0
    For i = 1 To Len(Fajlnev)
        If Mid(Fajlnev, i, 1) = "\" Then j = i
    Next i
    Konyvtara = Mid(Fajlnev, 1, j)
End Function
Public Function CsakANeve(Eleres As String)
    CsakANeve = Mid(Eleres, Len(Konyvtara(Eleres)) + 1, Len(Eleres))
End Function
Public Sub loggol(Szoveg As String)
    log.Text = log.Text & Szoveg & vbCrLf
End Sub
