VERSION 5.00
Begin VB.Form frmPalette 
   AutoRedraw      =   -1  'True
   Caption         =   "Color Paletter"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClip 
      Caption         =   "Copy To Clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CommandButton cmdSaveBMP 
      Caption         =   "Save 8 bit Bitmap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Get From Clipboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox DestPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2160
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame frmLocations 
      Caption         =   "Load Palette From... / Save Bitmap To..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   6495
      Begin VB.DriveListBox MyDrive 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6255
      End
      Begin VB.DirListBox MyDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.FileListBox MyFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3600
         Pattern         =   "*.pal"
         TabIndex        =   3
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.PictureBox SourcePic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Proximity Filter (recomended)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Color Type Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Lineal Regression Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Paleta(0 To 255) As Long
Dim Paletting As Boolean
Dim strImage As String
Dim CMP As String


Private Sub cmdClip_Click()
Clipboard.Clear
Clipboard.SetData DestPic.Image
End Sub

Private Sub Form_Load()
MyDir.Path = "."
MyDrive.Drive = Left(MyDir.Path, 2) & "\"
MyFile.Path = "."
'Debug.Print "'" & NumToBytes(1078, 4) & "'"
End Sub

Private Sub cmdCopy_Click()
If Clipboard.GetFormat(2) Then
    Image1.Picture = Clipboard.GetData
    SourcePic.Picture = Image1.Picture
    SourcePic.Move SourcePic.Left, SourcePic.Top, Image1.Width, Image1.Height
    DestPic.Move DestPic.Left, DestPic.Top, Image1.Width, Image1.Height
Else
    Beep
End If
End Sub

Private Sub cmdSaveBMP_Click()
Dim Px As Long, Py As Long
Dim k As Long
Dim Str As String, Archivo As String
Px = DestPic.ScaleWidth
Py = DestPic.ScaleHeight
'Build BITMAP (I dunno how to use bitmaps in here, so I use the Format code manually)
'HEADER
Str = "BM" & NumToBytes(1078 + Px * Py, 4) & String(4, 0) & NumToBytes(1078, 4)
'BITMAP INFO
Str = Str & NumToBytes(40, 4) & NumToBytes(Px, 4) & NumToBytes(Py, 4)
Str = Str & NumToBytes(1, 2) & NumToBytes(8, 2) & NumToBytes(0, 4) & NumToBytes(Px * Py, 4)
Str = Str & NumToBytes(0, 8) & NumToBytes(256, 4) & NumToBytes(0, 4)
For k = 0 To 255
    Str = Str + NumToBytes(Paleta(k), 3) + Chr(0)
Next k
Str = Str + strImage
Archivo = MyDir.Path
If Not Right(Archivo, 1) = "\" Then Archivo = Archivo + "\"
Open Archivo + InputBox("Write ONLY the name of the bitmap file", "Save BitMap", "Bitmap") + ".bmp" For Binary As #1
Put #1, 1, Str
Close #1
Beep
End Sub

Private Sub cmdProcess_click()
If cmdProcess.Caption = "Process!" Then
    Paletting = True
    cmdProcess.Caption = "Stop"
    DestPic.Visible = False
    cmdSaveBMP.Enabled = False
    cmdCopy.Enabled = False
    DoPalette
Else
    Paletting = False
    DestPic.Visible = True
    cmdSaveBMP.Enabled = True
    cmdCopy.Enabled = True
    cmdProcess.Caption = "Process!"
End If
End Sub

Private Sub MyDir_Change()
MyFile.Path = MyDir.Path
End Sub


Private Sub MyDrive_Change()
MyDir.Path = Left(MyDrive.Drive, 2) + "\"
End Sub

Private Sub MyFile_Click()
Dim k As Byte, l As Byte
CMP = MyFile.Path & "\" & MyFile.List(MyFile.ListIndex)
CargarPAL CMP
For l = 0 To 15
    For k = 0 To 15
        frmPalette.Line (280 + k * 10, 8 + l * 10)-(289 + k * 10, 17 + l * 10), Paleta(k + l * 16), BF
        'Form1.ForeColor = RGB(255, 255, 255) - Paleta(k + l * 16)
        'Form1.PSet (280 + k * 10, 8 + l * 10), Paleta(k + l * 16)
        'Form1.Print (k + l * 16) Mod 10
    Next k
Next l
End Sub


Sub CargarPAL(ByVal Cual As String)
Dim Archivo As Integer, colorear As Byte
Dim trash As String
Dim k As Integer, l As Integer, comp(0 To 3) As Byte
Archivo = FreeFile
trash = "H"
Open Cual For Binary As #Archivo
    colorear = 0
    For k = 25 To LOF(Archivo)
        Get Archivo, k, trash
        comp(colorear) = Asc(trash)
        colorear = colorear + 1
        If colorear > 3 Then
            colorear = 0
            Paleta(Int((k - 25) / 4)) = RGB(comp(0), comp(1), comp(2))
        End If
    Next k
Close #Archivo
End Sub

Function Tipo(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As Byte
If R = G Then
    If G = B Then Tipo = 0
    If G > B Then Tipo = 1
    If G < B Then Tipo = 2
End If
If R > G Then
    If G = B Then Tipo = 3
    If G > B Then Tipo = 4
    If G < B Then
        If B = R Then Tipo = 5
        If B > R Then Tipo = 6
        If B < R Then Tipo = 7
    End If
End If
If R < G Then
    If G = B Then Tipo = 8
    If G > B Then
        If B = R Then Tipo = 9
        If B > R Then Tipo = 10
        If B < R Then Tipo = 11
    End If
    If G < B Then Tipo = 12
End If
End Function

Function NumToBytes(ByVal Number As Long, ByVal Chars As Long) As String
Dim l As Long
Dim c(0 To 7) As String * 1
For l = Chars - 1 To 0 Step -1
    If Int(Number / (256 ^ l)) <= 255 Then
        c(l) = Chr(Int(Number / (256 ^ l)))
        Number = Number - Int(Number / (256 ^ l)) * (256 ^ l)
    End If
Next l
NumToBytes = ""
For l = 0 To Chars - 1
    NumToBytes = NumToBytes + c(l)
Next l
End Function

Sub DoPalette()
'Variables not altered
Dim R As Long, G As Long, B As Long
Dim Ar As Long, Ag As Long, Ab As Long
Dim Px As Integer, Py As Integer
Dim k As Integer, UseCol As Integer

'Variables altered
Dim MinC As Single
Dim Beta0 As Single, Beta1 As Single
Dim Valor As Byte, MaxValor As Byte
For k = chkFilter.LBound To chkFilter.UBound
    MaxValor = MaxValor + chkFilter(k).Value
Next k
If MaxValor = 0 Then MsgBox "No filters checked", vbCritical: Exit Sub
strImage = ""
DestPic.Cls
'------------------------------------------------------------------
Py = SourcePic.ScaleHeight - 1
Px = 0
Do
    If Not Paletting Then Exit Do
'Extraer componentes de color de imagen
        B = Int(SourcePic.Point(Px, Py) / (256 ^ 2))
        G = Int((SourcePic.Point(Px, Py) - (B * (256 ^ 2))) / 256)
        R = Int(SourcePic.Point(Px, Py) - (B * (256 ^ 2)) - G * 256)
        MinC = 1024
        Beta0 = 10000
        Beta1 = 10000
        UseCol = -1
        For k = 0 To 255
'Extraer componentes de color de Paleta(k)
            Valor = 0
            Ab = Int(Paleta(k) / (256 ^ 2))
            Ag = Int((Paleta(k) - (Ab * (256 ^ 2))) / 256)
            Ar = Int(Paleta(k) - (Ab * (256 ^ 2)) - Ag * 256)
            'Proximity Filter
            If Not chkFilter(0).Value = 0 Then If MinC > Abs(R - Ar) * 2 + Abs(G - Ag) + Abs(B - Ab) / 2 Then Valor = Valor + 1
            'ColorType Filter
            If Not chkFilter(1).Value = 0 Then If Tipo(R, G, B) = Tipo(Ar, Ag, Ab) Then Valor = Valor + 1
            'Regression Filter
            If Not chkFilter(2).Value = 0 Then If Abs((B - R) / 2 - (Ab - Ar) / 2) < Beta1 And Abs((5 * R + 2 * G - B) / 6 - (5 * Ar + 2 * Ag - Ab) / 6) < Beta0 Then Valor = Valor + 1
            If Valor >= MaxValor Then
                MinC = Abs(R - Ar) * 2 + Abs(G - Ag) + Abs(B - Ab) / 2
                Beta0 = Abs((5 * R + 2 * G - B) / 6 - (5 * Ar + 2 * Ag - Ab) / 6)
                Beta1 = Abs((B - R) / 2 - (Ab - Ar) / 2)
                UseCol = k
            End If
        Next k
        If UseCol < 0 Then
            Debug.Print "Color no detectado para pixel " & Px & "," & Py
        Else
            strImage = strImage + Chr(UseCol)
            DestPic.PSet (Px, Py), Paleta(UseCol)
        End If
    Px = Px + 1
    If Px = DestPic.ScaleWidth Then
        Py = Py - 1
        If Py < 0 Then Exit Do
        Px = 0
        Caption = 100 - Int((Py / SourcePic.ScaleHeight) * 100) & "%"
    End If
    DoEvents
Loop
Beep
Paletting = False
DestPic.Visible = True
cmdSaveBMP.Enabled = True
cmdCopy.Enabled = True
cmdProcess.Caption = "Process!"
End Sub
