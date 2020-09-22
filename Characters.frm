VERSION 5.00
Begin VB.Form CharEditor 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Benutzerdefinierte Zeichen"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Characters.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   420
      TabIndex        =   14
      Top             =   2340
      Width           =   1035
   End
   Begin VB.Frame frmPreview 
      Caption         =   "Voransicht"
      Height          =   1035
      Left            =   1500
      TabIndex        =   5
      Top             =   840
      Width           =   1095
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   7
         Left            =   660
         Picture         =   "Characters.frx":000C
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   13
         Top             =   540
         Width           =   210
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   6
         Left            =   480
         Picture         =   "Characters.frx":0092
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   12
         Top             =   540
         Width           =   210
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   5
         Left            =   300
         Picture         =   "Characters.frx":0118
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   11
         Top             =   540
         Width           =   210
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   4
         Left            =   120
         Picture         =   "Characters.frx":019E
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   10
         Top             =   540
         Width           =   210
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   3
         Left            =   660
         Picture         =   "Characters.frx":0224
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   9
         Top             =   240
         Width           =   225
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   2
         Left            =   480
         Picture         =   "Characters.frx":02AA
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   8
         Top             =   240
         Width           =   210
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   1
         Left            =   300
         Picture         =   "Characters.frx":0330
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   7
         Top             =   240
         Width           =   210
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         Height          =   300
         Index           =   0
         Left            =   120
         Picture         =   "Characters.frx":03B6
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   6
         Top             =   240
         Width           =   210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zeichen-Nr."
      Height          =   735
      Left            =   1500
      TabIndex        =   3
      Top             =   60
      Width           =   1095
      Begin VB.ComboBox cboCurSym 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Pixel "
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   4140
      Width           =   735
      Begin VB.Image imgPixelOff 
         Height          =   180
         Left            =   360
         Picture         =   "Characters.frx":043C
         Top             =   240
         Width           =   180
      End
      Begin VB.Image imgPixelOn 
         Height          =   180
         Left            =   120
         Picture         =   "Characters.frx":04B8
         Top             =   240
         Width           =   180
      End
   End
   Begin VB.Frame frmSymbol 
      Caption         =   "Zeichen bearb."
      Height          =   2115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1335
      Begin VB.PictureBox picSymbol 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1850
         Left            =   90
         ScaleHeight     =   123
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   2
         Top             =   240
         Width           =   1190
      End
   End
End
Attribute VB_Name = "CharEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim NewX As Long
Dim NewY As Long
Dim LstX As Long
Dim LstY As Long
Private Type XKoord
  x(0 To 4) As Long
End Type

Private Type YKoord
  y(0 To 7) As XKoord
End Type

Dim UserChar(0 To 7) As YKoord

Private Sub cboCurSym_Click()
For y = 0 To 7
  For x = 0 To 4
    If UserChar(cboCurSym.Text).y(y).x(x) = 0 Then
      picSymbol.PaintPicture imgPixelOff.Picture, x * 15 + 2, y * 15 + 2
    Else
      picSymbol.PaintPicture imgPixelOn.Picture, x * 15 + 2, y * 15 + 2
    End If
  Next x
Next y
End Sub

Private Function LoadChars()
On Error GoTo Fehler:
Dim Temp As String
For Z = 0 To 7
  picPreview(Z).Cls
  For y = 0 To 7
    Temp = INIlesen("UserChars" & Z, "Line" & y)
    For x = 0 To 4
      UserChar(Z).y(y).x(x) = Mid(Temp, x + 1, 1)
      If UserChar(Z).y(y).x(x) = 1 Then
        picPreview(Z).PSet (x * 2 + 2, y * 2 + 2)
        picPreview(Z).PSet (x * 2 + 3, y * 2 + 2)
        picPreview(Z).PSet (x * 2 + 2, y * 2 + 3)
        picPreview(Z).PSet (x * 2 + 3, y * 2 + 3)
      End If
    Next
  Next y
Next Z
cboCurSym.Text = 0
Exit Function
Fehler:
MsgBox "Optionen.ini ist beschädigt oder nicht vorhanden." & vbNewLine & _
       "Der Vorgang wird abgebrochen.", vbOKOnly + vbExclamation
End Function

Private Function SaveChars()
Dim Temp As String
For Z = 0 To 7
  For y = 0 To 7
    Temp = vbNullString
    For x = 0 To 4
      Temp = Temp & UserChar(Z).y(y).x(x)
    Next
    INISchreiben "UserChars" & Z, "Line" & y, Temp
  Next y
Next Z
End Function

Private Sub cmdAbort_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
SaveChars
Unload Me
End Sub

Private Sub Form_Load()
For y = 0 To 7
  For x = 0 To 4
    picSymbol.PaintPicture imgPixelOff.Picture, x * 15 + 2, y * 15 + 2
  Next x
  cboCurSym.AddItem y
Next y
LstX = -1
LstY = -1
cboCurSym.Text = 0
LoadChars
End Sub

Private Sub picPreview_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
  cboCurSym.Text = Index
End If
End Sub

Private Sub picSymbol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
  NewX = Fix((x) / 15)
  NewY = Fix((y) / 15)
  If NewX > 4 Or NewX < 0 Or NewY > 7 Or NewY < 0 Then Exit Sub
  
  If NewX <> LstX Or NewY <> LstY Then
    If UserChar(cboCurSym.Text).y(NewY).x(NewX) = 0 Then
      picSymbol.PaintPicture imgPixelOn.Picture, NewX * 15 + 2, NewY * 15 + 2
      UserChar(cboCurSym.Text).y(NewY).x(NewX) = 1
    Else
      picSymbol.PaintPicture imgPixelOff.Picture, NewX * 15 + 2, NewY * 15 + 2
      UserChar(cboCurSym.Text).y(NewY).x(NewX) = 0
    End If
  End If
  LstX = NewX
  LstY = NewY
End If
End Sub

Private Sub picSymbol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
  NewX = Fix((x) / 15)
  NewY = Fix((y) / 15)
  If NewX > 4 Or NewX < 0 Or NewY > 7 Or NewY < 0 Then Exit Sub
  
  If NewX <> LstX Or NewY <> LstY Then
    If UserChar(cboCurSym.Text).y(NewY).x(NewX) = 0 Then
      picSymbol.PaintPicture imgPixelOn.Picture, NewX * 15 + 2, NewY * 15 + 2
      UserChar(cboCurSym.Text).y(NewY).x(NewX) = 1
    Else
      picSymbol.PaintPicture imgPixelOff.Picture, NewX * 15 + 2, NewY * 15 + 2
      UserChar(cboCurSym.Text).y(NewY).x(NewX) = 0
    End If
  End If
  LstX = NewX
  LstY = NewY
End If
End Sub

Private Sub picSymbol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
  picPreview(cboCurSym.Text).Cls
  For y = 0 To 7
    For x = 0 To 4
      If UserChar(cboCurSym.Text).y(y).x(x) = 1 Then
        picPreview(cboCurSym.Text).PSet (x * 2 + 2, y * 2 + 2)
        picPreview(cboCurSym.Text).PSet (x * 2 + 3, y * 2 + 2)
        picPreview(cboCurSym.Text).PSet (x * 2 + 2, y * 2 + 3)
        picPreview(cboCurSym.Text).PSet (x * 2 + 3, y * 2 + 3)
      End If
    Next x
  Next y
  LstX = -1
  LstY = -1
End If
End Sub
