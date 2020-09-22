VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "LC-Display Testprogramm"
   ClientHeight    =   5745
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frmOptions 
      Caption         =   "Anzeige:"
      Height          =   2355
      Left            =   3780
      TabIndex        =   37
      Top             =   1560
      Width           =   3555
      Begin VB.CheckBox chkClearbeforeOut 
         Caption         =   "Displayinhalt vor der Ausgabe löschen"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   1440
         Value           =   1  'Aktiviert
         Width           =   3195
      End
      Begin VB.CheckBox chkMoveDisplay 
         Caption         =   "Displayinhalt statt Cursor verschieben"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox chkIncreCursor 
         Caption         =   "Cursor inkrementieren"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   960
         Value           =   1  'Aktiviert
         Width           =   2175
      End
      Begin VB.HScrollBar scrBacklight 
         Height          =   255
         Left            =   180
         Max             =   255
         TabIndex        =   45
         Top             =   1980
         Width           =   2175
      End
      Begin VB.CheckBox chkDisplayOn 
         Caption         =   "Display an"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   1515
      End
      Begin VB.CheckBox chkShowCursor 
         Caption         =   "Cursor anzeigen"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1635
      End
      Begin VB.CheckBox chkBlinkCursor 
         Caption         =   "blinkender Cursor"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblBacklight 
         Caption         =   "0"
         Height          =   195
         Left            =   2460
         TabIndex        =   47
         Top             =   2010
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Hintergrundgeleuchtung: (Hardware erf.)"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1740
         Width           =   3315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Erweitert:"
      Height          =   1695
      Left            =   60
      TabIndex        =   18
      Top             =   3960
      Width           =   7275
      Begin VB.CommandButton cmdClearDisplay 
         Caption         =   "Displayinhalt löschen"
         Height          =   375
         Left            =   1740
         TabIndex        =   48
         Top             =   1140
         Width           =   1635
      End
      Begin VB.CommandButton cmdCursorRight 
         Caption         =   ">"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   720
         Width           =   195
      End
      Begin VB.CommandButton cmdManChar 
         Caption         =   "übertragen"
         Height          =   375
         Left            =   6180
         TabIndex        =   36
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtManChar 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   4260
         MaxLength       =   3
         TabIndex        =   34
         Text            =   "000"
         Top             =   1260
         Width           =   400
      End
      Begin VB.CommandButton cmdManText 
         Caption         =   "übertragen"
         Height          =   375
         Left            =   6180
         TabIndex        =   32
         Top             =   550
         Width           =   975
      End
      Begin VB.CommandButton cmdCursorHome 
         Caption         =   "Cursor auf Home"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1140
         Width           =   1395
      End
      Begin VB.CommandButton cmdSetCursor 
         Caption         =   "Cursor setzen"
         Height          =   375
         Left            =   1980
         TabIndex        =   26
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox txtRow 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "01"
         Top             =   720
         Width           =   315
      End
      Begin VB.VScrollBar scrRow 
         Height          =   285
         Left            =   540
         Max             =   1
         Min             =   4
         TabIndex        =   24
         Top             =   720
         Value           =   1
         Width           =   195
      End
      Begin VB.TextBox txtColumn 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   900
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "01"
         Top             =   720
         Width           =   315
      End
      Begin VB.VScrollBar scrColumn 
         Height          =   285
         Left            =   1200
         Max             =   1
         Min             =   27
         TabIndex        =   22
         Top             =   720
         Value           =   1
         Width           =   195
      End
      Begin VB.CommandButton cmdCursorLeft 
         Caption         =   "<"
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
         Left            =   1500
         TabIndex        =   21
         Top             =   720
         Width           =   195
      End
      Begin VB.TextBox txtManText 
         Height          =   285
         Left            =   3780
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Code:"
         Height          =   195
         Left            =   3780
         TabIndex        =   35
         Top             =   1290
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Zeichen ausgeben:"
         Height          =   195
         Left            =   3720
         TabIndex        =   33
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Text ungeparst ausgeben:"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Spalte:"
         Height          =   195
         Left            =   840
         TabIndex        =   29
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblRow 
         Caption         =   "Reihe:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "Cursor:"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   240
         Width           =   3315
      End
   End
   Begin VB.Frame frmDisplay 
      Caption         =   "Display:"
      Height          =   1635
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   3615
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'Kein
         Height          =   435
         Left            =   1140
         TabIndex        =   41
         Top             =   1080
         Width           =   1395
         Begin VB.OptionButton optIntf 
            Caption         =   "4 Bit"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optIntf 
            Caption         =   "8 Bit"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   42
            Top             =   240
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.ComboBox cboPort 
         Height          =   315
         ItemData        =   "Form1.frx":000C
         Left            =   1080
         List            =   "Form1.frx":0019
         TabIndex        =   12
         Top             =   660
         Width           =   1155
      End
      Begin VB.ComboBox cboDisplaySize 
         Height          =   315
         ItemData        =   "Form1.frx":002C
         Left            =   1080
         List            =   "Form1.frx":002E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblIntf 
         Caption         =   "Interface:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label lblPortInfo 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2340
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Port:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Größe/Typ:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Textausgabe:"
      Height          =   2175
      Left            =   60
      TabIndex        =   2
      Top             =   1740
      Width           =   3615
      Begin VB.CommandButton cmdShowVaria 
         Caption         =   "Variablen anzeigen"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubmitText 
         Caption         =   "Text übertragen"
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame frmCommands 
      Caption         =   "Befehle:"
      Height          =   1455
      Left            =   3780
      TabIndex        =   0
      Top             =   60
      Width           =   3555
      Begin VB.CommandButton cmdTestSpeed 
         Caption         =   "Übtr.-Geschw. testen"
         Height          =   375
         Left            =   1560
         TabIndex        =   51
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdEditChars 
         Caption         =   "bearbeiten"
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton cmdSubmitChars 
         Caption         =   "übertragen"
         Height          =   375
         Left            =   1260
         TabIndex        =   7
         Top             =   960
         Width           =   1035
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "Initialisieren"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Benutzerdefinierte Zeichen:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   3315
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDisplaySize_Click()
Lcd.Size = cboDisplaySize.ItemData(cboDisplaySize.ListIndex)
If cboDisplaySize.ItemData(cboDisplaySize.ListIndex) < 6 Then
  txtLine(2).Enabled = False
  txtLine(3).Enabled = False
Else
  txtLine(2).Enabled = True
  txtLine(3).Enabled = True
End If
cmdInit.FontBold = True
End Sub

Private Sub cboPort_Click()
Lcd.Port = Val(cboPort.Text)
INISchreiben "Display", "Port", Val(cboPort.Text)
cmdInit.FontBold = True
End Sub

Private Sub chkBlinkCursor_Click()
If Lcd.Port = "0" Then
  MsgBox "Der Port wurde nicht bestimmt.", vbOKOnly + vbExclamation
Else
  Lcd.Displayset IIf(chkDisplayOn.value = 1, True, False), _
                 IIf(chkShowCursor.value = 1, True, False), _
                 IIf(chkBlinkCursor.value = 1, True, False)
End If
End Sub

Private Sub chkDisplayOn_Click()
If Lcd.Port = "0" Then
  MsgBox "Der Port wurde nicht bestimmt.", vbOKOnly + vbExclamation
Else
  Lcd.Displayset IIf(chkDisplayOn.value = 1, True, False), _
                 IIf(chkShowCursor.value = 1, True, False), _
                 IIf(chkBlinkCursor.value = 1, True, False)
End If
End Sub

Private Sub chkIncreCursor_Click()
Entrymodeset IIf(chkIncreCursor.value = 1, True, False), _
             IIf(chkMoveDisplay.value = 1, True, False)
End Sub

Private Sub chkMoveDisplay_Click()
Entrymodeset IIf(chkIncreCursor.value = 1, True, False), _
             IIf(chkMoveDisplay.value = 1, True, False)
End Sub

Private Sub chkShowCursor_Click()
If Lcd.Port = "0" Then
  MsgBox "Der Port wurde nicht bestimmt.", vbOKOnly + vbExclamation
Else
  Lcd.Displayset IIf(chkDisplayOn.value = 1, True, False), _
                 IIf(chkShowCursor.value = 1, True, False), _
                 IIf(chkBlinkCursor.value = 1, True, False)
End If
End Sub

Private Sub cmdClearDisplay_Click()
Lcd.ClearDisplay
End Sub

Private Sub cmdCursorHome_Click()
Lcd.SetCursorHome
End Sub

Private Sub cmdCursorLeft_Click()
Lcd.MoveCursor dLeft, False
End Sub

Private Sub cmdCursorRight_Click()
Lcd.MoveCursor dRight, False
End Sub

Private Sub cmdEditChars_Click()
CharEditor.Show
End Sub

Private Sub cmdInit_Click()
If Lcd.Port = "0" Then
  MsgBox "Der Port wurde nicht bestimmt.", vbOKOnly + vbExclamation
Else
  Lcd.Init False, False, cBoth
  Lcd.ClearDisplay
End If
cmdInit.FontBold = False
End Sub

Private Sub cmdManChar_Click()
Lcd.WriteChar txtManChar.Text
End Sub

Private Sub cmdManText_Click()
Lcd.WriteText txtManText.Text
End Sub

Private Sub cmdSetCursor_Click()
Lcd.SetPos Val(txtRow.Text), Val(txtColumn.Text)
End Sub

Private Sub cmdSubmitChars_Click()
Dim TmpUserChar As DspOwnSymbol
If Lcd.Port = "0" Then
  MsgBox "Der Port wurde nicht bestimmt.", vbOKOnly + vbExclamation
Else
  For Z = 0 To 7
    For x = 0 To 7
      TmpUserChar.sLine(x) = INIlesen("UserChars" & Z, "Line" & x)
      Lcd.DefineChar Z, TmpUserChar
    Next
  Next
End If
End Sub

Private Sub cmdSubmitText_Click()
If Lcd.Port = "0" Then
  MsgBox "Der Port wurde nicht bestimmt.", vbOKOnly + vbExclamation
Else
  If chkClearbeforeOut.value = 1 Then Lcd.ClearDisplay
  For x = 0 To 3
    If cboDisplaySize.ItemData(cboDisplaySize.ListIndex) < 5 And x >= 2 Then Exit For
      Lcd.SetPos x + 1, 1
      Lcd.WriteText Lcd.ParseText(ParseUserChars(txtLine(x).Text))
  Next
End If
End Sub

Private Function ParseUserChars(ByVal Text As String) As String
For x = 255 To 0 Step -1
  Text = Replace(Text, "$" & x, Chr(x))
Next
Text = Replace(Text, "$Time", Time)
Text = Replace(Text, "$Date", Date)
Text = Replace(Text, "$ShortDayname", WeekdayName(Weekday(Now, vbUseSystemDayOfWeek), True, vbUseSystemDayOfWeek))
Text = Replace(Text, "$Dayname", WeekdayName(Weekday(Now, vbUseSystemDayOfWeek), False, vbUseSystemDayOfWeek))
Text = Replace(Text, "$Day", Day(Now))
Text = Replace(Text, "$ShortMonthname", MonthName(Month(Now), True))
Text = Replace(Text, "$Monthname", MonthName(Month(Now), False))
Text = Replace(Text, "$Year", Year(Now))
Text = Replace(Text, "$Hour", Hour(Now))
Text = Replace(Text, "$Minute", Minute(Now))
Text = Replace(Text, "$Second", Second(Now))
Text = Replace(Text, "$Backlight", Lcd.Backlight)
Text = Replace(Text, "$Port", Lcd.Port)
Text = Replace(Text, "$LCDtype", cboDisplaySize.Text)
Text = Replace(Text, "$+ ", "")
'Text = Replace(Text, "$", "")
ParseUserChars = Text
End Function
Private Sub cmdShowVaria_Click()
MsgBox "Variablen für die Texteingabe:" & vbNewLine & _
       "$0 - $255: Gibt den Zeichen mit angegebenen ASCII-Code aus" & vbNewLine & _
       "$Time: gibt die aktuelle Uhrzeit aus*" & vbNewLine & _
       "$Date: gibt das aktuelle Datum aus*" & vbNewLine & _
       "$ShortDayname: Gibt die Kurzform des aktuellen Tages aus (Mo, Di, Mi,...)*" & vbNewLine & _
       "$Dayname: Gibt den Namen des aktuellen Tages aus (Montag, Dienstag, Mittwoch)*" & vbNewLine & _
       "$Day: Gibt den aktuellen Tag im Monat aus" & vbNewLine & _
       "$ShortMonthname: Gibt kurzen Namen des aktuellen Monats aus (Jan, Feb, Mär,...)*" & vbNewLine & _
       "$Monthname: Gibt den vollstängigen Namen des aktuellen Monats aus (Januar, Februar,...)*" & vbNewLine & _
       "$Year: Gibt das aktuelle Jahr aus*" & vbNewLine & _
       "$Hour: Gibt die Stunde der Uhrzeit aus*" & vbNewLine & _
       "$Minute: Gibt die Minute der Uhrzeit aus*" & vbNewLine & _
       "$Second: Gibt die Sekunden der Uhrzeit aus*" & vbNewLine & _
       "$Backlight: Gibt den Helligkeitswert der Hintergrundbeleuchtung von 0-255 aus" & vbNewLine & _
       "$Port: Gibt den verwendeten Port aus" & vbNewLine & _
       "$LCDtype: Gibt den Typ des verwendeten Displays aus (Combobox oben)" & vbNewLine & _
       "$+: Entfernt das nachfolgende Leerzeichen" & vbNewLine & _
       "*) Format wie in Windows vorgegeben", vbInformation + vbOKOnly ' & vbNewLine & _
       "$" & vbNewLine & _
       "$" & vbNewLine
End Sub


Private Sub Form_Load()
Dim CurrDisplay As String
For x = 0 To 3
  txtLine(x).Text = INIlesen("Standard-Text", "Line" & x)
Next
For x = 1 To Val(INIlesen("Display", "Sizes"))
  CurrDisplay = INIlesen("Display", "Size" & x)
  cboDisplaySize.AddItem Mid(CurrDisplay, InStr(1, CurrDisplay, ",") + 1)
  cboDisplaySize.ItemData(x - 1) = Left(CurrDisplay, InStr(1, CurrDisplay, ","))
Next
Temp = INIlesen("Display", INIlesen("Display", "LastSize"))
cboDisplaySize.Text = Mid(Temp, InStr(1, Temp, ",") + 1)
Lcd.Size = cboDisplaySize.ItemData(cboDisplaySize.ListIndex)
cboPort.Text = INIlesen("Display", "Port")
Lcd.Port = Val(cboPort.Text)
Lcd.Mode = m8Bit
End Sub

Private Sub lblPortInfo_Click()
MsgBox "LPT1: 888" & vbCrLf & "LPT2: 632" & vbCrLf & "LPT3: 956", vbInformation + vbOKOnly
End Sub

Private Sub mnuFileEnd_Click()
End
End Sub

Private Sub cmdTestSpeed_Click()
MsgBox "Übertragungszeit eines kompletten Speicherinhalts: " & TestLcdSpeed & "ms", vbInformation + vbOKOnly
End Sub

Private Sub optIntf_Click(Index As Integer)
If Index = 0 Then
  Lcd.Mode = m4Bit
Else
  Lcd.Mode = m8Bit
End If
cmdInit.FontBold = True
End Sub

Private Sub scrBacklight_Change()
lblBacklight.Caption = scrBacklight.value
Lcd.Backlight = scrBacklight.value
End Sub

Private Sub scrColumn_Change()
txtColumn.Text = Format(scrColumn.value, "00")
End Sub

Private Sub txtColumn_Change()
On Error Resume Next
scrColumn.value = txtColumn.Text
End Sub

Private Sub txtColumn_Click()
txtColumn.SelStart = 0
txtColumn.SelLength = Len(txtColumn.Text)
End Sub

Private Sub txtColumn_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 75: 'Nichts tun
  Case 8 'Nichts tun
  Case Else: KeyAscii = 0
End Select
End Sub

Private Sub scrRow_Change()
txtRow.Text = Format(scrRow.value, "00")
End Sub

Private Sub txtManChar_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 75: 'Nichts tun
  Case 8 'Nichts tun
  Case 13: Call cmdManChar_Click: KeyAscii = 0
  Case Else: KeyAscii = 0
End Select
End Sub

Private Sub txtManText_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 13: Call cmdManText_Click: KeyAscii = 0
End Select
End Sub

Private Sub txtRow_Change()
On Error Resume Next
scrRow.value = txtRow.Text
End Sub

Private Sub txtRow_Click()
txtRow.SelStart = 0
txtRow.SelLength = Len(txtColumn.Text)
End Sub

Private Sub txtRow_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 75: 'Nichts tun
  Case 8 'Nichts tun
  Case Else: KeyAscii = 0
End Select
End Sub
'--
