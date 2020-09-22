Attribute VB_Name = "Lcd"
Option Explicit

'LCD-Ansteuerung von Christof Rueß, www.hobby-elektronik.de.vu
'Die Verwendung dieses Moduls ist kostenfrei, solange diese Zeilen erhalten bleiben.
'Fragen, Kritik und Anregungen an: hobbyelektronik@gmx.net
'--
'Anschlussbelegung für HD4478(kompatible)-Displays
'PC D-SUB-Stecker - Display
'   1 (Datatrobe) - Enable 1
'          2 (D0) - DB0*
'          3 (D1) - DB1*
'          4 (D2) - DB2*
'          5 (D3) - DB3*
'          6 (D4) - DB4
'          7 (D5) - DB5
'          8 (D6) - DB6
'          9 (D7) - DB7
'   14 (Autofeed) - R/W - Read/Write (bei 4x27 und 4x40 für Hintergrundbeleuchtung)
'       16 (Init) - RS - Register Select
'     17 (Select) - Enable 2 bzw. Enable für Hintergrundbeleuchtungssteuerung
'   18 - 25 (GND) - GND
' *) Für den 4-Bit-Betrieb nicht nötig
'Zusätzlich an das Display:
'Masse an Pin 1 (Vss) & evtl. an 16 (LED K)
' +5V  an Pin 2 (Vdd) & evtl. (mit Vorwiderstand! - etwa 10-50 Ohm) an 15 (A/Vee)
'Kontrastspannung an Pin 3 (10kOhm-Poti - Schleifer an Pin 3 die anderen zwei jeweils an +5V und GND)

'Deklaration für die Portansteuerung (auch für Win2k & WinXP)
Public Declare Sub PortOut32 Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal value As Integer)
Public Declare Function PortIn32 Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
 
'Da man mit mit der "normalen" Anschlussbelegung des Displays das Busy-Flag nicht auslesen kann, werden diese Routinen benötigt.
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'Falls QueryPerformance* nicht verfügbar ist, wird das normale Sleep verwendet.
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)

'Damit das Auslesen der Bilder für die selbstdefinierten Zeichen schneller geht, bediene ich mich dieser API.
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'Für die Einstellung, ob das LCD im 8- oder 4-Bit-Modus betrieben werden soll
'mDefault ist nur für Modulinterne Verarbeitungen bestimmt und sollte von außen nicht gesetzt werden
Public Enum DspMode
  mDefault = 0
  m4Bit = 1
  m8Bit = 2
End Enum

'Wichtig, wenn ein Display mit mehreren Controllern verwendet wird
Public Enum DspContr
  cUp = 0
  cDown = 1
  cBoth = 2
End Enum

'Die Displaygrößen für die Adressfindung definieren.
Public Enum DspSize
  s1x16 = 0
  s2x08 = 1
  s2x16 = 2
  s2x20 = 3
  s2x24 = 4
  s2x40 = 5
  s4x16 = 6
  s4x20 = 7
  s4x27 = 8
  s4x40 = 9
End Enum

'Wird für MoveCursor verwendet um die Richtung des Curosrs zu bestimmen
Public Enum DspDir
  dLeft = 0
  dRight = 1
End Enum
 
'Hier werden die Zeilen für die eigenen Zeichen definiert.
'Die Function zum "richtigen" definieren der Zeichen folgt später...
Public Type DspOwnSymbol
  sLine(0 To 7) As String
End Type

'Variable für die Displayeigenschaften
Private LcdPort As Long 'Port des LCDs (intern)
Private LcdSize As DspSize 'Größe des LCDs (intern - siehe auch Enum "DspSize")
Private LcdCurrCntrl As DspContr 'Aktueller Controller des Displays (intern - siehe auch Enum "DspContr")
Private LcdBright As Long 'Helligkeit der Hintergrundbeleuchtung (intern - Hardware erforderlich)
Private LcdMode As DspMode '"Bandbreite" der Ansteuerung des Displays (intern - siehe auch Enum "DspMode")
Public SwapRSRW As Boolean 'RS und R/W getauscht? (extern!)

'Für WaitforLCD notwendig
Private PerfFrequ As Currency
Private PerfCntA As Currency
Private PerfCntB As Currency
Private PerfCntC As Currency
Private PerfCntAvl As Long

'Hiermit wird der Port, an dem das LCD angschlossen ist festgelegt
Property Let Port(NewPort As Long)
LcdPort = NewPort
End Property

'Damit der Port auch gelesen werden kann, ist dies hier notwendig:
Property Get Port() As Long
Port = LcdPort
End Property

'Setzen der Displaygröße
Property Let Size(NewSize As DspSize)
LcdSize = NewSize
End Property

'Lesen der Displaygröße
Property Get Size() As DspSize
Size = LcdSize
End Property

'Setzen des Displaymodi
Property Let Mode(NewMode As DspMode)
LcdMode = NewMode
End Property

'Lesen des Displaymodi
Property Get Mode() As DspMode
Mode = LcdMode
End Property

'Setzen der Helligkeit der Hintergrundbeleuchtung
'Setzt Steuerung voraus: http://www.hobby-elektronik.de.vu/EloPC/LCDLight
Property Let Backlight(Bright As Long)
On Error GoTo Fehler
If LcdSize < 8 Then
  PortOut32 LcdPort, Bright
  PortOut32 LcdPort + 2, 10 ' ConvertBintoDez("00001010")
  PortOut32 LcdPort + 2, 2 ' ConvertBintoDez("00000010")
  LcdBright = Bright
Else
  PortOut32 LcdPort, Bright
  PortOut32 LcdPort + 2, 0 ' ConvertBintoDez("00000000")
  PortOut32 LcdPort + 2, 2 ' ConvertBintoDez("00000010")
  LcdBright = Bright
End If
Fehler:
End Property

'Mit dieser Funktion kann die Variable "Backlight" ausgelesen werden.
'Falls die Variable auch nach dem Schreiben noch 0 ist, konnte die Helligkeit nicht gesetzt werden
Property Get Backlight() As Long
  Backlight = LcdBright
End Property

'Neue Variante für die Pause, die die CPU etwas entlastet und zugleich noch einiges schneller ist...
Function WaitforLCD()
PerfCntAvl = QueryPerformanceFrequency(PerfFrequ)
'Auslesen der Performance-Frequenz für die Anwendung
'"Nur" unter Win9x, ME, NT, 2k und XP möglich

If PerfCntAvl = 0 Then 'Wenn der hochauflösende Timer nicht verfügbar ist...
  Sleep 1 '...soll die Pause nach dem "alten" Verfahren gemacht werden
Else
  QueryPerformanceCounter PerfCntA 'den Startwert für den Counter in PerfCntA schreiben
  Do 'so lange die Schleife ausführen bis...
    QueryPerformanceCounter PerfCntB
    PerfCntC = CDbl((PerfCntB - PerfCntA) / PerfFrequ)
  Loop Until PerfCntC > 0.0001 '...die Zeit von 0.8ms vorbei ist
End If
End Function
 
'Diese Funktion zwei Funktionen stammen (leider) nicht von mir.
'Sie sind aber sehr nützlich...
'Hier werden Binäre Zahlen zu Dezimalen umgewandelt...
Function ConvertBintoDez(ByVal Bin As String) As Long
  Dim x&, y&
    Bin = Bin & String$(8 - Len(Bin), "0")
    For x = 1 To Len(Bin)
      If Mid$(Bin, x, 1) = "1" Then
        y = y + 2 ^ (8 - x)
      End If
    Next x
    ConvertBintoDez = y
End Function

'...und hier Dezimale zu Binären.
Function ConvertDeztoBin(ByVal Dez As Long) As String
  Dim x%
   If Dez >= 2 ^ 32 Then
     Call MsgBox("Zahl ist größer als 32 Bit")
     Exit Function
   End If
   Do
     If (Dez And 2 ^ x) Then
       ConvertDeztoBin = "1" & ConvertDeztoBin
     Else
       ConvertDeztoBin = "0" & ConvertDeztoBin
     End If
     x = x + 1
   Loop Until 2 ^ x > Dez
   ConvertDeztoBin = Format(ConvertDeztoBin, "00000000")
End Function

'Ausgabe an das LCD zum experimentieren. Diese Funktion kann weggelassen werden.
'In den Zeilen ist auch gut zu sehen, dass das R/W- und Enable-Signal wegen einer Eigenschaft des Druckerports invertiert werden müssen.
Function TestingOutLcd(ByVal LcdDB7 As Boolean, ByVal LcdDB6 As Boolean, _
                              ByVal LcdDB5 As Boolean, ByVal LcdDB4 As Boolean, ByVal LcdDB3 As Boolean, _
                              ByVal LcdDB2 As Boolean, ByVal LcdDB1 As Boolean, ByVal LcdDB0 As Boolean, _
                              ByVal LcdRW As Boolean, ByVal LcdRS As Boolean)
PortOut32 LcdPort, ConvertBintoDez(IIf(LcdDB7 = True, 1, 0) & IIf(LcdDB6 = True, 1, 0) & IIf(LcdDB5 = True, 1, 0) & _
                                   IIf(LcdDB4 = True, 1, 0) & IIf(LcdDB3 = True, 1, 0) & IIf(LcdDB2 = True, 1, 0) & _
                                   IIf(LcdDB1 = True, 1, 0) & IIf(LcdDB0 = True, 1, 0))
  'So, hier werden die Daten an die DB-Kanäle ausgegeben. Da man Booleans (komischerweise) nicht direkt als 1/0-Werte ausgeben kann,
  'greife ich hier auf die IIF-Befehle zurück. Dann noch schnell nach Dezimal konvertiert und ab ans Display
PortOut32 LcdPort + 2, ConvertBintoDez("00000" & IIf(LcdRS = True, 1, 0) & IIf(LcdRW = True, 0, 1) & 0)
WaitforLCD
PortOut32 LcdPort + 2, ConvertBintoDez("00000" & IIf(LcdRS = True, 1, 0) & IIf(LcdRW = True, 0, 1) & 1)
WaitforLCD
End Function

 
'Ausgabe an das LCD, die ideal ist. Die DB-Kanäle werden über LcdData ausgegeben, Der Rest sollte bekannt sein ;)
Function OutLcd(ByVal LcdData As Byte, ByVal LcdRW As Boolean, ByVal LcdRS As Boolean, _
                Optional ByVal Controller As DspContr = cUp, Optional ByVal Mode As DspMode = mDefault)

'Festlegen der temporären Variablen
Dim OutE2 As Long
Dim OutRS As Long
Dim OutRW As Long
Dim OutE1 As Long
Dim CurrMode As DspMode

'Wenn der Parallel-Port nicht gesetzt wurde, wird nichts ausgegeben
If LcdPort = 0 Then Exit Function

'Wenn der Mode "Standard" ist, wird er aus den festgelegten Einstellungen gelesen.
If Mode = mDefault Then
  CurrMode = LcdMode
Else
  CurrMode = Mode
End If

'Die Fälle für Controllers bestimmen:
Select Case Controller
  Case 0: OutE1 = 1: OutE2 = 0 'Oberer Controller: E1 = 1; E2 = 0
  Case 1: OutE1 = 0: OutE2 = 8 'Unterer Controller: E1 = 0; E2 = 8
  Case 2: OutE1 = 1: OutE2 = 8 'Beide Controller: E1 = 1; E2 = 8
End Select

'Da bei manchen Belegungen (Beispiel: Pollin) RS und R/W vertauscht sind, wird dies nun auch berücksichtigt:
If SwapRSRW = False Then
  'RS festlegen
  If LcdRS = True Then OutRS = 4 Else OutRS = 0

  'R/W festlegen (eigentlich nicht nötig, da nicht möglich)
  If LcdRW = True Then OutRW = 0 Else OutRW = 2
Else
  If LcdRS = True Then OutRS = 0 Else OutRS = 2
  If LcdRW = True Then OutRW = 4 Else OutRW = 0
End If

If CurrMode = m8Bit Then
  PortOut32 LcdPort, LcdData
    'Hier ist die Dateiausgabe sehr simpel. "LcdData" wird schon als Byte erwartet.
    'So ist ein Runtime-Error schon von vorne weg ausgeschlossen...
    
  PortOut32 LcdPort + 2, OutRS + OutRW
    'Hier werden Enable1 & 2 auf low gesetzt um das Display für die Datenaufnahme bereit zu machen
  WaitforLCD 'kurz auf das Display warten
  PortOut32 LcdPort + 2, OutE2 + OutRS + OutRW + OutE1
    'Enable1 bzw. 2 auf high setzen
  WaitforLCD 'noch einmal auf das Display warten

  'Nun könnte man Enable1 & 2 wieder auf low setzen. Ist aber nicht unbedingt erforderlich.
Else
  Dim LcdBinData As String
  LcdBinData = ConvertDeztoBin(LcdData)
  
  PortOut32 LcdPort, ConvertBintoDez(Left(LcdBinData, 4))

  PortOut32 LcdPort + 2, OutRS + OutRW
    'Hier werden Enable1 & 2 auf low gesetzt um das Display für die Datenaufnahme bereit zu machen
  WaitforLCD 'kurz auf das Display warten
  PortOut32 LcdPort + 2, OutE2 + OutRS + OutRW + OutE1
    'Enable1 bzw. 2 auf high setzen
  WaitforLCD 'noch einmal auf das Display warten

  PortOut32 LcdPort, ConvertBintoDez(Right(LcdBinData, 4))

  PortOut32 LcdPort + 2, OutRS + OutRW
    'Hier werden Enable1 & 2 auf low gesetzt um das Display für die Datenaufnahme bereit zu machen
  WaitforLCD 'kurz auf das Display warten
  PortOut32 LcdPort + 2, OutE2 + OutRS + OutRW + OutE1
    'Enable1 bzw. 2 auf high setzen
  WaitforLCD 'noch einmal auf das Display warten
End If
End Function
 
'Da die HD44780-Displays (und kompatible) einen Charset haben,
'der etwas von dem ASCII-Zeichensatz abweicht, habe ich einen
'kleinen Parser geschrieben. Es ist zwar nicht immer passend, aber immerhin...
Function ParseText(ByVal Text As String) As String
Dim x As Long
Dim Parsed As Long
Dim Temp As String
For x = 1 To Len(Text)
    Select Case Asc(Mid(Text, x, 1))
        Case 176: Parsed = 223 ' °
        Case 167: Parsed = 32  ' § (als Leerzeichen, da kein Char vorhanden ist)
        Case 178: Parsed = 32  ' ² (als Leerzeichen, da kein Char vorhanden ist)
        Case 179: Parsed = 32  ' ³ (als Leerzeichen, da kein Char vorhanden ist)
        Case 92:  Parsed = 47  ' \
        Case 180: Parsed = 96  ' ´
        Case 126: Parsed = 243 ' ~
        Case 246: Parsed = 239 ' ö
        Case 228: Parsed = 225 ' ä
        Case 252: Parsed = 245 ' ü
        Case 246: Parsed = 239 ' Ö (als kleines ö)
        Case 228: Parsed = 225 ' Ä (als kleines ä)
        Case 252: Parsed = 245 ' Ü (als kleines ü)
        Case 223: Parsed = 226 ' ß
        Case 181: Parsed = 228 ' µ
        'Case xxx: Parsed = xxx  zum fortführen
        Case Else: Parsed = Asc(Mid(Text, x, 1))
    End Select
Temp = Temp & Chr(Parsed)
Next
ParseText = Temp
End Function

'Hier kann man die DD-RAM-Adresse aus der Reihe und des Zeichens herausfinden.
Function GetDDRamAddress(ByVal LcdRow As Long, ByVal LcdColumn) As Long
Select Case LcdSize
Case 0 '1x16
  LcdCurrCntrl = cUp
  If LcdColumn <= 8 Then
    GetDDRamAddress = -1 + LcdColumn
  Else
    GetDDRamAddress = 63 + LcdColumn
  End If
Case 1, 2, 3, 4, 5 '2x8, 2x16, 2x20, 2x24 und 2x40
  LcdCurrCntrl = cUp
  Select Case LcdRow
    Case 1: GetDDRamAddress = -1 + LcdColumn
    Case 2: GetDDRamAddress = 63 + LcdColumn
  End Select
Case 6 '4x16
  LcdCurrCntrl = cUp
  Select Case LcdRow
    Case 1: GetDDRamAddress = -1 + LcdColumn
    Case 2: GetDDRamAddress = 63 + LcdColumn
    Case 3: GetDDRamAddress = 15 + LcdColumn
    Case 4: GetDDRamAddress = 79 + LcdColumn
  End Select
Case 7 '4x20
  LcdCurrCntrl = cUp
  Select Case LcdRow
    Case 1: GetDDRamAddress = -1 + LcdColumn
    Case 2: GetDDRamAddress = 63 + LcdColumn
    Case 3: GetDDRamAddress = 19 + LcdColumn
    Case 4: GetDDRamAddress = 83 + LcdColumn
  End Select
Case 8, 9 '4x40 und 4x27
  Select Case LcdRow
    Case 1, 3: GetDDRamAddress = -1 + LcdColumn
    Case 2, 4: GetDDRamAddress = 63 + LcdColumn
  End Select
  Select Case LcdRow
    Case 1, 2: LcdCurrCntrl = cUp
    Case 3, 4: LcdCurrCntrl = cDown
  End Select
End Select
End Function

'Hier kann man die Position des Cursors auf dem Display setzen:
Function SetPos(ByVal LcdRow As Long, ByVal LcdColumn As Long)
OutLcd GetDDRamAddress(LcdRow, LcdColumn) + 128, False, False, LcdCurrCntrl
End Function
 
'Das hier ist so ziemlich der einzige Initialisierbefehl:
Function Functionset(ByVal Interface8Bit As Boolean, ByVal Multiline As Boolean, _
                     Optional ByVal Controller As DspContr = cUp)
OutLcd ConvertBintoDez("001" & IIf(Interface8Bit = True, 1, 0) & IIf(Multiline = True, 1, 0) & "000"), False, False, Controller
End Function
 
'Hier gibts noch etwas für das Display einzustellen. LCD an/aus, Cursor anzeigen und/oder Corsor blinken...
Function Displayset(ByVal DisplayOn As Boolean, ByVal ShowCursor As Boolean, ByVal CursorBlink As Boolean, _
                    Optional ByVal Controller As DspContr = cUp)
OutLcd ConvertBintoDez("00001" & IIf(DisplayOn = True, 1, 0) & IIf(ShowCursor = True, 1, 0) & IIf(CursorBlink = True, 1, 0)), False, False, Controller
End Function
 
'Hier werden einfach nur die beiden vorhergehenden Funktionen vereinigt.
'Dies fördert eigentlich nur die Faulheit des Programmierers - aber egal - so wird es sehr simpel.
Function Init(Optional ByVal ShowCursor As Boolean, Optional ByVal CursorBlink As Boolean, _
              Optional ByVal Controller As DspContr = cUp, Optional ByVal Mode As DspMode = mDefault)
              
Dim CurrMode As DspMode

If Mode = mDefault Then
  CurrMode = LcdMode
Else
  CurrMode = Mode
End If

If CurrMode = m8Bit Then
  Functionset True, True, Controller
  Functionset True, True, Controller
  Displayset True, ShowCursor, CursorBlink, Controller
  'das zweite mal ist nur nötig, wenn das LCD vorher im 4-Bit-Modus war.
  Functionset True, True, Controller
  Displayset True, ShowCursor, CursorBlink, Controller
Else
  OutLcd 48, False, False, Controller, m8Bit '48 =00110000
  'Sleep 15 'die Pause kann notwendig sein, muss aber nicht
  OutLcd 48, False, False, Controller, m8Bit
  OutLcd 48, False, False, Controller, m8Bit
  'Nun kann das Display endlich auf 4 Bit eingestellt werden.
  OutLcd 32, False, False, Controller, m8Bit '32 = 00100000
  'Das Display nocheinmal auf 4 Bit einstellen.
  Functionset False, True, Controller
  'Display anschalten und je nach Bedarf den Cursor aktivieren
  Displayset True, ShowCursor, CursorBlink, Controller
End If

ClearDisplay Controller 'Display noch löschen
SetCursorHome Controller ' und den Cursor auf 1, 1 setzen
End Function
 
'Mit diesem Befehl kann man das Display kurz und schmerzlos löschen.
'Hier wird nur der DD-RAM gelöscht - nicht der CG-RAM.
Function ClearDisplay(Optional ByVal Controller As DspContr = cUp)
OutLcd 1, False, False, Controller
End Function
 
'Diese Funktion setzt den Cursor auf die DD-RAM-Adresse 0 zurück.
'Das kann man auch mit einem größeren aufwand mit der LcdSetPos-Funktion machen.
Function SetCursorHome(Optional ByVal Controller As DspContr = cUp)
OutLcd 2, False, False, Controller
End Function

'Hier ist der eigentliche Textausgabe. Hierzu braucht man nur den Port und den gewünschten Text...
'...dieser wird dann Buchstabe für Buchstabe als ASCII-Code ausgegeben.
Function WriteText(ByVal LcdText As String)
Dim x&
For x = 1 To Len(LcdText)
  OutLcd Asc(Mid(LcdText, x, 1)), False, True, LcdCurrCntrl
Next
End Function
 
'Um benutzerdefinierte Symbole oder andere Sonderzeichen auszugeben, dient diese Funktion.
'die Benutzerdefinierten Zeichen haben den Wert 0-7.
Function WriteChar(ByVal LcdChar As Byte)
OutLcd LcdChar, False, True, LcdCurrCntrl
End Function
 
'Hier werden die benutzerdefinierten Symbole in den CG-RAM geschrieben.
'Komischerweise kann man die "LcdSymbol"-Variable nicht ByVal definieren. Dies sollte aber nicht stören.
Function DefineChar(ByVal SymbolNumber As Long, LcdSymbol As DspOwnSymbol, Optional ByVal Controller As DspContr = cUp)
Dim x&
OutLcd 64 + 8 * SymbolNumber, False, False, Controller
For x = 0 To 7
  OutLcd ConvertBintoDez("000" & LcdSymbol.sLine(x)), False, True, Controller
Next
End Function
 
'Der Einfachheit halber kann man jetzt auch 5x8 Pixel große Bitmaps als Icons setzen.
'So entfällt einiges an Code, wenn die Zeichen schon fest vom Programm definiert wurden.
'Leider funktioniert das Ganze bis jetzt nur mit Pictureboxen.
Function DefineCharbyPic(ByVal SymbolNumber As Long, SymbolPic As PictureBox, Optional ByVal Controller As DspContr = cUp)
Dim LineXdata As String
Dim OldScaleMode As ScaleModeConstants
Dim x&, y&
OldScaleMode = SymbolPic.ScaleMode
SymbolPic.ScaleMode = 3 ' Pixel
OutLcd 64 + 8 * SymbolNumber, False, False
For y = 0 To 7
  For x = 0 To 4
   If GetPixel(SymbolPic.hDC, x, y) = vbBlack Then
     LineXdata = LineXdata & "1"
   Else
     LineXdata = LineXdata & "0"
   End If
  Next
  OutLcd LcdPort, ConvertBintoDez("000" & LineXdata), False, True, Controller
  LineXdata = vbNullString
Next
SymbolPic.ScaleMode = OldScaleMode
End Function

'Verschiebt den Cursor nach links/rechts und auf Wunsch auch noch den Displayinhalt.
Function MoveCursor(ByVal Direction As DspDir, ByVal AlsoDisplay As Boolean, Optional ByVal Controller As DspContr = cUp)
OutLcd ConvertBintoDez("0001" & IIf(AlsoDisplay = True, 1, 0) & Direction & "00"), False, False, Controller
End Function
 
'Hier kann man festlegen, ob sich der Cursor inkrementiert oder dekrementiert verhalten soll (schreibt man das so?)
'Bei Movedisplay kann der Displayinhalt ählich wie bei LcdMoveCursor verschoben werden. Ich weiß aber immernoch nicht richtig, wie das abläuft.
Function Entrymodeset(ByVal PointerIncrement As Boolean, ByVal MoveDisplay As Boolean, Optional ByVal Controller As DspContr = cUp)
OutLcd ConvertBintoDez("000001" & IIf(PointerIncrement = True, 1, 0) & IIf(MoveDisplay, 1, 0)), False, False, Controller
End Function


Function TestLcdSpeed() As Currency
Dim x
Dim PerfFrequ As Currency
Dim PerfCntA As Currency
Dim PerfCntB As Currency
Dim PerfCntC As Currency
Dim PerfCntAvl As Long
Randomize
PerfCntAvl = QueryPerformanceFrequency(PerfFrequ) '"Kalibriert" die Uhr
QueryPerformanceCounter PerfCntA 'den Startwert für den Counter in PerfCntA schreiben
SetCursorHome
For x = 1 To 80
  WriteChar Rnd * 255
Next
QueryPerformanceCounter PerfCntB 'den Endwert für den Counter in PerfCntB schreiben
TestLcdSpeed = CDbl((PerfCntB - PerfCntA) / PerfFrequ) * 1000

End Function
