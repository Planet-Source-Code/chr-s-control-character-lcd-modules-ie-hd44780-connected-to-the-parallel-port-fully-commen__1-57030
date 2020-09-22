Attribute VB_Name = "INIDateien"
Dim Dateiname$

Private Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As _
   String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
   ByVal lpReturnedString As String, ByVal nSize As Long, _
   ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
   As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
   ByVal lpFileName As String) As Long

Public Function INISchreiben(ByVal DieSektion As String, ByVal DerEintrag As String, ByVal Wert As String) As Long
   WriteINI = WritePrivateProfileString(DieSektion, DerEintrag, Wert, App.Path & "\Optionen.ini")
End Function

Public Function INIlesen(DieSektion As String, DerEintrag As String) As String
   Temp$ = String(255, 0)
   X = GetPrivateProfileString(DieSektion, DerEintrag, "", Temp$, 255, App.Path & "\Optionen.ini")
   Temp$ = Left$(Temp$, X)
   INIlesen = Temp$
End Function


