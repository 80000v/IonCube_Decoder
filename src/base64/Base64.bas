Attribute VB_Name = "modBase64"
Option Explicit

'Private Declare Function ArrPtr Lib "msvbvm50.dll" _
'    Alias "VarPtr" (Ptr() As Any) As Long '<-- VB5
Private Declare Function ArrPtr Lib "msvbvm60.dll" _
    Alias "VarPtr" (Ptr() As Any) As Long '<-- VB6
Private Declare Sub PokeLng Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal dest As Long, Source As Long, _
    Optional ByVal Bytes As Long = 4)

Private Base64Initialized As Boolean
Private Base64EncodeByte(0 To 63) As Byte
Private Base64EncodeWord(0 To 63) As Integer
Private Base64DecodeByte(0 To 255) As Byte
Private Base64DecodeWord(0 To 255) As Integer
Private Unicode2Ascii(0 To 16383) As Integer
Private Ascii2Unicode(0 To 255) As Integer
Private Const Base64EmptyByte As Byte = 61 '='
Private Const Base64EmptyWord As Integer = 61


Public Sub Base64Init(Optional Base64Charset As String = _
      "0123456789" & _
      "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
      "abcdefghijklmnopqrstuvwxyz" & _
      "+/")
  
   Dim Base64CharsetLen
   Base64CharsetLen = Len(Base64Charset)
   If Base64CharsetLen <> 64 Then
      Err.Raise vbObjectError, , "The of Base64Charset is not 64! It's " & Base64CharsetLen & "! Now guess what that algo is called Base64 and why we need we need 64 chars here. :)"
   End If
  
      
  'Deklarationen:
  Dim i As Integer
  Dim Code As Integer
  
  'Base64-Tabellen füllen:
  For i = 0 To 63
    Code = Asc(Mid$(Base64Charset, i + 1, 1))
    Base64EncodeByte(i) = Code
    Base64DecodeByte(Code) = i
    Base64EncodeWord(i) = Code
    Base64DecodeWord(Code) = i
  Next i
  
  'Unicode-Tabellen füllen:
  For i = 0 To 255
    Code = AscW(Chr$(i))
    Ascii2Unicode(i) = Code
    Unicode2Ascii(Code) = i
  Next i
  
  Base64Initialized = True
End Sub


Public Sub Base64EncodeArray( _
    ByRef Bytes() As Byte, _
    ByRef OutBytes() As Byte _
  )
  'Deklarationen:
  Dim LB As Long
  Dim UB As Long
  Dim OutUB As Long
  Dim i As Long
  Dim j As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  
  'Input-Array checken:
  LB = LBound(Bytes)
  UB = UBound(Bytes)
  If UB < LB Then Exit Sub
  
  'Benötigte Größe berechnen:
  OutUB = LB + ((UB - LB) \ 3) * 4 + 3
  ReDim OutBytes(LB To OutUB)
  
  'Los gehts:
  If Not Base64Initialized Then Base64Init
  j = LB
  For i = LB To UB - 2 Step 3
    'Aus 3 Bytes...
    b1 = Bytes(i)
    b2 = Bytes(i + 1)
    b3 = Bytes(i + 2)
    
    '...werden 4 Base64-Bytes:
    OutBytes(j) = Base64EncodeByte(b1 \ &H4)
    OutBytes(j + 1) = Base64EncodeByte((b1 And &H3) * &H10 Or b2 \ &H10)
    OutBytes(j + 2) = Base64EncodeByte((b2 And &HF) * &H4 Or b3 \ &H40)
    OutBytes(j + 3) = Base64EncodeByte(b3 And &H3F)
    
    j = j + 4
  Next i
  
  'Ggf. fehlende Bytes berücksichtigen:
  Select Case UB - i
  Case 0 '2 Bytes fehlen:
    b1 = Bytes(i)
    
    OutBytes(j) = Base64EncodeByte(b1 \ &H4)
    OutBytes(j + 1) = Base64EncodeByte((b1 And &H3) * &H10)
    OutBytes(j + 2) = Base64EmptyByte
    OutBytes(j + 3) = Base64EmptyByte
  Case 1 '1 Byte fehlt:
    b1 = Bytes(i)
    b2 = Bytes(i + 1)
    
    OutBytes(j) = Base64EncodeByte(b1 \ &H4)
    OutBytes(j + 1) = Base64EncodeByte((b1 And &H3) * &H10 Or b2 \ &H10)
    OutBytes(j + 2) = Base64EncodeByte((b2 And &HF) * &H4)
    OutBytes(j + 3) = Base64EmptyByte
  End Select
End Sub


Public Sub Base64DecodeArray( _
    ByRef Bytes() As Byte, _
    ByRef OutBytes() As Byte _
  )
  'Deklarationen:
  Dim LB As Long
  Dim UB As Long
  Dim OutUB As Long
  Dim i As Long
  Dim j As Long
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  
  'Input-Array checken:
  LB = LBound(Bytes)
  UB = UBound(Bytes)
  If UB < LB Then Exit Sub
  
  'Benötigte Größe berechnen:
  If Bytes(UB) = Base64EmptyByte Then UB = UB - 1
  If Bytes(UB) = Base64EmptyByte Then UB = UB - 1
  OutUB = LB + (UB - LB) * 3 \ 4
  ReDim OutBytes(LB To OutUB)
  
  'Los gehts:
  If Not Base64Initialized Then Base64Init
  j = LB
  For i = LB To UB - 3 Step 4
    'Aus 4 Base64-Bytes...
    b1 = Base64DecodeByte(Bytes(i))
    b2 = Base64DecodeByte(Bytes(i + 1))
    b3 = Base64DecodeByte(Bytes(i + 2))
    b4 = Base64DecodeByte(Bytes(i + 3))
    
    '...werden 3 Bytes:
    OutBytes(j) = b1 * &H4 Or b2 \ &H10
    OutBytes(j + 1) = (b2 And &HF) * &H10 Or b3 \ &H4
    OutBytes(j + 2) = (b3 And &H3) * &H40 Or b4
    
    j = j + 3
  Next i
  
  'Ggf. fehlende Bytes berücksichtigen:
  Select Case OutUB - j
  Case 0 '1 Byte fehlt:
    b1 = Base64DecodeByte(Bytes(i))
    b2 = Base64DecodeByte(Bytes(i + 1))
    
    OutBytes(j) = b1 * &H4 Or b2 \ &H10
  Case 1 '2 Bytes fehlen:
    b1 = Base64DecodeByte(Bytes(i))
    b2 = Base64DecodeByte(Bytes(i + 1))
    b3 = Base64DecodeByte(Bytes(i + 2))
    
    OutBytes(j) = b1 * &H4 Or b2 \ &H10
    OutBytes(j + 1) = (b2 And &HF) * &H10 Or b3 \ &H4
  End Select
End Sub


Public Function Base64EncodeUnicode2( _
    ByRef Text As String _
  ) As String
  'Platz für Ergebnis deklarieren:
  Dim Bytes() As Byte
  
  If Len(Text) Then
    Base64EncodeArray StrConv(Text, vbFromUnicode), Bytes()
    Base64EncodeUnicode2 = StrConv(Bytes, vbUnicode)
  End If
End Function


Public Function Base64DecodeUnicode2( _
    ByRef Text As String _
  ) As String
  'Platz für Ergebnis deklarieren:
  Dim Bytes() As Byte
  
  If Len(Text) Then
    Base64DecodeArray StrConv(Text, vbFromUnicode), Bytes()
    Base64DecodeUnicode2 = StrConv(Bytes, vbUnicode)
  End If
End Function


Public Function Base64EncodeAscii( _
    ByRef Text As String _
  ) As String
  Dim chars() As Integer 'Unicode-Darstellung des Textes
  Dim SavePtr As Long    'Original Daten-Pointer
  Dim SADescrPtr As Long 'Safe Array Descriptor
  Dim DataPtr As Long    'pvData - Daten-Pointer
  Dim CountPtr As Long   'Pointer zu nElements
  Dim TextLen As Long
  Dim i As Long
  
  Dim chars64() As Integer 'Unicode-Darstellung des Base64-Textes
  Dim SavePtr64 As Long    'Original Daten-Pointer
  Dim SADescrPtr64 As Long 'Safe Array Descriptor
  Dim DataPtr64 As Long    'pvData - Daten-Pointer
  Dim CountPtr64 As Long   'Pointer zu nElements
  Dim TextLen64 As Long
  Dim j As Long
  
  Dim b1 As Integer
  Dim b2 As Integer
  Dim b3 As Integer
  
  'Platzbedarf checken:
  TextLen = Len(Text)
  If TextLen = 0 Then Exit Function
  TextLen64 = ((TextLen + 2) \ 3) * 4
  Base64EncodeAscii = Space$(TextLen64)
  
  'Input-String durch Integer-Array mappen:
  ReDim chars(1 To 1)
  SavePtr = VarPtr(chars(1))
  PokeLng VarPtr(SADescrPtr), ByVal ArrPtr(chars)
  DataPtr = SADescrPtr + 12
  CountPtr = SADescrPtr + 16
  PokeLng DataPtr, StrPtr(Text)
  PokeLng CountPtr, TextLen
  
  'Output-String: (Base64) durch Integer-Array mappen:
  ReDim chars64(0 To 0)
  SavePtr64 = VarPtr(chars64(0))
  PokeLng VarPtr(SADescrPtr64), ByVal ArrPtr(chars64)
  DataPtr64 = SADescrPtr64 + 12
  CountPtr64 = SADescrPtr64 + 16
  PokeLng DataPtr64, StrPtr(Base64EncodeAscii)
  PokeLng CountPtr64, TextLen64
  
  'Los gehts:
  If Not Base64Initialized Then Base64Init
  For i = 1 To TextLen - 2 Step 3
    b1 = chars(i)
    b2 = chars(i + 1)
    b3 = chars(i + 2)
    
    chars64(j) = Base64EncodeWord(b1 \ &H4)
    chars64(j + 1) = Base64EncodeWord((b1 And &H3) * &H10 Or b2 \ &H10)
    chars64(j + 2) = Base64EncodeWord((b2 And &HF) * &H4 Or b3 \ &H40)
    chars64(j + 3) = Base64EncodeWord(b3 And &H3F)
    
    j = j + 4
  Next i
  
  'Ggf. fehlende Bytes berücksichtigen:
  Select Case TextLen - i
  Case 0 '2 Bytes fehlen:
    b1 = chars(i)
    
    chars64(j) = Base64EncodeWord(b1 \ &H4)
    chars64(j + 1) = Base64EncodeByte((b1 And &H3) * &H10)
    chars64(j + 2) = Base64EmptyWord
    chars64(j + 3) = Base64EmptyWord
  Case 1 '1 Byte fehlt:
    b1 = chars(i)
    b2 = chars(i + 1)
    
    chars64(j) = Base64EncodeWord(b1 \ &H4)
    chars64(j + 1) = Base64EncodeWord((b1 And &H3) * &H10 Or b2 \ &H10)
    chars64(j + 2) = Base64EncodeWord((b2 And &HF) * &H4)
    chars64(j + 3) = Base64EmptyWord
  End Select
  
  'Integer-Arrays restaurieren:
  PokeLng DataPtr64, SavePtr64
  PokeLng CountPtr64, 1
  '
  PokeLng DataPtr, SavePtr
  PokeLng CountPtr, 1
End Function


Public Function Base64EncodeUnicode( _
    ByRef Text As String _
  ) As String
  'Input-Variablen (Unicode):
  Dim chars() As Integer 'Unicode-Darstellung des Textes
  Dim SavePtr As Long    'Original Daten-Pointer
  Dim SADescrPtr As Long 'Safe Array Descriptor
  Dim DataPtr As Long    'pvData - Daten-Pointer
  Dim CountPtr As Long   'Pointer zu nElements
  Dim TextLen As Long
  Dim i As Long
  'Output-Variablen (Base64):
  Dim chars64() As Integer 'Unicode-Darstellung des Base64-Textes
  Dim SavePtr64 As Long    'Original Daten-Pointer
  Dim SADescrPtr64 As Long 'Safe Array Descriptor
  Dim DataPtr64 As Long    'pvData - Daten-Pointer
  Dim CountPtr64 As Long   'Pointer zu nElements
  Dim TextLen64 As Long
  Dim j As Long
  'Sonstiges:
  Dim b1 As Integer
  Dim b2 As Integer
  Dim b3 As Integer
  
  'Platzbedarf bestimmen:
  TextLen = Len(Text)
  If TextLen = 0 Then Exit Function
  TextLen64 = ((TextLen + 2) \ 3) * 4
  Base64EncodeUnicode = Space$(TextLen64)
  
  'Input-String durch Integer-Array mappen:
  ReDim chars(1 To 1)
  SavePtr = VarPtr(chars(1))
  PokeLng VarPtr(SADescrPtr), ByVal ArrPtr(chars)
  DataPtr = SADescrPtr + 12
  CountPtr = SADescrPtr + 16
  PokeLng DataPtr, StrPtr(Text)
  PokeLng CountPtr, TextLen
  
  'Output-String durch Integer-Array mappen:
  ReDim chars64(0 To 0)
  SavePtr64 = VarPtr(chars64(0))
  PokeLng VarPtr(SADescrPtr64), ByVal ArrPtr(chars64)
  DataPtr64 = SADescrPtr64 + 12
  CountPtr64 = SADescrPtr64 + 16
  PokeLng DataPtr64, StrPtr(Base64EncodeUnicode)
  PokeLng CountPtr64, TextLen64
  
  'Los gehts:
  If Not Base64Initialized Then Base64Init
  For i = 1 To TextLen - 2 Step 3
    'Aus 3 Unicode-Words...
    b1 = Unicode2Ascii(chars(i))
    b2 = Unicode2Ascii(chars(i + 1))
    b3 = Unicode2Ascii(chars(i + 2))
    
    '...werden 4 Base64-Words:
    chars64(j) = Base64EncodeWord(b1 \ &H4)
    chars64(j + 1) = Base64EncodeWord((b1 And &H3) * &H10 Or b2 \ &H10)
    chars64(j + 2) = Base64EncodeWord((b2 And &HF) * &H4 Or b3 \ &H40)
    chars64(j + 3) = Base64EncodeWord(b3 And &H3F)
    
    j = j + 4
  Next i
  
  'Ggf. fehlende Words berücksichtigen:
  Select Case TextLen - i
  Case 0 '2 Words fehlen:
    b1 = Unicode2Ascii(chars(i))
    
    chars64(j) = Base64EncodeWord(b1 \ &H4)
    chars64(j + 1) = Base64EncodeWord((b1 And &H3) * &H10)
    chars64(j + 2) = Base64EmptyWord
    chars64(j + 3) = Base64EmptyWord
  Case 1 '1 Word fehlt:
    b1 = Unicode2Ascii(chars(i))
    b2 = Unicode2Ascii(chars(i + 1))
    
    chars64(j) = Base64EncodeWord(b1 \ &H4)
    chars64(j + 1) = Base64EncodeWord((b1 And &H3) * &H10 Or b2 \ &H10)
    chars64(j + 2) = Base64EncodeWord((b2 And &HF) * &H4)
    chars64(j + 3) = Base64EmptyWord
  End Select
  
  'Integer-Arrays restaurieren:
  PokeLng DataPtr64, SavePtr64
  PokeLng CountPtr64, 1
  '
  PokeLng DataPtr, SavePtr
  PokeLng CountPtr, 1
End Function


Public Function Base64DecodeAscii( _
    ByRef Text As String _
  ) As String
  Dim TextLen As Long
  Dim j As Long
  'Sonstiges:
  Dim b1 As Integer
  Dim b2 As Integer
  Dim b3 As Integer
  Dim b4 As Integer
  
   
  
  'Vorab-Prüfung:

   Dim TextLen64 As Long
   Dim PosEnd&
   PosEnd = InStrRev(Text, "=")
   If PosEnd = 0 Then
      TextLen64 = Len(Text)
   Else
     ' -1 because "=" has the length 1
      TextLen64 = PosEnd - 1
   End If
  
   If TextLen64 = 0 Then Exit Function



 'Input
  Dim chars64() As Byte
  chars64 = StrConv(Text, vbFromUnicode)
 
 'because chars64 array starts with 0 (and not 1)
  Dec TextLen64

'
'  'Input-String durch Integer-Array mappen:
'
'
'  'Input-Variablen (Unicode):
'  Dim chars64() As Integer 'Unicode-Darstellung des Base64-Textes
'  Dim SavePtr64 As Long    'Original Daten-Pointer
'  Dim SADescrPtr64 As Long 'Safe Array Descriptor
'  Dim DataPtr64 As Long    'pvData - Daten-Pointer
'  Dim CountPtr64 As Long   'Pointer zu nElements
'
'  ReDim chars64(1 To 1)
'
'  SavePtr64 = VarPtr(chars64(1))
'  PokeLng VarPtr(SADescrPtr64), ByVal ArrPtr(chars64)
'
'  DataPtr64 = SADescrPtr64 + 12
'  PokeLng DataPtr64, StrPtr(Text)
'
'  CountPtr64 = SADescrPtr64 + 16
'  PokeLng CountPtr64, TextLen64
'
  'Platzbedarf bestimmen:
   
  If chars64(TextLen64) = Base64EmptyWord Then TextLen64 = TextLen64 - 1
  If chars64(TextLen64) = Base64EmptyWord Then TextLen64 = TextLen64 - 1
  TextLen = ((TextLen64 + 1) * 3 \ 4)
'  Base64DecodeAscii = Space$(TextLen)
'

 'Output
  Dim chars() As Byte
  ReDim chars(TextLen - 1)



'  'Output-String durch Integer-Array mappen:
'  'Output-Variablen (Base64):
'  Dim Chars() As Integer 'Unicode-Darstellung des Textes
'  Dim SavePtr As Long    'Original Daten-Pointer
'  Dim SADescrPtr As Long 'Safe Array Descriptor
'  Dim DataPtr As Long    'pvData - Daten-Pointer
'  Dim CountPtr As Long   'Pointer zu nElements
'
'  ReDim Chars(0 To 0)
'  SavePtr = VarPtr(Chars(0))
'  PokeLng VarPtr(SADescrPtr), ByVal ArrPtr(Chars)
'
'  DataPtr = SADescrPtr + 12
'  PokeLng DataPtr, StrPtr(Base64DecodeAscii)
'
'  CountPtr = SADescrPtr + 16
'  PokeLng CountPtr, TextLen
  
  'Los gehts:
  If Not Base64Initialized Then Base64Init
  
  Dim i As Long
  For i = 0 To TextLen64 - 3 Step 4
    'Aus 4 Base64-Words...
    b1 = Base64DecodeWord(chars64(i))
    b2 = Base64DecodeWord(chars64(i + 1))
    b3 = Base64DecodeWord(chars64(i + 2))
    b4 = Base64DecodeWord(chars64(i + 3))
    
    '...werden 3 Words:
'    Chars(j) = b1 * &H4 Or b2 \ &H10 '<<4
'    Chars(j + 1) = (b2 And &HF) * &H10 Or b3 \ &H4 '<<2
'    Chars(j + 2) = (b3 And &H3) * &H40 Or b4
    
    chars(j) = b1 * &H4 Or b2 \ &H10          '<<4
    chars(j + 1) = 255 And (b2 * &H10 Or b3 \ &H4) '<<2
    chars(j + 2) = 255 And (b3 * &H40 Or b4)
    
    
    j = j + 3
  Next i
  
  'Ggf. fehlende Words berücksichtigen:
  Select Case TextLen64 - i
  Case 1 '1 Word fehlt:
    b1 = Base64DecodeWord(chars64(i))
    b2 = Base64DecodeWord(chars64(i + 1))
    
    chars(j) = b1 * &H4 Or b2 \ &H10
  Case 2 '2 Words fehlen:
    b1 = Base64DecodeWord(chars64(i))
    b2 = Base64DecodeWord(chars64(i + 1))
    b3 = Base64DecodeWord(chars64(i + 2))
    
    chars(j) = b1 * &H4 Or b2 \ &H10
    chars(j + 1) = (b2 And &HF) * &H10 Or b3 \ &H4
  End Select
  
'  'Integer-Arrays restaurieren:
'  PokeLng DataPtr, SavePtr
'  PokeLng CountPtr, 1
'  '
'  PokeLng DataPtr64, SavePtr64
'  PokeLng CountPtr64, 1
  
  Base64DecodeAscii = StrConv(chars, vbUnicode)
End Function

Private Function GetNextValidChar&(ByRef mArray, ByRef index&)
   GetNextValidChar = mArray(index)
   Do While (GetNextValidChar < &H20) And (index < UBound(mArray))
      index = index + 1
      GetNextValidChar = mArray(index)
   Loop
End Function
Public Function Base64DecodeUnicode( _
    ByRef Text As String _
  ) As String
  'Input-Variablen (Base64):
  Dim chars64() As Integer 'Unicode-Darstellung des Base64-Textes
  Dim SavePtr64 As Long    'Original Daten-Pointer
  Dim SADescrPtr64 As Long 'Safe Array Descriptor
  Dim DataPtr64 As Long    'pvData - Daten-Pointer
  Dim CountPtr64 As Long   'Pointer zu nElements
  Dim TextLen64 As Long
  Dim i As Long
  'Output-Variablen (Unicode):
  Dim chars() As Integer 'Unicode-Darstellung des Textes
  Dim SavePtr As Long    'Original Daten-Pointer
  Dim SADescrPtr As Long 'Safe Array Descriptor
  Dim DataPtr As Long    'pvData - Daten-Pointer
  Dim CountPtr As Long   'Pointer zu nElements
  Dim TextLen As Long
  Dim j As Long
  'Sonstiges:
  Dim b1 As Integer
  Dim b2 As Integer
  Dim b3 As Integer
  Dim b4 As Integer
  
  'Vorab-Prüfung:
  TextLen64 = Len(Text)
  If TextLen64 = 0 Then Exit Function
  
  'Input-String durch Integer-Array mappen:
  ReDim chars64(1 To 1)
  SavePtr64 = VarPtr(chars64(1))
  PokeLng VarPtr(SADescrPtr64), ByVal ArrPtr(chars64)
  DataPtr64 = SADescrPtr64 + 12
  CountPtr64 = SADescrPtr64 + 16
  PokeLng DataPtr64, StrPtr(Text)
  PokeLng CountPtr64, TextLen64
  
  'Platzbedarf bestimmen:
  If chars64(TextLen64) = Base64EmptyWord Then TextLen64 = TextLen64 - 1
  If chars64(TextLen64) = Base64EmptyWord Then TextLen64 = TextLen64 - 1
  TextLen = TextLen64 * 3 \ 4
  Base64DecodeUnicode = Space$(TextLen)
  
  'Output-String durch Integer-Array mappen:
  ReDim chars(0 To 0)
  SavePtr = VarPtr(chars(0))
  PokeLng VarPtr(SADescrPtr), ByVal ArrPtr(chars)
  DataPtr = SADescrPtr + 12
  CountPtr = SADescrPtr + 16
  PokeLng DataPtr, StrPtr(Base64DecodeUnicode)
  PokeLng CountPtr, TextLen
  
  'Los gehts:
  If Not Base64Initialized Then Base64Init
  For i = 1 To TextLen64 - 3 Step 4
  
'   Debug.Assert j < &H36
    'Aus 4 Base64-Words...
    
    b1 = Base64DecodeWord(GetNextValidChar(chars64, i))
    b2 = Base64DecodeWord(GetNextValidChar(chars64, i + 1))
    b3 = Base64DecodeWord(GetNextValidChar(chars64, i + 2))
    b4 = Base64DecodeWord(GetNextValidChar(chars64, i + 3))

    
    '...werden 3 Words:
    chars(j) = Ascii2Unicode(b1 * &H4 Or b2 \ &H10)
    chars(j + 1) = Ascii2Unicode((b2 And &HF) * &H10 Or b3 \ &H4)
    chars(j + 2) = Ascii2Unicode((b3 And &H3) * &H40 Or b4)
    
    j = j + 3
  Next i
  
  'Ggf. fehlende Words berücksichtigen:
  Select Case TextLen64 - i
  Case 1 '1 Word fehlt:
    b1 = Base64DecodeWord(chars64(i))
    b2 = Base64DecodeWord(chars64(i + 1))
    
    chars(j) = Ascii2Unicode(b1 * &H4 Or b2 \ &H10)
    j = j + 1
  Case 2 '2 Words fehlen:
    b1 = Base64DecodeWord(chars64(i))
    b2 = Base64DecodeWord(chars64(i + 1))
    b3 = Base64DecodeWord(chars64(i + 2))
    
    chars(j) = Ascii2Unicode(b1 * &H4 Or b2 \ &H10)
    chars(j + 1) = Ascii2Unicode((b2 And &HF) * &H10 Or b3 \ &H4)
    j = j + 2
  End Select
   
   'Grösse Korrigieren (wegen möglicher Zeilenumbrüche)
   Base64DecodeUnicode = Mid(Base64DecodeUnicode, 1, j)
   
  'Integer-Arrays restaurieren:
  PokeLng DataPtr, SavePtr
  PokeLng CountPtr, 1
  '
  PokeLng DataPtr64, SavePtr64
  PokeLng CountPtr64, 1
End Function
