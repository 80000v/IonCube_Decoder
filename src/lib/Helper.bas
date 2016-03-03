Attribute VB_Name = "Helper"
Option Explicit
Option Compare Text


Public Const ERR_CANCEL_ALL& = vbObjectError Or &H1000

Public Const ERR_SKIP& = vbObjectError Or &H2000

Public Cancel As Boolean
Public CancelAll As Boolean

Public Skip As Boolean

'Konstantendeklationen für Registry.cls

'Registrierungsdatentypen
Public Const REG_SZ As Long = 1                         ' String
Public Const REG_BINARY As Long = 3                     ' Binär Zeichenfolge
Public Const REG_DWORD As Long = 4                      ' 32-Bit-Zahl

'Vordefinierte RegistrySchlüssel (hRootKey)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0

Public Const LocaleID_ENG = 1031

Public Const ERR_FILESTREAM = &H1000000
Public Const ERR_OPENFILE = vbObjectError Or ERR_FILESTREAM + 1
Private i, j As Integer

Public Declare Sub MemCopyAnyToAny Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Any, src As Any, ByVal Length&)
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal src As Any, ByVal Length&)
Public Declare Sub MemCopyX Lib "kernel32" Alias "RtlMoveMemory" _
(dest As Any, ByVal src As Long, ByVal Length&)

Public Declare Sub MemCopyAnyToStr Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal Length&)
Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As String, src As Long, ByVal Length&)

Public Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (dest As Long, ByVal src As String, ByVal Length&)
'Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As String, src As Long, ByVal Length&)
Public Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (dest As Long, ByVal src As Integer, ByVal Length&)
    
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SM_DBCSENABLED = 42
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer


Private BenchtimeA&, BenchtimeB&

'for mt_MT_Init to do a multiplation without 'overflow error'
Private Declare Function iMul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal dw2 As Long, ByVal dw3 As Long, ByVal dw4 As Long) As Long

Sub myStop()
   #If isDebug Then
'      Stop
   #End If
End Sub


'Ensure that 'myObjRegExp.MultiLine = True' else it will use the beginning of the string!
Function MulInt32&(a&, b&)
  MulInt32 = iMul(a, 0, b, 0)
End Function

Function AddInt32&(a As Double, b As Double)
  AddInt32 = "&h" & H32(a + b)
End Function


'Returns whether the user has DBCS enabled
Private Function isDBCSEnabled() As Boolean
   isDBCSEnabled = GetSystemMetrics(SM_DBCSENABLED)
End Function


Function LeftButton() As Boolean
    LeftButton = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Function RightButton() As Boolean
    RightButton = (GetAsyncKeyState(vbKeyRButton) And &H8000)
End Function

Function MiddleButton() As Boolean
    MiddleButton = (GetAsyncKeyState(vbKeyMButton) And &H8000)
End Function

Function MouseButton() As Integer
    If GetAsyncKeyState(vbKeyLButton) < 0 Then
        MouseButton = 1
    End If
    If GetAsyncKeyState(vbKeyRButton) < 0 Then
        MouseButton = MouseButton Or 2
    End If
    If GetAsyncKeyState(vbKeyMButton) < 0 Then
        MouseButton = MouseButton Or 4
    End If
End Function

Function KeyPressed(Key) As Boolean
   KeyPressed = GetAsyncKeyState(Key)
End Function

Public Function HexStringToString$(ByVal HexString$, Optional ByRef IsPrintable As Boolean)
   HexStringToString = Space(Len(HexString) \ 2)
   
   IsPrintable = True
   For i = 1 To Len(HexString) Step 2
      Dim tmpChar&
      tmpChar = "&h" & Mid$(HexString, i, 2)
      If IsPrintable Then
         IsPrintable = RangeCheck(tmpChar, &HFF, &H20)
      End If
      
      Mid$(HexStringToString, (i \ 2) + 1) = Chr(tmpChar)
   Next
End Function

Public Function HexvaluesToString$(Hexvalues$)
   Dim tmpChar
   For Each tmpChar In Split(Hexvalues)
      'HexvaluesToString = HexvaluesToString & ChrB("&h" & tmpchar) & ChrB(0)
      'Note ChrB("&h98") & ChrB(0) is not correct translated
      HexvaluesToString = HexvaluesToString & Chr("&h" & tmpChar)
   Next
End Function

Public Function ValuesToHexString$(Data As StringReader, Optional Seperator = " ")
'ValuesToHexString = ""
   With Data
      .EOS = False
   Dim DisableAutoMove_old As Boolean
   DisableAutoMove_old = .DisableAutoMove
   .DisableAutoMove = False
   
   
      Do Until .EOS
         ValuesToHexString = ValuesToHexString & H8(.int8) & Seperator
      Loop
   
   .DisableAutoMove = DisableAutoMove_old
   
   End With
  
End Function


Sub MaxProc(ParamArray values())
   Dim item
   For Each item In values
      values(LBound(values)) = IIf(Max < item, item, Max)
   Next
End Sub


Function Max(ParamArray values())
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function Min(ParamArray values())
   Dim item
   Min = &H7FFFFFFF
   For Each item In values
      Min = IIf(Min > item, item, Min)
   Next
End Function

Function limit(value&, Optional ByVal upperLimit = &H7FFFFFFF, Optional lowerLimit = 0) As Long
   'limit = IIf(Value > upperLimit, upperLimit, IIf(Value < lowerLimit, lowerLimit, Value))

   If (value > upperLimit) Then _
      limit = upperLimit _
   Else _
      If (value < lowerLimit) Then _
         limit = lowerLimit _
      Else _
         limit = value
   
End Function

Function isEven(Number As Long) As Boolean
   isEven = ((Number And 1) = 0)
End Function

Function RangeCheck(ByVal value&, Max&, Optional Min& = 0, Optional ErrText, Optional ErrSource$) As Boolean
   RangeCheck = (Min <= value) And (value <= Max)
   If (RangeCheck = False) And (IsMissing(ErrText) = False) Then _
       Err.Raise vbObjectError, ErrSource, _
           ErrText & " Value must between '" & Min & "'  and '" & Max & "' !"
End Function

Public Function H8(ByVal value As Long)
   H8 = Right(String(1, "0") & Hex(value), 2)
End Function

Public Function H16(ByVal value As Long)
   H16 = Right(String(3, "0") & Hex(value), 4)
End Function

Public Function H32(ByVal value As Double)
   If value >= 0 Then
      H32 = Hex(value)
   Else
    ' split Number in High a Low part...
      Dim High&, Low&
      High = Int(value / &H10000)
      Low = value - (CDbl(High) * &H10000)
      
      H32 = H16(High) & H16(Low)
   End If
   
   H32 = Right(String(7, "0") & H32, 8)
End Function

Public Function Swap(ByRef a, ByRef b)
   Swap = b
   b = a
   a = Swap
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_r  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_l("Summe",4) -> "umme"
Public Function BlockAlign_r(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Right(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_r = RawString & Space(Blocksize - Len(RawString))
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_l  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_l("Summe",4) -> "umme"
Public Function BlockAlign_l(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Left(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_l = Space(Blocksize - Len(RawString)) & RawString
End Function

Public Function qw()
   Cancel = True
   Do
      DoEvents
   Loop While Cancel = True
End Function
Public Function szNullCut$(zeroString$)
   Dim nullCharPos&
   nullCharPos = InStr(1, zeroString, Chr(0))
   If nullCharPos Then
      szNullCut = Left(zeroString, nullCharPos - 1)
   Else
      szNullCut = zeroString
   End If
   
End Function


Public Function Inc(ByRef value, Optional Increment& = 1)
   value = value + Increment
   Inc = value
End Function

Public Function Dec(ByRef value, Optional DeIncrement& = 1)
   value = value - DeIncrement
   Dec = value
End Function



Public Function CollectionToArray(Collection As Collection) As Variant
   
   Dim tmp
   ReDim tmp(Collection.Count - 1)
   
   Dim i
   i = LBound(tmp)
   
   Dim item
   For Each item In Collection
      tmp(i) = item
      Inc i
   Next
   
   CollectionToArray = tmp
   
End Function
Public Function isString(StringToCheck) As Boolean
   'isString = False
   Dim i&
   For i = 1 To Len(StringToCheck)
      If RangeCheck(Asc(Mid$(StringToCheck, i, 1)), &H7F, &H20) Then
      
      Else
         Exit Function
      End If
   Next
   
   isString = True
   
End Function



'Searches for some string and then starts there to crop
Function strCropWithSeek$(Text$, LeftString$, RightString$, Optional errorvalue, Optional SeektoStrBeforeSearch$)
   strCropWithSeek = strCrop1(Text$, LeftString$, RightString$, errorvalue, _
            InStr(1, Text, SeektoStrBeforeSearch))
End Function


Function strCrop1$(ByVal Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1)
   
   Dim cutend&, cutstart&
      cutstart = InStr(StartSearchAt, Text, LeftString)
   If cutstart Then
      cutstart = cutstart + Len(LeftString)
      cutend = InStr(cutstart, Text, RightString)
      If cutend > cutstart Then
         strCrop1 = Mid$(Text, cutstart, cutend - cutstart)
      Else
        'is Rightstring empty?
         If RightString = "" Then
            strCrop1 = Mid$(Text, cutstart)
         Else
            strCrop1 = errorvalue
         End If
      End If
   Else
      strCrop1 = errorvalue
   End If

End Function

Function strCropAndDelete(Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1, Optional ReplaceString$ = "")
   strCropAndDelete = strCrop1(Text$, LeftString$, RightString$, errorvalue, StartSearchAt)
   Text = Replace(Text, LeftString & strCropAndDelete & RightString, ReplaceString, , , vbTextCompare)
End Function



Function strCrop$(Text$, LeftString$, RightString$, Optional errorvalue, Optional StartSearchAt = 1)
   
   Dim cutend&, cutstart&
      cutend = InStr(StartSearchAt, Text, RightString)
   If cutend Then
      cutstart = InStrRev(Text, LeftString, cutend, vbBinaryCompare) + Len(LeftString)
      strCrop = Mid$(Text, cutstart, cutend - cutstart)
   Else
      strCrop = errorvalue
   End If

End Function

Function MidMbcs(ByVal str As String, Start, Length)
    MidMbcs = StrConv(MidB$(StrConv(str, vbFromUnicode), Start, Length), vbUnicode)
End Function


Function strCutOut$(str$, pos&, Length&, Optional TextToInsert = "")
   strCutOut = Mid(str, pos, Length)
   str$ = Mid(str, 1, pos - 1) & TextToInsert & Mid(str, pos + Length)
End Function
'<<
Public Function shl8(value&, bits&)
   shl8 = (value * (2 ^ bits)) And &HFF
End Function

'>>
Public Function shr8(value&, bits&)
   shr8 = Int(value / (2 ^ bits)) And &HFF
End Function


'<<
Public Function shl16(value&, bits&)
   shl16 = (value * (2 ^ bits)) And 65535
End Function

'>>
Public Function shr16(value&, bits&)
   If value < 0 Then
    'negative values
'      value = &HD63862D8 '&hD63862D8
'      value = &HAA3862D8 '&hD63862D8
      shr16 = Int(value / (2 ^ bits))
      shr16 = shr16 And 65535
   Else
      shr16 = Int((value / (2 ^ bits))) And 65535
   End If
   
   Debug.Assert Left(H32(value), 4) = H16(shr16)
End Function

Public Function ror(value&, bits&)
   ror = (shl8(value, bits) Or shr8(value, 8 - bits)) And &HFF
End Function

Public Function rol(value&, bits&)
   rol = (shr8(value, bits) Or shl8(value, 7 - bits)) And &HFF
End Function

Public Function Int16ToUInt32&(value%)
      Const N_0x8000& = 32767
      If value >= 0 Then
         Int16ToUInt32 = value
      Else
         Int16ToUInt32 = CLng(value And N_0x8000) + N_0x8000
      End If
      
End Function




Public Function BenchStart()

   BenchtimeA = GetTickCount

End Function
Public Function BenchEnd()

   BenchtimeB = GetTickCount
   Debug.Print Time & " - " & BenchtimeB - BenchtimeA

End Function


Public Function FileExists(FileName) As Boolean
   On Error GoTo FileExists_err
   FileExists = FileLen(FileName)

FileExists_err:
End Function

Public Function StrQuote(ByRef Text) As String
   StrQuote = """" & Text & """"
End Function

Public Function Brackets(ByRef Text As String) As String
   Brackets = "(" & Text & ")"
End Function


Public Function IsAlreadyInCollection(CollectionToTest As Collection, Key$) As Boolean
   Dim description$, Number&, Source$
   description = Err.description
   Number = Err.Number
   Source = Err.Source
   
      On Error Resume Next
      CollectionToTest.item Key
      IsAlreadyInCollection = (Err = 0)
      
   Err.description = description
   Err.Number = Number
   Err.Source = Source


End Function

'Public Sub ArrayEnsureBounds(Arr)
'
''   Dim tmp_ptr&
''   MemCopy tmp_ptr, VarPtr(Arr) + 8, 4 ' resolve Variant
''   MemCopy tmp_ptr, tmp_ptr, 4               ' get arraypointer
''
''   Dim bIsNullArray As Boolean
''   bIsNullArray = (tmp_ptr = 0)
'' On Error Resume Next
'
'   Dim bIsNullArray As Boolean
'   bIsNullArray = (Not Not Arr) = 0 'use vbBug to get pointer to Arr
'
''   Rnd 1 ' catch Expression too complex error that is cause by the bug
''On Error GoTo 0
'
''   Exit Function
'
'   If bIsNullArray Then
'
'   ElseIf (UBound(Arr) - LBound(Arr)) < 0 Then
'   Else
'      Exit Function
'   End If
'
'   ReDim Arr(0)
'   ArrayEnsureBounds = True
'   Exit Function

Public Function ArrayEnsureBounds(Arr) As Boolean

On Error GoTo Array_err
  ' IsArray(Arr)=False        ->  13 - Type Mismatch
  ' [Arr has no Elements]     ->  9 - Subscript out of range
  ' ZombieArray[arr=Array()]  -> GoTo Array_new
   If UBound(Arr) - LBound(Arr) < 0 Then GoTo Array_new
Exit Function
Array_err:
Select Case Err
    Case 9, 13
Array_new:
      ArrayDelete Arr
      ArrayEnsureBounds = True

'   Case Else
'      Err.Raise Err.Number, "", "Error in ArrayEnsureBounds: " & Err.Description

End Select

End Function



Public Sub ArrayAdd(Arr, Optional Element = "")
  'Extent Array if not completely new array
   If ArrayEnsureBounds(Arr) <> True Then
      ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
   End If
   Arr(UBound(Arr)) = Element

End Sub


'Public Sub ArrayAdd(Arr As Variant, Optional element = "")
'' Is that already a Array?
'   If IsArray(Arr) Then
'      ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
'
' ' VarType(Arr) = vbVariant must be
'   Else 'If VarType(Arr) = vbVariant Then
'      ReDim Arr(0)
'   End If
'
'   Arr(UBound(Arr)) = element
'
'End Sub

Public Sub ArrayRemoveLast(Arr)
   ReDim Preserve Arr(LBound(Arr) To UBound(Arr) - 1)
End Sub

Public Sub ArrayDelete(Arr)
   ReDim Arr(0)
   'Arr = Array()
   'Set Arr = Nothing
End Sub


Public Function ArrayGetLast(Arr)
ArrayEnsureBounds Arr
   ArrayGetLast = Arr(UBound(Arr))
End Function
Public Sub ArraySetLast(Arr, Element)
ArrayEnsureBounds Arr
    Arr(UBound(Arr)) = Element
End Sub
Public Sub ArrayAppendLast(Arr(), Element)
ArrayEnsureBounds Arr
    Arr(UBound(Arr)) = Arr(UBound(Arr)) & Element
End Sub


Public Function ArrayGetFirst(Arr)
ArrayEnsureBounds Arr
   ArrayGetFirst = Arr(LBound(Arr))
End Function
Public Sub ArraySetFirst(Arr, Element)
ArrayEnsureBounds Arr
    Arr(LBound(Arr)) = Element
End Sub
Public Sub ArrayAppendFirst(Arr, Element)
ArrayEnsureBounds Arr
    Arr(LBound(Arr)) = Arr(LBound(Arr)) & Element
End Sub




Function DelayedReturn(Now As Boolean) As Boolean
   Static LastState As Boolean
   
   DelayedReturn = LastState
   
   LastState = Now
   
End Function







'Private Sub QuickSort( _
'                      ByRef ArrayToSort As Variant, _
'                      ByVal Low As Long, _
'                      ByVal High As Long)
'Dim vPartition As Variant, vTemp As Variant
'Dim i As Long, j As Long
'  If Low > High Then Exit Sub  ' Rekursions-Abbruchbedingung
'  ' Ermittlung des Mittenelements zur Aufteilung in zwei Teilfelder:
'  vPartition = ArrayToSort((Low + High) \ 2)
'  ' Indizes i und j initial auf die äußeren Grenzen des Feldes setzen:
'  i = Low: j = High
'  Do
'    ' Von links nach rechts das linke Teilfeld durchsuchen:
'    Do While ArrayToSort(i) < vPartition
'      i = i + 1
'    Loop
'    ' Von rechts nach links das rechte Teilfeld durchsuchen:
'    Do While ArrayToSort(j) > vPartition
'      j = j - 1
'    Loop
'    If i <= j Then
'      ' Die beiden gefundenen, falsch einsortierten Elemente
'austauschen:
'      vTemp = ArrayToSort(j)
'      ArrayToSort(j) = ArrayToSort(i)
'      ArrayToSort(i) = vTemp
'      i = i + 1
'      j = j - 1
'    End If
'  Loop Until i > j  ' Überschneidung der Indizes
'  ' Rekursive Sortierung der ausgewählten Teilfelder. Um die
'  ' Rekursionstiefe zu optimieren, wird (sofern die Teilfelder
'  ' nicht identisch groß sind) zuerst das kleinere
'  ' Teilfeld rekursiv sortiert.
'  If (j - Low) < (High - i) Then
'    QuickSort ArrayToSort, Low, j
'    QuickSort ArrayToSort, i, High
'  Elsea
'    QuickSort ArrayToSort, i, High
'    QuickSort ArrayToSort, Low, j
'  End If
'End Sub
'
'

Public Sub myDoEvents()
   DoEvents
   
   Skip_Test
   CancelAll_Test
End Sub

Public Sub Skip_Test()
   If Skip = True Then
      
      Skip = False
      Err.Raise ERR_SKIP, , "User pressed the skip key."
      
   End If
  
End Sub



Public Sub CancelAll_Test()
   If CancelAll = True Then
      
      CancelAll = False
      Err.Raise ERR_CANCEL_ALL, , "User pressed the cancel key."
      
   End If
  
End Sub

Public Function FileLoad$(FileName$)
   Dim File As New FileStream
   With File
      .Create FileName, False, False, True
      FileLoad = .FixedString(-1)
      .CloseFile
   End With
End Function

Public Sub FileSave(FileName$, Data$)
   Dim File As New FileStream
   With File
      .Create FileName, True, False, False
      .FixedString(-1) = Data
      .CloseFile
   End With
End Sub
Function ShowFlags$(value&, ParamArray descriptions())
   Dim outString As New clsStrCat
   
   Dim i
   For i = 0 To 15
      Dim testFlag&
      testFlag = 2 ^ i
      If (testFlag And value) Then
         
         Dim description$
         description = "Bit " & i
         On Error Resume Next
         description = descriptions(i)
         
         outString.Concat "  0x" & H16(testFlag) & "  " & description & vbCrLf
      End If
   Next
   
   ShowFlags = outString.value
End Function

Function ShowFlagsSimple(value&, Seperator$, ParamArray descriptions())
   
#If isDebug Then
#Else
   Dim outString()
   
   Dim i
   For i = 0 To 15
      Dim testFlag&
      testFlag = 2 ^ i
      If (testFlag And value) Then
         
         Dim description$
         description = "Bit " & i
         On Error Resume Next
         description = descriptions(i)
         
         ArrayAdd outString, description
      End If
   Next
   
   ShowFlagsSimple = Join(outString, Seperator)
   
#End If
   
End Function


Function joinCol(C As Collection, Optional Delimitier$ = " ")
   Dim tmpArr
   ReDim tmpArr(1 To C.Count)
   
   Dim i
   For i = 1 To C.Count
      tmpArr(i) = C(i)
   Next
   
   joinCol = Join(tmpArr, Delimitier)
End Function


Public Function Adler32$(Data As StringReader, Optional InitL& = 1, Optional InitH& = 0)
   With Data
'            Dim a
            
            Dim L&, h&
            h = InitH: L = InitL
'            a = GetTickCount
' taken out for performance reason
'               .EOS = False
'               .DisableAutoMove = False
'               Do Until .EOS
'                 'The largest prime less than 2^16
'                  l = (.int8 + l) Mod 65521 '&HFFF1
'                  H = (H + l) Mod 65521 '&HFFF1
'                  If (l And 8) Then myDoEvents
'               Loop
'
'            Debug.Print "a: ", GetTickCount - a 'Benchmark: 20203

 '           a = GetTickCount
               
               Dim StrCharPos&, tmpBuff$
               tmpBuff = StrConv(.mvardata, vbFromUnicode, LocaleID_ENG)
'               tmpBuff = .mvardata
               For StrCharPos = 1 To Len(.mvardata)
                  'The largest prime less than 2^16
                  L = (AscB(MidB$(tmpBuff, StrCharPos, 1)) + L) ' Mod 65521 '&HFFF1
                  h = (h + L) ' Mod 65521 '&HFFF1
                  
                ' Do Mod &HFFF1 only if h might overflow
                  'Originally   the
                  'algorithm had  to calculate  the modulo every  step,  but   this   revised  form only does  the modulo
                  'calculation, every 5552 iterations. Why this number? Since h is an unsigned long int (thus with an
                  'upper limit of  2^32-1) we have to stop the summation before we have an overflow. How can we
                  'calculate this? Assume that our buffer is filled with bytes with the value 0xFF (255), the maximum
                  'value for an unsigned char. Then s2 will get filled up pretty quickly ,  and we can see that it will
                  'exceed the limit 2^32-1 after 5552 iterations,
                  '/* 5552 is the largest n such that 255n(n+1)/2 + (n+1)(BASE-1) <= 2^32-1 */
                
                  If (0 = (StrCharPos Mod 5552)) Or _
                           (StrCharPos = Len(.mvardata)) Then
                     L = L Mod 65521  '&HFFF1
                     h = h Mod 65521  '&HFFF1
                     myDoEvents
                  End If
                  
               Next
'            Debug.Print "b: ", GetTickCount - a 'Benchmark: 5969

      Adler32 = H16(h) & H16(L)
   End With
End Function


'Reverse a Dword Example:
' 0x00112233 gets 0x33221100  or
' 0x01234567 gets 0x76543210
Public Function DwordReverse&(value&)
   
'TODO: Faster implemetation
   Dim tmp As New StringReader
   With tmp
      .int32 = value
      .Position = 0
      DwordReverse = "&h" & H8(.int8) & H8(.int8) & H8(.int8) & H8(.int8) '.int32
   End With
End Function
