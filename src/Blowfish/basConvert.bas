Attribute VB_Name = "basConvert"
Option Explicit
Option Base 0

' basConvert: Utilities to convert between byte arrays, hex strings,
' strings containing binary values, and 32-bit word arrays.

' NB: On 32-bit Unicode/CJK systems you may need to do a global
' replace of Asc() and Chr() with AscW() and ChrW() respectively.

' Version 2. November 2003: removed cv_BytesFromString which can be
' done with abBytes = StrConv(strInput, vbFromUnicode).
' - Added error handling to catch empty arrays.
' - Made HexFromByte public.
' Version 1. First published January 2002
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2000-2 D.I. Management Services Pty Limited,
' all rights reserved.

' You are free to use this code as part of your own applications
' provided you keep this copyright notice intact.

' This code may only be used as part of an application. It may
' not be reproduced or distributed separately by any means without
' the express written permission of the author.

' David Ireland and DI Management Services Pty Limited make no
' representations concerning either the merchantability of this
' software or the suitability of this software for any particular
' purpose. It is provided "as is" without express or implied
' warranty of any kind.

' Please forward comments or bug reports to <code@di-mgt.com.au>.
' The latest version of this source code can be downloaded from
' www.di-mgt.com.au/crypto.html.
'****************** END OF COPYRIGHT NOTICE*************************

' The Public Functions in this module are:
' cv_BytesFromHex(sInputHex): Returns array of bytes
' cv_WordsFromHex(sHex): Returns array of words (Longs)
' cv_HexFromWords(aWords): Returns hex string
' cv_HexFromBytes(aBytes()): Returns hex string
' cv_HexFromString(str): Returns hex string
' cv_StringFromHex(strHex): Returns string of ascii characters
' cv_GetHexByte(sInputHex, iIndex): Extracts iIndex'th byte from hex string
' RandHexByte(): Returns random byte as a 2-digit hex string
' HexFromByte(x): Returns 2-digit hex string representing byte x

Public Function cv_BytesFromHex(ByVal sInputHex As String) As Variant
' Returns array of bytes from hex string in big-endian order
' E.g. sHex="FEDC80" will return array {&HFE, &HDC, &H80}
    Dim i As Long
    Dim M As Long
    Dim aBytes() As Byte
    If Len(sInputHex) Mod 2 <> 0 Then
        sInputHex = "0" & sInputHex
    End If
    
    M = Len(sInputHex) \ 2
    If M <= 0 Then
        ' Version 2: Returns empty array
        cv_BytesFromHex = aBytes
        Exit Function
    End If
    
    ReDim aBytes(M - 1)
    
    For i = 0 To M - 1
        aBytes(i) = Val("&H" & Mid$(sInputHex, i * 2 + 1, 2))
    Next
    
    cv_BytesFromHex = aBytes

End Function

Public Function cv_WordsFromHex(ByVal sHex As String) As Variant
' Converts string <sHex> with hex values into array of words (long ints)
' E.g. "fedcba9876543210" will be converted into {&HFEDCBA98, &H76543210}
    Const ncLEN As Integer = 8
    Dim i As Long
    Dim nWords As Long
    Dim aWords() As Long
    
    nWords = Len(sHex) \ ncLEN
    If nWords <= 0 Then
        ' Version 2: Returns empty array
        cv_WordsFromHex = aWords
        Exit Function
    End If
    
    ReDim aWords(nWords - 1)
    For i = 0 To nWords - 1
        aWords(i) = Val("&H" & Mid(sHex, i * ncLEN + 1, ncLEN))
    Next
    
    cv_WordsFromHex = aWords
    
End Function

Public Function cv_HexFromWords(aWords) As String
' Converts array of words (Longs) into a hex string
' E.g. {&HFEDCBA98, &H76543210} will be converted to "FEDCBA9876543210"
    Const ncLEN As Integer = 8
    Dim i As Long
    Dim nWords As Long
    Dim sHex As String * ncLEN
    Dim iIndex As Long
    
    'Set up error handler to catch empty array
    On Error GoTo ArrayIsEmpty
    If Not IsArray(aWords) Then
        Exit Function
    End If
    
    nWords = UBound(aWords) - LBound(aWords) + 1
    cv_HexFromWords = String(nWords * ncLEN, " ")
    iIndex = 0
    For i = 0 To nWords - 1
        sHex = Hex(aWords(i))
        sHex = String(ncLEN - Len(sHex), "0") & sHex
        Mid$(cv_HexFromWords, iIndex + 1, ncLEN) = sHex
        iIndex = iIndex + ncLEN
    Next
    
ArrayIsEmpty:

End Function

Public Function cv_HexFromBytes(aBytes() As Byte) As String
' Returns hex string from array of bytes
' E.g. aBytes() = {&HFE, &HDC, &H80} will return "FEDC80"
    Dim i As Long
    Dim iIndex As Long
    Dim nLen As Long
    
    'Set up error handler to catch empty array
    On Error GoTo ArrayIsEmpty

    nLen = UBound(aBytes) - LBound(aBytes) + 1

    cv_HexFromBytes = String(nLen * 2, " ")
    iIndex = 0
    For i = LBound(aBytes) To UBound(aBytes)
        Mid$(cv_HexFromBytes, iIndex + 1, 2) = HexFromByte(aBytes(i))
        iIndex = iIndex + 2
    Next
    
ArrayIsEmpty:
    
End Function

Public Function cv_HexFromString(str As String) As String
' Converts string <str> of ascii chars to string in hex format
' str may contain chars of any value between 0 and 255.
' E.g. "abc." will be converted to "6162632E"
    Dim byt As Byte
    Dim i As Long
    Dim n As Long
    Dim iIndex As Long
    Dim sHex As String
    
    n = Len(str)
    sHex = String(n * 2, " ")
    iIndex = 0
    For i = 1 To n
        byt = CByte(Asc(Mid$(str, i, 1)) And &HFF)
        Mid$(sHex, iIndex + 1, 2) = HexFromByte(byt)
        iIndex = iIndex + 2
    Next
    cv_HexFromString = sHex
    
End Function

Public Function cv_StringFromHex(strHex As String) As String
' Converts string <strHex> in hex format to string of ascii chars
' with value between 0 and 255.
' E.g. "6162632E" will be converted to "abc."
    Dim i As Integer
    Dim nBytes As Integer
    
    nBytes = Len(strHex) \ 2
    cv_StringFromHex = String(nBytes, " ")
    For i = 0 To nBytes - 1
        Mid$(cv_StringFromHex, i + 1, 1) = Chr$(Val("&H" & Mid$(strHex, i * 2 + 1, 2)))
    Next
    
End Function

Public Function cv_GetHexByte(ByVal sInputHex As String, iIndex As Long) As Byte
' Extracts iIndex'th byte from hex string (starting at 1)
' E.g. cv_GetHexByte("fecdba98", 3) will return &HBA
    Dim i As Long
    i = 2 * iIndex
    If i > Len(sInputHex) Or i <= 0 Then
        cv_GetHexByte = 0
    Else
        cv_GetHexByte = Val("&H" & Mid$(sInputHex, i - 1, 2))
    End If
    
End Function

Public Function RandHexByte() As String
'   Returns a random byte as a 2-digit hex string
    Static stbInit As Boolean
    If Not stbInit Then
        Randomize
        stbInit = True
    End If
    
    RandHexByte = HexFromByte(CByte((Rnd * 256) And &HFF))
End Function

Public Function HexFromByte(ByVal x) As String
' Returns a 2-digit hex string for byte x
    x = x And &HFF
    If x < 16 Then
        HexFromByte = "0" & Hex(x)
    Else
        HexFromByte = Hex(x)
    End If
End Function


Public Function testWordsHex()
    Dim aWords
    
    aWords = cv_WordsFromHex("FEDCBA9876543210")
    Debug.Print cv_HexFromWords(aWords)
    
End Function



