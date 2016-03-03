Attribute VB_Name = "basTestRadix64"
Option Explicit
Option Base 0

' basTestRadix64: Tests for Radix 64 en/decoding functions
' Version 4a: Updated November 2003.
' Version 4. Updated 17 August 2002 to do comparisons with older versions.
' Version 2. Updated 16 January 2002.
' Version 1. Published 28 December 2000
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2000-2 D.I. Management Services Pty Limited,
' all rights reserved.

' You are free to use this code as part of your own applications
' provided you keep this copyright notice intact and acknowledge
' its authorship with the words:

'   "Contains cryptography software by David Ireland of
'   DI Management Services Pty Ltd <www.di-mgt.com.au>."

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
'--------------------------------------------------------------
' Timer Functions - from Litwin, Getz, Gilbert
'--------------------------------------------------------------
Declare Function wu_GetTime Lib "winmm.dll" Alias _
    "timeGetTime" () As Long
Private mlStartTime As Long
Private Sub ap_StartTimer()
    mlStartTime = wu_GetTime()
End Sub
Private Function ap_EndTimer() As Long
    ap_EndTimer = wu_GetTime() - mlStartTime
End Function
'--------------------------------------------------------------

' NOTE: The function cv_HexFromBytes()
' is in module basConvert.txt

Public Function TestBytesEnc64Rand()
    Dim nLen As Integer, i As Integer
    Dim abBytes() As Byte
    Dim sBase64 As String
    Dim abDecoded() As Byte
    ' Fill array with a random no of random binary values
    Randomize
    nLen = Int(32 * Rnd) + 1
    ReDim abBytes(nLen - 1)
    For i = 0 To nLen - 1
       abBytes(i) = CByte((Rnd * 256) And &HFF)
    Next
    
    ' Print hex values, encode it, then decode it again
    Debug.Print "Input:", cv_HexFromBytes(abBytes)
    sBase64 = EncodeBytes64(abBytes)
    Debug.Print "Encoded:", sBase64
    abDecoded = DecodeBytes64(sBase64)
    Debug.Print "Decoded:", cv_HexFromBytes(abDecoded)
    ' Compare byte arrays using hex conversion
    If cv_HexFromBytes(abBytes) <> cv_HexFromBytes(abDecoded) Then
        MsgBox "Radix64 Error"
    End If

End Function

Public Function TestEnc64Time()
    Dim nLen As Long, i As Long
    Dim lTime As Long
    Dim abBytes() As Byte
    Dim sBase64 As String
    Dim abDecoded() As Byte
    Dim sInput As String, sDecoded As String
    
    ' Fill a string with a lot of random binary values
    Randomize
    nLen = 100000 + Rnd() * 5
    ReDim abBytes(nLen - 1)
    For i = 0 To nLen - 1
       abBytes(i) = CByte((Rnd * 256) And &HFF)
    Next
    
    ' 1. Use latest version
    ' Encode it, then decode it again
    Call ap_StartTimer
    sBase64 = EncodeBytes64(abBytes)
    lTime = ap_EndTimer()
    Debug.Print "Encode with Bytes: " & nLen & " chars took " & lTime & " milliseconds"
    Call ap_StartTimer
    abDecoded = DecodeBytes64(sBase64)
    lTime = ap_EndTimer()
    Debug.Print "Decode with Bytes: " & nLen & " chars took " & lTime & " milliseconds"
    ' Check result equals the input
    If cv_HexFromBytes(abBytes) <> cv_HexFromBytes(abDecoded) Then
        MsgBox "Radix64 Error"
    End If

    ' 2. Compare times using deprecated String version
    ' Encode it, then decode it again
    sInput = StrConv(abBytes, vbUnicode)
    Call ap_StartTimer
    sBase64 = EncodeStr64(sInput)
    lTime = ap_EndTimer()
    Debug.Print "Encode with String: " & nLen & " chars took " & lTime & " milliseconds"
    Call ap_StartTimer
    sDecoded = DecodeStr64(sBase64)
    lTime = ap_EndTimer()
    Debug.Print "Decode with String: " & nLen & " chars took " & lTime & " milliseconds"
    ' Check result equals the input
    If sInput <> sDecoded Then
        MsgBox "Radix64 Error"
    End If


End Function




