Attribute VB_Name = "basBlowfish"
Option Explicit
Option Base 0

' basBlowfish: Bruce Schneier's Blowfish algorithm in VB
' Core routines.

' Version 6. November 2003. Removed redundant functions blf_Enc()
' and blf_Dec().
' Version 5: January 2002. Speed improvements.
' Version 4: 12 May 2001. Fixed maxkeylen size from bits to bytes.
' First published October 2000.
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2000-2 D.I. Management Services Pty Limited,
' all rights reserved.

' You are free to use this code as part of your own applications
' provided you keep this copyright notice intact and acknowledge
' its authorship with the words:

'   "Contains cryptography software by David Ireland of
'   DI Management Services Pty Ltd <www.di-mgt.com.au>."

' If you use it as part of a web site, please include a link
' to our site in the form
' <A HREF="http://www.di-mgt.com.au/crypto.html">Cryptography
' Software Code</a>

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

' Public Functions in this module:
' blf_EncipherBlock: Encrypts two words
' blf_DecipherBlock: Decrypts two words
' blf_Initialise: Initialise P & S arrays using key
' blf_KeyInit: Initialise using byte-array key
' blf_EncryptBytes: Encrypts an block of 8 bytes
' blf_DecryptBytes: Decrypts an block of 8 bytes
'
' Superseded functions:
' blf_Key: Initialise using byte-array and its length
' blf_Enc: Encrypts an array of words
' blf_Dec: Decrypts an array of words

Private Const ncROUNDS  As Integer = 16
Private Const ncMAXKEYLEN As Integer = 56
' Version 4: ncMAXKEYLEN was previously incorrectly set as 448
' (bits vs bytes)
' Thanks to Robert Garofalo for pointing this out.

Private Function blf_F(x As Long) As Long
    Dim a As Byte, b As Byte, C As Byte, d As Byte
    Dim y As Long
    
    Call uwSplit(x, a, b, C, d)
    
    y = uw_WordAdd(blf_S(0, a), blf_S(1, b))
    y = y Xor blf_S(2, C)
    y = uw_WordAdd(y, blf_S(3, d))
    blf_F = y
    
End Function

Public Function blf_EncipherBlock(xL As Long, xR As Long)
    Dim i As Integer
    Dim temp As Long
    
    'Reverse Xl and XR
    
    For i = 0 To ncROUNDS - 1
        xL = xL Xor blf_P(i)
        xR = blf_F(xL) Xor xR
        
      ' Swap xL and xR
        temp = xL
        xL = xR
        xR = temp
        
    Next
 
 ' Swap xL and xR
    temp = xL
    xL = xR
    xR = temp
    
    xR = xR Xor blf_P(ncROUNDS)
    xL = xL Xor blf_P(ncROUNDS + 1)
        
End Function

Public Function blf_DecipherBlock(xL As Long, xR As Long)
    Dim i As Integer
    Dim temp As Long
    
  'Difference to Encrypt (Indexes)
    For i = ncROUNDS + 1 To 2 Step -1
        xL = xL Xor blf_P(i)
        xR = blf_F(xL) Xor xR
        temp = xL
        xL = xR
        xR = temp
    Next
      
    temp = xL
    xL = xR
    xR = temp
    
    xR = xR Xor blf_P(1)
    xL = xL Xor blf_P(0)
        
End Function

Public Function blf_Initialise(aKey() As Byte, nKeyBytes As Integer)
    Dim i As Integer, j As Integer, K As Integer
    Dim wData As Long, wDataL As Long, wDataR As Long
    
    Call blf_LoadArrays     ' Initialise P and S arrays

 ' Init P-Box with key
    j = 0
    For i = 0 To (ncROUNDS + 2 - 1)
        wData = &H0
        For K = 0 To 3
            wData = uw_ShiftLeftBy8(wData) Or aKey(j)
            j = j + 1
            If j >= nKeyBytes Then j = 0
        Next K
        blf_P(i) = blf_P(i) Xor wData
'        Debug.Print "P-Box: " i, H32(blf_P(i))
    Next i
    

'Stop
  ' Init for P-Box
    wDataL = &H0
    wDataR = &H0
    For i = 0 To (ncROUNDS + 2 - 1) Step 2
        Call blf_EncipherBlock(wDataL, wDataR)
        
        blf_P(i) = wDataL
        blf_P(i + 1) = wDataR
    Next i
    
  ' Init S-Box
    For i = 0 To 3
        For j = 0 To 255 Step 2
            Call blf_EncipherBlock(wDataL, wDataR)
    
            blf_S(i, j) = wDataL
            blf_S(i, j + 1) = wDataR
        Next j
    Next i

End Function

Public Function blf_Key(aKey() As Byte, nKeyLen As Integer) As Boolean
    blf_Key = False
    If nKeyLen < 0 Or nKeyLen > ncMAXKEYLEN Then
        Exit Function
    End If
    
    Call blf_Initialise(aKey, nKeyLen)
    
    blf_Key = True
End Function

Public Function blf_KeyInit(aKey() As Byte) As Boolean
' Added Version 5: Replacement for blf_Key to avoid specifying keylen
' Version 6: Added error checking for input
    Dim nKeyLen As Integer
    
    blf_KeyInit = False
    
    'Set up error handler to catch empty array
    On Error GoTo ArrayIsEmpty

    nKeyLen = UBound(aKey) - LBound(aKey) + 1
    If nKeyLen < 0 Or nKeyLen > ncMAXKEYLEN Then
        Exit Function
    End If
    
    Call blf_Initialise(aKey, nKeyLen)
    
    
    blf_KeyInit = True
    
ArrayIsEmpty:

End Function

Public Function blf_EncryptBytes(aBytes() As Byte)
' aBytes() must be 8 bytes long
' Revised Version 5: January 2002. To use faster uwJoin and uwSplit fns.
    Dim wordL As Long, wordR As Long
    
    ' Convert to 2 x words
    wordL = uwJoin(aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    wordR = uwJoin(aBytes(4), aBytes(5), aBytes(6), aBytes(7))
    ' Encrypt it
    Call blf_EncipherBlock(wordL, wordR)
    ' Put back into bytes
    Call uwSplit(wordL, aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    Call uwSplit(wordR, aBytes(4), aBytes(5), aBytes(6), aBytes(7))

End Function

Public Function blf_DecryptBytes(aBytes() As Byte)
' aBytes() must be 8 bytes long
' Revised Version 5:: January 2002. To use faster uwJoin and uwSplit fns.
    Dim wordL As Long, wordR As Long
    
    ' Convert to 2 x words
    wordL = uwJoin(aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    wordR = uwJoin(aBytes(4), aBytes(5), aBytes(6), aBytes(7))
    
    ' Decrypt it
    Call blf_DecipherBlock(wordL, wordR)
    
    ' Put back into bytes
    Call uwSplit(wordL, aBytes(0), aBytes(1), aBytes(2), aBytes(3))
    Call uwSplit(wordR, aBytes(4), aBytes(5), aBytes(6), aBytes(7))

End Function

