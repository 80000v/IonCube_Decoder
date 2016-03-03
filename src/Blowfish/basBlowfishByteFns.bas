Attribute VB_Name = "basBlowfishByteFns"
Option Explicit
Option Base 0

' basBlowfishByteFns: Wrapper functions to call Blowfish algorithms

' Version 6. November 2003. Added this module with new Byte functions
' Blowfish in Visual Basic first published October 2000.
'************************* COPYRIGHT NOTICE*************************
' This code was originally written in Visual Basic by David Ireland
' and is copyright (c) 2000-3 D.I. Management Services Pty Limited,
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

' The functions in this module are:
' blf_BytesRaw(abData, bEncrypt): En/Deciphers bytes abData without padding
' blf_BytesEnc(abData): Pads and enciphers byte array abData with current key
' blf_BytesDec(abData): Deciphers byte array abData with current key and unpads
' PadBytes(abData): Pads byte array to next multiple of 8 bytes
' UnpadBytes(abData): Removes padding after decryption
' blf_BytesEncRawCBC(abData, abInitV): Encrypts abData in CBC mode
' blf_BytesEncCBC(abData, abInitV): Pads and encrypts abData in CBC mode
' blf_BytesDecRawCBC(abData, abInitV): Decrypts abData in CBC mode
' blf_BytesDecCBC(abData, abInitV): Decrypts abData in CBC mode and unpads

' To set current key, call blf_KeyInit(aKey())
'   where aKey() is the key as an array of Bytes

' NB The functions in this module deal with data of any length, but
' if you only want to deal with an 8-byte block, use
' blf_EncryptBytes() and blf_DecryptBytes() in module basBlowfish

' Use faster API call to copy bytes
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal lpDestination As Any, ByVal lpSource As Any, ByVal Length As Long)

Public Function blf_BytesRaw(abData() As Byte, bEncrypt As Boolean) As Variant
' New function added version 6.
' Encrypts or decrypts byte array abData without padding using to current key.
' Similar to blf_BytesEnc and blf_BytesDec, but does not add padding
' and ignores trailing odd bytes.
' ECB mode - each block is en/decrypted independently
    Dim nLen As Long
    Dim nBlocks As Long
    Dim iBlock As Long
    Dim j As Long
    Dim abOutput() As Byte
    Dim abBlock(7) As Byte
    Dim iIndex As Long
    
    ' Calc number of 8-byte blocks (ignore odd trailing bytes)
    nLen = UBound(abData) - LBound(abData) + 1
    nBlocks = nLen \ 8
    
    ReDim abOutput(nBlocks * 8 - 1)
    
    ' Work through in blocks of 8 bytes
    iIndex = 0
    For iBlock = 1 To nBlocks
        ' Get the next block of 8 bytes
        CopyMemory VarPtr(abBlock(0)), VarPtr(abData(iIndex)), 8&

        ' En/Decrypt the block according to flag
        If bEncrypt Then
            Call blf_EncryptBytes(abBlock())
        Else
            Call blf_DecryptBytes(abBlock())
        End If
        
        ' Copy to output string
        CopyMemory VarPtr(abOutput(iIndex)), VarPtr(abBlock(0)), 8&
        
        iIndex = iIndex + 8
    Next
    
    blf_BytesRaw = abOutput
    
End Function

Public Function blf_BytesEnc(abData() As Byte) As Variant
' Encrypts byte array abData after adding PKCS#5/RFC2630/RFC3370 padding
' NB always adds padding - use blf_BytesRaw() if you don't want padding
' ECB mode
' Returns encrypted byte array as a variant.
' Requires key and boxes to be already set up.
' New in Version 6.

    Dim abOutput() As Byte
    
    abOutput = PadBytes(abData)
    abOutput = blf_BytesRaw(abOutput, True)
    
    blf_BytesEnc = abOutput
End Function

Public Function blf_BytesDec(abData() As Byte) As Variant
' Decrypts byte array abData assuming PKCS#5/RFC2630/RFC3370 padding and ECB mode
' NB always removes valid padding - use blf_BytesRaw() if you don't want padding
' Returns encrypted byte array as a variant.
' Requires key and boxes to be already set up.
' New in Version 6.

    Dim abOutput() As Byte
    
    abOutput = blf_BytesRaw(abData, False)
    abOutput = UnpadBytes(abOutput)
    
    blf_BytesDec = abOutput
End Function

Public Function PadBytes(abData() As Byte) As Variant
' Pad data bytes to next multiple of 8 bytes as per PKCS#5/RFC2630/RFC3370
    Dim nLen As Long
    Dim nPad As Integer
    Dim abPadded() As Byte
    Dim i As Long
    
    'Set up error handler for empty array
    On Error GoTo ArrayIsEmpty

    nLen = UBound(abData) - LBound(abData) + 1
    nPad = ((nLen \ 8) + 1) * 8 - nLen
    
    ReDim abPadded(nLen + nPad - 1)  ' Pad with # of pads (1-8)
    If nLen > 0 Then
        CopyMemory VarPtr(abPadded(0)), VarPtr(abData(0)), nLen
    End If
    For i = nLen To nLen + nPad - 1
        abPadded(i) = CByte(nPad)
    Next
    
ArrayIsEmpty:
    PadBytes = abPadded

End Function

Public Function UnpadBytes(abData() As Byte) As Variant
' Strip PKCS#5/RFC2630/RFC3370-style padding
    Dim nLen As Long
    Dim nPad As Long
    Dim abUnpadded() As Byte
    Dim i As Long
    
    'Set up error handler for empty array
    On Error GoTo ArrayIsEmpty
    
    nLen = UBound(abData) - LBound(abData) + 1
    If nLen = 0 Then GoTo ArrayIsEmpty
    ' Get # of padding bytes from last char
    nPad = abData(nLen - 1)
    If nPad > 8 Then nPad = 0   ' In case invalid
    If nLen - nPad > 0 Then
        ReDim abUnpadded(nLen - nPad - 1)
        CopyMemory VarPtr(abUnpadded(0)), VarPtr(abData(0)), nLen - nPad
    End If

ArrayIsEmpty:
    UnpadBytes = abUnpadded
    
End Function

Public Function TestPadBytes()
    Dim abData() As Byte
    
    abData = StrConv("abc", vbFromUnicode)
    abData = PadBytes(abData)
    Stop
    abData = UnpadBytes(abData)
    Stop
    
End Function

Private Sub bXorBytes(aByt1() As Byte, aByt2() As Byte, nBytes As Long)
' XOR's bytes in array aByt1 with array aByt2
' Returns results in aByt1
' i.e. aByt1() = aByt1() XOR aByt2()
    Dim i As Long
    For i = 0 To nBytes - 1
        aByt1(i) = aByt1(i) Xor aByt2(i)
    Next
End Sub

Public Function blf_BytesEncRawCBC(abData() As Byte, abInitV() As Byte) As Variant
' Encrypts byte array <abData> in CBC mode
' using byte array <abInitV> as initialisation vector.
' Returns ciphertext as variant array of bytes.
' Requires key and boxes to be already set up.
' New in Version 6.
    Dim nLen As Long
    Dim nBlocks As Long
    Dim iBlock As Long
    Dim abBlock(7) As Byte
    Dim iIndex As Long
    Dim abReg(7) As Byte    ' Feedback register
    Dim abOutput() As Byte
    
    ' Initialisation vector should be a 8-byte array
    ' so ReDim just to make sure
    ' This will add zero bytes if too short or chop off any extra
    ReDim Preserve abInitV(7)
    
    ' Calc number of 8-byte blocks
    nLen = UBound(abData) - LBound(abData) + 1
    nBlocks = nLen \ 8
    
    ' Dimension output
    ReDim abOutput(nBlocks * 8 - 1)
    
    ' C_0 = IV
    CopyMemory VarPtr(abReg(0)), VarPtr(abInitV(0)), 8&
    
    ' Work through string in blocks of 8 bytes
    iIndex = 0
    For iBlock = 1 To nBlocks
        ' Fetch next block from input
        CopyMemory VarPtr(abBlock(0)), VarPtr(abData(iIndex)), 8&
        
        
        ' XOR with feedback register = Pi XOR C_i-1
        Call bXorBytes(abBlock, abReg, 8)
        
        ' Encrypt the block Ci = Ek(Pi XOR C_i-1)
        Call blf_EncryptBytes(abBlock())
        
        
        ' Store in feedback register Reg = Ci
        CopyMemory VarPtr(abReg(0)), VarPtr(abBlock(0)), 8&
        
        ' Copy to output string
        CopyMemory VarPtr(abOutput(iIndex)), VarPtr(abBlock(0)), 8&

        iIndex = iIndex + 8
    Next
    
    blf_BytesEncRawCBC = abOutput
    
End Function

Public Function blf_BytesDecRawCBC(abData() As Byte, abInitV() As Byte) As Variant
' Decrypts byte array <abData> in CBC mode
' using byte array <abInitV> as initialisation vector.
' Returns plaintext as variant array of bytes.
' Requires key and boxes to be already set up.
' New in Version 6.
'    Dim strIn As String
'    Dim strOut As String
    
    Dim nLen As Long
    Dim nBlocks As Long
    Dim iBlock As Long
    Dim abBlock(7) As Byte
    Dim iIndex As Long
    Dim abReg(7) As Byte    ' Feedback register
    Dim abStore(7) As Byte
    Dim abOutput() As Byte
    
    ' Initialisation vector should be a 8-byte array
    ' so ReDim just to make sure
    ' This will add zero bytes if too short or chop off any extra
    ReDim Preserve abInitV(7)
    
    ' Calc number of 8-byte blocks
    nLen = UBound(abData) - LBound(abData) + 1
    nBlocks = nLen \ 8
    
    ' Dimension output
    ReDim abOutput(nBlocks * 8 - 1)
    
    ' C_0 = IV
    CopyMemory VarPtr(abReg(0)), VarPtr(abInitV(0)), 8&
    
    ' Work through string in blocks of 8 bytes
    iIndex = 0
    For iBlock = 1 To nBlocks
        ' Fetch next block from input
        CopyMemory VarPtr(abBlock(0)), VarPtr(abData(iIndex)), 8&
        
      
        ' Save C_i-1
        CopyMemory VarPtr(abStore(0)), VarPtr(abBlock(0)), 8&
        
        ' Decrypt the block Dk(Ci)
        Call blf_DecryptBytes(abBlock())
        
        ' XOR with feedback register = C_i-1 XOR Dk(Ci)
        Call bXorBytes(abBlock, abReg, 8)
        
        
        ' Store in feedback register Reg = C_i-1
        CopyMemory VarPtr(abReg(0)), VarPtr(abStore(0)), 8&
                
        ' Copy to output string
        CopyMemory VarPtr(abOutput(iIndex)), VarPtr(abBlock(0)), 8&

        iIndex = iIndex + 8
    Next
    
    blf_BytesDecRawCBC = abOutput
    
End Function

Public Function blf_BytesEncCBC(abData() As Byte, abInitV() As Byte) As Variant
' Encrypts byte array abData after adding PKCS#5/RFC2630/RFC3370 padding
' NB always adds padding - use blf_BytesEncRawCBC() if you don't want padding
' CBC mode
' Returns encrypted byte array as a variant.
' Requires key and boxes to be already set up.
' New in Version 6.

    Dim abOutput() As Byte
    
    abOutput = PadBytes(abData)
    abOutput = blf_BytesEncRawCBC(abOutput, abInitV)
    
    blf_BytesEncCBC = abOutput
End Function

Public Function blf_BytesDecCBC(abData() As Byte, abInitV() As Byte) As Variant
' Decrypts byte array abData assuming PKCS#5/RFC2630/RFC3370 padding and CBC mode
' NB always removes valid padding - use blf_BytesDecRawCBC() if you don't want padding
' Returns encrypted byte array as a variant.
' Requires key and boxes to be already set up.
' New in Version 6.

    Dim abOutput() As Byte
    
    abOutput = blf_BytesDecRawCBC(abData, abInitV)
    abOutput = UnpadBytes(abOutput)
    
    blf_BytesDecCBC = abOutput
End Function



