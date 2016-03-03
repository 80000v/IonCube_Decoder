VERSION 5.00
Begin VB.Form frmBlowfish 
   Caption         =   "Blowfish Testbed Demo Version 6"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtCipher64 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4680
      Width           =   4695
   End
   Begin VB.Frame grpKeyForm 
      Caption         =   "Key form:"
      Height          =   615
      Left            =   1200
      TabIndex        =   16
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton optAlphaKey 
         Caption         =   "Alpha"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optHexKey 
         Caption         =   "Hex string"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCipherHex 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4200
      Width           =   4695
   End
   Begin VB.TextBox txtCipher 
      BackColor       =   &H80000004&
      Height          =   405
      Left            =   1200
      TabIndex        =   14
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox txtDecrypt 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt It"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtKeyAsString 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   4695
   End
   Begin VB.CommandButton cmdSetKey 
      Caption         =   "&Set Key"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt It"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtKey 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "fedcba9876543210"
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtPlain 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label Label7 
      Caption         =   "(Radix64):"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      Caption         =   "Copyright (C) 2000-3 DI Management Services Pty Ltd <www.di-mgt.com.au>"
      Height          =   495
      Left            =   1350
      TabIndex        =   13
      Top             =   6525
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "(In hex):"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Deciphered"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Active key:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Key:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Cipher text:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Plain text:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "frmBlowfish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' frmBlowfish.frm

' Examples of how to use the basic Blowfish functions
' to set the key, encrypt and decrypt using Blowfish
' and some of the byte-handling utilities.

' This is just a test-bed demo. It is not meant to be
' represesentative of good security code practices.

' Version 6. Published November 2003. Replaced deprecated
' String functions with Byte versions.
' Version 5. Published January 2002. Speed improvements
' and revised byte and word conversion functions.
' Replaced basByteUtils with basConvert.
' Thanks to Robert Garofalo and Doug Ward for advice
' and speed fixes incorporated here.
' Version 3. Published 20 January 2001. Fixed minor bug -
' stored ciphertext in a string instead of text box to
' avoid stripping trailing ascii zeroes.
' Thanks to Jim McCusker of epotec.com for this bug fix.
' Version 2. Published 28 December 2000. Added Radix64 fns and
' option to have alpha or hex key input.
' Version 1. Published 28 November 2000.
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


Dim aKey() As Byte
Dim abCipher() As Byte     ' Used to store ciphertext

Private Sub cmdEncrypt_Click()
    Dim abPlain() As Byte
    abPlain = StrConv(Me.txtPlain, vbFromUnicode)
    
    abCipher = blf_BytesEnc(abPlain)
    
    Me.txtCipher = StrConv(abCipher, vbUnicode)
    Me.txtCipherHex = cv_HexFromBytes(abCipher)
    Me.txtCipher64 = EncodeBytes64(abCipher)
    
    Me.cmdDecrypt.Enabled = True
End Sub

Private Sub cmdDecrypt_Click()
    Dim abPlain() As Byte
    abPlain = blf_BytesDec(abCipher)
    Me.txtDecrypt = StrConv(abPlain, vbUnicode)
End Sub

Private Sub cmdSetKey_Click()
' Get key bytes from user's string

    ' What format is it in
    If Me.optHexKey Then
        ' In hex format
        aKey() = cv_BytesFromHex(Me.txtKey)
    Else
        ' User has provided a plain alpha string
        aKey() = StrConv(Me.txtKey, vbFromUnicode)
    End If
    
    ' Show key
    Me.txtKeyAsString = cv_HexFromBytes(aKey())
    
    'Initialise key
    Call blf_KeyInit(aKey)
    
    ' Allow encrypt
    Me.cmdEncrypt.Enabled = True
    
    ' Put user in plaintext box
    Me.txtPlain.SetFocus

End Sub

