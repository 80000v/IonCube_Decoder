VERSION 5.00
Begin VB.Form FrmDeCodeLicense 
   Caption         =   "Decode License File"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "FrmDecodeLicenseFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmd_close 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Na what will happen if you click that one? Hint pressing ESC 'll do da same as clicking this button."
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Cmd_Decrypt 
      Caption         =   "Decrypt"
      Default         =   -1  'True
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      ToolTipText     =   "Opens LicenseFile manually and shows all data(including these make_license.exe hides from you)"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txt_Lic 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   6975
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "Patches & runs IC's 'make_license.exe'"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Txt_MemberID 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "<MemberID>"
      ToolTipText     =   $"FrmDecodeLicenseFile.frx":1CFA
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Txt_PassPhrase 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "<PassPhrase>"
      ToolTipText     =   "<PassPhrase> This is really essential for decryption"
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox txt_LicFile 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "<LicFile>"
      ToolTipText     =   "<LicFile> Relative or full path to IC-License file"
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "FrmDeCodeLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MakeLic_File$ = "data\make_license.exe"
Const MakeLic_PatchOffset& = &H29BCC
Const MakeLic_PE_BuildTimeStamp& = &H4A16A8F0

Public ByteCode_Key&

Private preamble_data$
Dim ICLD_DataEntries&

Private WithEvents Console As Console
Attribute Console.VB_VarHelpID = -1

Private Function GetEncStr$(ByRef InputStream As StringReader, Optional Enforcesflag)


   Dim Record As New StringReader
   Dim DataSize&
   DataSize = InputStream.int32
   Record = InputStream.FixedString(DataSize)
'   Record = InputStream.FixedString(InputStream.int32)
   'Debug.Print "ICLD-Record(" & H16(DataSize) & "): " & ValuesToHexString(Record)
   
   
   With Record
      .Position = 0
      
    ' The enforces byte setting is only present in for Value
    '(If you read the attrib name - just leave this optional parameter out)
      If IsMissing(Enforcesflag) = False Then
         Enforcesflag = .FixedString(1)
      End If
      
      
      Const XORKey32& = &HE9FC23B1 '3925615537
      Const XORKey16& = XORKey32 And &HFFFF&
      
      Dim Size&
      Size = .int16 Xor XORKey16 '&H23B1& '9137
      
      
      Dim StringLen&
      StringLen = .Length - .Position '= (.Length - FillBytes - 2) 'StrSize; NullByte String Terminator
      
      Dim ExtraBytes&
      ExtraBytes = Size - StringLen
      If ExtraBytes <> 0 Then
         Debug.Print "LCLD-Record ExtraBytes: " & ExtraBytes
         Stop
      End If
      
      
      Dim XorKey As New StringReader
      XorKey.int32 = XORKey32
      XorKey.EOS = False
      
      Dim output As New StringReader
      Do Until output.Length >= Size
         
         output.int8 = .int8 Xor XorKey.int8
       
       ' cycle through key
         If XorKey.EOS Then XorKey.EOS = False
      Loop
            
   End With
   
   GetEncStr = output
   
End Function
Private Function GenPassWordHash$(PassPhrase$, MemberID&)

 ' Generate full password...
   Dim SHA_PassWord As New StringReader
   With SHA_PassWord
      .Position = 0
      
      .FixedString = PassPhrase
     '+ 0x17 Extradata add
      .int32 = &H111C0702 '0
      .int32 = MemberID      '4
      'SomeObj            '8
      .int32 = &H3900040A     '0
      .int32 = &H38010F       '4
      .int16 = &H138          '8
      .int8 = 0               'A
      'EndObj             '8 + B = 0x17)
         
   End With
 
 ' ...  to Hash it with SHA256
   
   GenPassWordHash = SHA256(SHA_PassWord.Data)

End Function
Private Function SHA256$(Data$, Optional HashAsBinaryData As StringReader)
   Dim HashAsDwords()
 
   Dim mySHA256 As New CSHA256
   SHA256 = mySHA256.SHA256(Data, HashAsDwords)
   
   Set HashAsBinaryData = New StringReader
   With HashAsBinaryData
      Dim i&
      For i = LBound(HashAsDwords) To UBound(HashAsDwords)
         .int32 = HashAsDwords(i)
      Next
   End With
End Function
   
 ' CTR -> CounTerR mode
 ' http://en.wikipedia.org/wiki/Block_cipher_modes_of_operation#Counter_.28CTR.29
Private Function blf_Decrypt_CTR(Data_In As StringReader, Out As StringReader)
   
   With Data_In
      
   '  Do Counter init - Get Seed from data
   ' (Note about style: in IC that part is also done inside blf_KeyInit)
      Dim wDataL&, wDataR&
      Dim IV_1&, IV_2&
      
      IV_1 = .int32
      wDataL = DwordReverse(IV_1)
      
      IV_2 = .int32
      wDataR = DwordReverse(IV_2)
         
         'TestDataIn
         ' wDataL 0xFB303510    -80726768
         ' wDataR 0xB7E6371D  -1209649379
      Call blf_EncipherBlock(wDataL, wDataR)
      
      wDataL = DwordReverse(wDataL)
      wDataR = DwordReverse(wDataR)
         'TestdataOut
         ' wDataL 0x78D8AA28   2027465256
         ' wDataR 0x2F220D4C    790760780



    ' Now do Blowfish decryption in CTR mode
      Do Until .EOS = True
       
       ' Decipher data
       ' wDataL
         Out.int32 = .int32 Xor wDataL
         Out.int32 = .int32 Xor wDataR
         
   '     Debug.Print Out, ValuesToHexString(Out)
         
        
        ' increase Counter (Consider IV_1 and IV_2 as UInt64)
          Inc IV_1
         'TODO: if OverFlow - increase IV_2
          If IV_1 = 0 Then
            'TODO Correct OverFlow - increase DataR
             Stop
             Inc IV_2
             If IV_2 = 0 Then Stop
   
          End If
   
         'Gen new encryption vector (= IV -> Initialisation Vector)
         ' 11 35 30 FB 1D 37 E6 B7
         ' 0xFB303511    -80726769
         ' 0xB7E6371D  -1209649379
         wDataL = DwordReverse(IV_1)
         wDataR = DwordReverse(IV_2)
         
         Call blf_EncipherBlock(wDataL, wDataR)
         ' B4 65 93 32 2E 32 A2 A1
         ' 0x329365B4
         ' 0xA1A2322E
        
         wDataL = DwordReverse(wDataL)
         wDataR = DwordReverse(wDataR)
   
      Loop
   
   End With

End Function

Private Function LicDecode_ICLF$(Text$, Key$)
   
   
   Dim Data_In As New StringReader
   Dim Data_Out As New clsStrCat
   
   
   With Data_In
      .Data = Text
      
      .Position = 0
            
      Dim Signature$
      Signature = .FixedString(4)
      If Signature <> "ICLF" Then
         Err.Raise vbObjectError, , "ICLF-Signature invalid. Probaly because decryption failed."
      End If
                  
      Dim VersionData&
      VersionData = .int16
      log_verbose "ICLF-VersionData: " & H16(VersionData)
            
            
      blf_KeyInit cv_BytesFromHex(Key)
      
   End With

   Dim Out As New StringReader
   
   blf_Decrypt_CTR Data_In, Out
            
   
 ' Truncate OutputData
   With Out
      .Position = Data_In.Length - 4 - 2 - (2 * 4) ' because of ICLF,VersionData, IV1,IV2
      .Truncate
      
'      Debug.Print Out.Length, Out, ValuesToHexString(Out)
      
      .Position = 0
      LicDecode_ICLF = .FixedString

      
   End With


End Function

' Translates 'random value' from Mersenne_twister (value range 0..0x3f) into that alphabet:
' "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+/"
' SKIPPED (and adds a "=" at the end)
Private Function GenKey0x40$(seed&)
   
   Dim Data_In As New StringReader
   Dim Data_Out As New clsStrCat
   
   Data_Out.Clear
   random_init seed
   
   Const KEYLEN& = &H3F
   
   Dim DuplicateFinder(0 To KEYLEN) As Boolean
   
   Dim i: i = 0
   Do While i <= KEYLEN
   
      Dim char&
      If seed = 0 Then
         char = i
      Else
         char = Mersenne_twister_random(KEYLEN)
'         Debug.Print i, H8(char), Data_Out.value
      End If
      
      If DuplicateFinder(char) = False Then
         DuplicateFinder(char) = True
      
         If char < 0 Then '..0x00
            'Error Range should be only inside 0..3f
            Stop
            
         ElseIf char < Asc(vbLf) Then   '0x00..0x0a
            char = char + Asc("0")   '0..9
         
         ElseIf char < Asc("$") Then '..0x24
            char = char + Asc("7")   'A..Z
                 
         ElseIf char < Asc(">") Then '..0x3d
            char = char + Asc("=")   'a..z
         
         ElseIf char = Asc(">") Then '0x3e
            char = Asc("+")          '+
            
         ElseIf char = Asc("?") Then '0x3f
            char = Asc("/")          '/
         
         Else
            'should not be reached since
            'char = char and &H3F
            'limits the range from 0..3f
            Stop
            
         End If
   
         Data_Out.Concat Chr(char)
         Inc i
         
      End If
   
   Loop
   
'   Data_Out.Concat "="
   
   GenKey0x40 = Data_Out.value

End Function

Private Function TextToBin$(Text$)
   
   Dim Data_In As New StringReader
   Dim Data_Out As New clsStrCat
   
   With Data_In
      .Data = Text
      
      Do Until .EOS
         
         Dim ByteH&
         ByteH = .int8
         If ByteH > Asc("9") Then
            ByteH = 2 - ByteH  'Asc(" ") / 0x10 -> 2
         End If
         ByteH = ByteH And &HF
         
         
         Dim ByteL&
         ByteL = .int8
         If ByteL > Asc("9") Then
            ByteL = Asc("c") - ByteL
         Else
            ByteL = ByteL - Asc("0")
         End If
         
        'Melt H und L nibbles together
         Data_Out.Concat Chr((ByteH * &H10) Or ByteL)
         
      Loop
   
      TextToBin = Data_Out.value
   
   End With
   
End Function

Private Sub cmd_close_Click()
   Unload Me
End Sub

Private Sub Cmd_Decrypt_Click()
   On Error GoTo Cmd_Decrypt_err
   
   LoadLicenseFile
   
   
'in range of 1023 byte skip '/r' '/n'
' ============= Convert Ioncube Licence File(ICLF) to binary =====================
   Dim Base64Data As New StringReader
   With Base64Data
      .Data = txt_Lic
      
    ' Get first 8 TextBytes (from Base64data) that will become 4 BinBytes -> 1 DWORD
      Dim TextToDword As New StringReader
      With TextToDword
         .Position = 0
         .DisableAutoMove = True
         .Data = TextToBin(Base64Data.FixedString(2 * 4))
         
         Dim seed&
         seed = .int32
         log_verbose "ICLF-Seed: 0x" & H32(seed)
         
      End With
      
    ' That DWORD is the seed to mixup the Base64Alphabet that is used for decoding Base64 to BinData
      Dim Base64Alphabet$  'Note for optimisation 'Base64Alphabet' is only uses(read) once
      Base64Alphabet = GenKey0x40(seed)
      Base64Init Base64Alphabet
      
    ' Now get the Rest of Base64 data and convert it into Bin
      Dim LicData_Bin As New StringReader
      LicData_Bin = Base64DecodeAscii(.FixedString)
      
'      Debug.Print ValuesToHexString(LicData_Bin)
'      Debug.Print LicData_Bin.Data
'      Debug.Print LicData_Bin.Length
      
   End With

' ============= Decrypt Ioncube Licence File(ICLF) =====================
 
   Dim LicData_ICLF As New StringReader
   With LicData_Bin
      
      random_init seed
      
      
      .Position = 0
      .EOS = False
      Do Until .EOS = True
         Dim byteVal As Byte
         byteVal = LicData_Bin.int8
         byteVal = byteVal Xor Mersenne_twister_random(&HFF&)
         LicData_ICLF.int8 = byteVal
      Loop
      
 '     Debug.Print ValuesToHexString(LicData_ICLF)
 '     Debug.Print LicData_ICLF.data
 '     Debug.Print LicData_ICLF.Length
      
      
   End With
   
 ' ============== Decrypt Ioncube Licence Data(ICLD) ==============
   
   Dim BlowFishPassword$ 'ICLF_Password
   BlowFishPassword = GenPassWordHash(Txt_PassPhrase, "&h" & Txt_MemberID)
   
   
   Dim LicData_ICLD As New StringReader
   LicData_ICLD.Data = LicDecode_ICLF(LicData_ICLF.Data, BlowFishPassword)
   
   txt_Lic = LicData_ICLD
   
'   Debug.Print ValuesToHexString(LicData_ICLD)
'   Debug.Print LicData_ICLD.Data
   
   Dim ICLD_PayLoad As New StringReader
   LicValidate_ICLD LicData_ICLD, ICLD_PayLoad
   
 ' Show(and set) important data(FrmMain.SrvRestrictionsItems & ByteCode_Key)
   txt_Lic = LicInterpret_ICLD(ICLD_PayLoad).value
   
Exit Sub
Cmd_Decrypt_err:
   txt_Lic = "ERROR: " & Err.description
   
End Sub
Private Sub LicValidate_ICLD(LicData_ICLD As StringReader, ICLD_PayLoad As StringReader)

 ' Test
   With LicData_ICLD
      .Position = 0
      
      Debug.Assert .Length >= 6
      
      Dim Signature$
      Signature = .FixedString(4) '4
      If Signature <> "ICLD" Then
         Err.Raise vbObjectError, , "ICLD-Signature invalid. Probably because decryption failed due to invalid passphrase or memberID"
      End If
      
      Dim VersionData_1 As Byte, VersionData_2 As Byte
      VersionData_1 = .int8 '5
      VersionData_2 = .int8 '6
      log_verbose "ICLD-VersionData: " & H8(VersionData_1) & " " & H8(VersionData_2)
      
      
      
      ICLD_DataEntries = .int32 'a
      log_verbose "ICLD_DataEntries: " & H32(ICLD_DataEntries)
      
      
      If (ICLD_DataEntries >= 4) Then
         If VersionData_1 = 2 Then
            'okay
         ElseIf VersionData_1 < 2 Then
'            myStop
            '8
            'The encoded file %s requires a license file.
            'The license file %s is an old version.
            'Please request a new license file from the PHP application provider.
         
         Else
'            myStop
            '9
             'The license file %s
             'for the encoded file %s
             'requires an updated Loader. Please install the latest Loader
             'From loaders.ioncube.com
         End If
      End If
      
      
      
      Dim ICLD_PayLoad_SHA256 As New StringReader
      ICLD_PayLoad_SHA256 = .FixedString(&H20) '2a
      
      Dim ICLD_PayLoadSize
      ICLD_PayLoadSize = .Length - .Position
      
'      Dim ICLD_PayLoad As New StringReader
      ICLD_PayLoad.Data = .FixedString
      
   End With
   
 ' Validating Hash
   Dim ICLD_PayLoad_SHA256_Generated$
   ICLD_PayLoad_SHA256_Generated = SHA256(ICLD_PayLoad.Data)
   ICLD_PayLoad_SHA256_Generated = HexStringToString(ICLD_PayLoad_SHA256_Generated)
   If ICLD_PayLoad_SHA256_Generated <> ICLD_PayLoad_SHA256 Then
      Dim DbgFileName$
      DbgFileName = "ICLD_Body_ThatShouldBe_SHA256_" & ValuesToHexString(ICLD_PayLoad_SHA256, "") & ".bin"
      
      log "ICLD_PayLoad SHA256 does not match!"
      log "Dumping data for debugging to: '" & DbgFileName & "'"
      FileSave DbgFileName, ICLD_PayLoad.Data
      
   Else
      log_verbose "ICLD-Data SHA-256(validated): " & ValuesToHexString(ICLD_PayLoad_SHA256)
      
   End If

End Sub
   
Private Function LicInterpret_ICLD(ICLD_PayLoad As StringReader) As clsStrCat
   
   Set LicInterpret_ICLD = New clsStrCat
   
   With ICLD_PayLoad
      Dim ICLD_DataEntry&
      For ICLD_DataEntry = 1 To ICLD_DataEntries

         Dim KeyName$
         KeyName = GetEncStr(ICLD_PayLoad)
         
         Dim Enforcesflag$
         
         Dim value As New StringReader
         value = GetEncStr(ICLD_PayLoad, Enforcesflag)
         
         
         Select Case KeyName
            Case "__expiry"
            
               Dim expireDate&
               expireDate = value + &H500DA46  '83941958
               value = FormatCTime(expireDate) & "   0x" & value & " '+0x500DA46'=> 0x" & H32(expireDate)
            
            Case "__member_id"
            
            
            Case "__mykey"
               ByteCode_Key = value
               value = value & " -> 0x" & H32(value)
               
            
            Case "__preamble_hash"
               
               Dim preamble_hash_Generated$
               preamble_hash_Generated = LCase(SHA256(preamble_data))
               
               Dim preamble_hash$
               preamble_hash = LCase(ValuesToHexString(value, ""))
'               value = preamble_hash
               
               If preamble_hash = preamble_hash_Generated Then
                  value = "SHA-256(matched): " & preamble_hash
               Else
                  value = "SHA-256(INVALID): " & preamble_hash
               End If
               
            
            Case Else
               If KeyName Like "__restriction_binary_*" Then
               
                ' important setting for morfkey
                ' (__mykey * &h92492493  'imul;  'Result_High64Part + __mykey  'add edx,__mykey; 'Result \ 4  'sar edx,2....)
                  FrmMain.SrvRestrictionsItems = Inc(FrmMain.SrvRestrictionsItems)
                  
                  Dim SRestr As New clsStrCat
                  
                  FrmMain.HandleServerRestrictions value, SRestr

                  value.Data = SRestr.value ' ValuesToHexString(Value)
                  SRestr.Clear
                  
                ElseIf KeyName Like "__*" Then
                
                Else
                  'Deserialise Value
               
                End If
            
            
         End Select
         
         Dim outputLine
         outputLine = KeyName & " = [" & Enforcesflag & "]" & value
         log "LICFILE " & outputLine
         
         LicInterpret_ICLD.Concat outputLine & vbCrLf
         

         
      Next

   End With
   
End Function

Private Sub removeWhitespaces(ByRef InOutData$)
   ReplaceDoMulti InOutData, " ", ""
   ReplaceDoMulti InOutData, vbCr, ""
   ReplaceDoMulti InOutData, vbLf, ""
   ReplaceDoMulti InOutData, vbTab, ""
   
End Sub


Private Sub LoadLicenseFile()

   Dim LicFileData$
   LicFileData = FileLoad(txt_LicFile)

   Const LicDataStartPattern$ = "------ LICENSE FILE DATA -------"
   Const LicDataEndPattern$ = "--------------------------------"
   
   Dim LicDataSplited
   LicDataSplited = Split(LicFileData, LicDataStartPattern)
   
 ' Note preamble_data is the data before "------ LICENSE FILE DATA -------"
   preamble_data = LicDataSplited(0)
   removeWhitespaces preamble_data
   
 ' Get Base64 LicData
   LicFileData = LicDataSplited(1)
   
 ' split off end Licdata Terminator("--------------------------------")
   LicFileData = Split(LicFileData, LicDataEndPattern)(0)
   
   removeWhitespaces LicFileData
   
   txt_Lic = LicFileData
   
End Sub
   
   

Private Sub cmdDecode_Click()
On Error GoTo cmdDecode_err
   
   Dim MemberID&
   MemberID = "&h" & Txt_MemberID
   
   WriteMemberIDIntoMakeLic MemberID
   
   txt_Lic = ""
   
   Dim MakeLic_Args()
   MakeLic_Args = Array( _
      "--decode-license", """" & txt_LicFile & """", _
      "--passphrase", """" & Txt_PassPhrase & """")
   
   Console.ShellExConsole _
      App.Path & "\" & MakeLic_File, _
      Join(MakeLic_Args)

Exit Sub
cmdDecode_err:
MsgBox Err.description, , "ERRoR"
   
End Sub

Private Sub WriteMemberIDIntoMakeLic(MemberID&)
   
   Dim MakeLicExe As New FileStream
   With MakeLicExe
      .Create App.Path & "\" & MakeLic_File, False, False, False
     
     ' seek to PE-Header
      .Position = &H3C
      .Position = .longValue
      
    ' Test Timestamp
      .Move 8
      If .longValue <> MakeLic_PE_BuildTimeStamp Then
         Err.Raise vbObjectError, , MakeLic_File & " unsupported version!"
      End If
      
    ' Write ID into file
      .Position = MakeLic_PatchOffset
      .longValue = MemberID
      
      .CloseFile
   
   End With
   
End Sub


Private Sub Console_OnOutput(TextLine As String, ProgramName As String)
 ' Bad Performance !!!
   txt_Lic = txt_Lic & TextLine
End Sub

Private Sub Form_Initialize()
   Set Console = New Console
End Sub

Private Sub Txt_MemberID_Validate(Cancel As Boolean)
   On Error Resume Next
   Txt_MemberID = H32("&h" & Txt_MemberID)
   If Err Then Txt_MemberID = "00000000"
End Sub
