VERSION 5.00
Object = "{E88121A0-9FA9-11CF-9D9F-00AA003A3AA3}#1.0#0"; "ZlibTool.ocx"
Begin VB.Form FrmMain 
   Caption         =   "My IronCube Decoder"
   ClientHeight    =   8565
   ClientLeft      =   1530
   ClientTop       =   675
   ClientWidth     =   9330
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   9330
   Begin VB.CommandButton cmd_DecodeLic 
      Caption         =   "Show License Decoder"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.CheckBox Chk_verbose 
      Caption         =   "Verbose LogOutput"
      Height          =   195
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Aktiviert
      Width           =   1800
   End
   Begin VB.ListBox ListLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   3
      ToolTipText     =   "Double click to force saving to ic.log !"
      Top             =   1200
      Width           =   9135
   End
   Begin VB.Timer Timer_OleDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   2160
   End
   Begin VB.TextBox Txt_Filename 
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   1
      Text            =   "Drag the Php script that was compiled/encoded with IronCube Encoder in here, or enter/paste path+filename."
      ToolTipText     =   "Drag in or type in da file"
      Top             =   360
      Width           =   9135
   End
   Begin ZLIBTOOLLib.ZlibTool ZlibTool 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Decompression status bar"
      Top             =   1080
      Width           =   9135
      _Version        =   65536
      _ExtentX        =   16113
      _ExtentY        =   238
      _StockProps     =   0
   End
   Begin VB.TextBox Txt_Script 
      Height          =   615
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "RawData"
      Top             =   5880
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.ListBox List_Source 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      ItemData        =   "FrmMain.frx":1CFA
      Left            =   120
      List            =   "FrmMain.frx":1CFC
      TabIndex        =   2
      Top             =   6360
      Width           =   9135
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SkipExtract As Boolean = False 'True
Const BYTECODEKEY_INIT_246& = &H363432
Const BYTECODEKEY_DELETED& = &H92A764C5

'Mersenne Twister
Private Declare Function MT_Init Lib "MT.DLL" (ByVal initSeed As Long) As Long
Private Declare Function MT_GetI8 Lib "MT.DLL" () As Long


Dim File As New FileStream

Dim DisableDeleteTmpFile As Boolean

Dim FilePath_for_Txt$
Dim Body_MT_Seed&

Dim IncludeXorKey&

Public SrvRestrictionsItems  As Byte

''Bugfix for ZlibTool.ocx (ZlibTool uses CreateWindowEx(...Class = "msctls_progress32"...) but forget to call that init)
'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Const IC_START_PATTERN$ = "<?php //0"
Const Digits& = 4

Dim PHP_Loader_Size&
      
'Private Sub Form_Initialize()
'   InitCommonControls
'End Sub

Function MidMbcs(ByVal str As String, Start, Length)
    MidMbcs = StrConv(MidB(StrConv(str, vbFromUnicode), Start, Length), vbUnicode)
End Function


Sub FL_verbose(Text)
   log_verbose H32(File.Position) & " -> " & Text
End Sub

Sub log_verbose(TextLine$)
   If Chk_verbose.value = vbChecked Then log TextLine
End Sub



Sub FL(Text)
   log H32(File.Position - 1) & " -> " & Text
End Sub

Public Sub LogSub(TextLine$)
   log "  " & TextLine
End Sub

Public Sub LogDecompiled(TextLine$)
   Txt_Script = Txt_Script & TextLine
End Sub



Public Sub log2(TextLine$)
'   log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Public Sub log(TextLine$)
On Error Resume Next
   Dim line
   For Each line In Split(TextLine, vbCrLf)
      ListLog.AddItem line
   Next
'   ListLog.AddItem H32(GetTickCount) & vbTab & TextLine
 
 ' Process windows messages (=Refresh display)
   If (ListLog.ListCount < 30000) Or (Rnd < 0.1) Then
       ' Scroll to last item ; when there are more than &h7fff items there will be an overflow error
      If (ListLog.ListCount >= 0) Then
      Dim ListCount&
         ListLog.ListIndex = ListLog.ListCount - 1
         DoEvents
      End If
   End If
   
   Err.Clear
   
End Sub

Public Function GetLogdata$(Listbox As Listbox)
   With Listbox
      Dim LogData As New clsStrCat
      LogData.Clear
      Dim i
      For i = 0 To IIf(.ListCount < 0, 32767, .ListCount)
         LogData.Concat .List(i) & vbCrLf
      Next
      
      GetLogdata = LogData.value
      
   End With
End Function


'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Public Sub log_clear()
On Error Resume Next
   ListLog.Clear
End Sub




Private Sub Init()
   

   icKnownFunctions = Split(FileLoad(App.Path & "\data\" & _
                              "php_functions_list.dat"), _
                            vbCrLf)

   icByteCodeNames = Split(FileLoad(App.Path & "\data\" & _
                              "ByteCodeNames_list.dat"), _
                            vbCrLf)
 
 
   op_extended_value = Array( _
                           "", _
                           "ZEND_DECLARE_CLASS", _
                           "ZEND_DECLARE_FUNCTION", _
                           "ZEND_DECLARE_INHERITED_CLASS," _
                       )

   
End Sub



Private Sub cmd_DecodeLic_Click()
   FrmDeCodeLicense.Show
End Sub

Private Sub Form_Load()
   
   
'   Dim value&
'   random_init 19
'   value = Mersenne_twister_random(&HFFFF)
'
'   Dim Value2&
'   MT_Init 19
'   MsgBox H32(value)
'   Value2 = MT_GetI8
'   MsgBox H32(value) & " - " & H32(Value2)
      
   
   FrmMain.Caption = FrmMain.Caption & " " & App.Major & "." & App.Minor & " build(" & App.Revision & ")"
   
   'Extent Listbox width
   Listbox_SetHorizontalExtent ListLog, 6000
   
   
   Init
'   frmCodeGen.Show vbModal
'   End
   
   
   
   Dim commandline As New commandline
   If commandline.NumberOfCommandLineArgs Then
      FilePath_for_Txt = commandline.getArgs()(0)
      Timer_OleDrag.Enabled = True
   End If
   
   
End Sub

Public Sub LogStage(Stage&, Title$)
   log vbCrLf & _
       "=========> S t a g e  " & Stage & "  -  " & Title & _
       vbCrLf
End Sub


Private Sub DecodePhp(FileName$)
   
   log_clear
   List_Source.Clear
   ArrayDelete Dumper_Data
   
   Dim objFileName As New ClsFilename
   objFileName.FileName = FileName
   ChDrive objFileName.Path
   ChDir objFileName.Path
   
   With App
      log "========== " & App.Title & " " & Join(Array(App.Major, App.Minor, App.Revision), ".") & "  LogFile =========="
   End With
   
LogStage 0, "Getting RawData (UnBase64-Raw.bin)"
   
   With File
   log "Opening " & FileName
      .Create FileName, False, False, True
     ' .Create "D:\PHP\ioncube-encoded-file.php", False, False, True
   
      Dim PHP_HeaderSig$
      PHP_HeaderSig = .FixedString(14)
      
      On Error GoTo Err_NotA_IC_File
      
'      Const IC_START_PATTERN$ = "<?php //0"
'      Const Digits& = 4
      
'      Dim PHP_Loader_Size&
      PHP_Loader_Size = "&h" & Split(PHP_HeaderSig, IC_START_PATTERN)(1)
      
      On Error GoTo Err_DecodePhp
      
      

      
      .Position = PHP_Loader_Size + Digits

'---  added Code (correct PHP_Loader_Size it is too big)
      'Ensure that data starts at NewLine
      .Move -1
      'Note: WindowsLineBreak is 0x0d 0x0a CRLF while
      '        LinuxLineBreak is      0x0a   LF
      If .char <> vbLf Then
         log "That file as an incorrect header length and is probably not loaded by the original IC-Loader."
         'Ensure the following data is just Base64
         ' I assume this this doesn't contain any space
         If .FindString(" ") <> -1 Then
            log "Skipping correcting PHP_Loader_Size since it seem that PHP_Loader_Size is to small - what is tolerate by this and the original IC-Loader."
'            Stop
         Else
            If .FindStringRev("?>" & vbLf) = -1 Then
               log "Whoops there was no linebreak(0x0A LineFineChar) found before the IC-Base64 data"
               Stop
            Else
'               .Move 1
               log "Corrected PHP_Loader_Size from " & H16(PHP_Loader_Size) & " -> " & _
                   H16(.Position - Digits) & "."
            End If
         End If
      End If
'----
      
     'Read Base64-Data
      Dim PHP_Data As New StringReader
      With PHP_Data
         
         FL_verbose "Start of ionCubeData..."
         .Data = File.FixedString(-1)
         
         FL_verbose H32(.Length) & " Bytes read."
         Txt_Script = .Data
      
         .DisableAutoMove = True

        'VersionCheck
         Dim PHP_Ver&
         PHP_Ver = .int32

         Dim bFound&
         bFound = False
         
         Dim encrypted As New StringReader
         Select Case PHP_Ver
            Case &HDEADC0DE, &H3FBC2883, &H882BC103, &H217582F, &H149FEC13, &H67A6BF45, &H9EB67AC2
               log "IC_Binary_Marker: " & H32(PHP_Ver)
             ' Uncommented because '.DisableAutoMove = True'
             ' .Move -4
               
               bFound = True
               
               Dim Base64ToRaw_SizeShrink&
               encrypted.Data = .FixedString
               Base64ToRaw_SizeShrink = 0

               

            Case Else

               Do
                  Dim PHP_Ver_Str$
                  PHP_Ver_Str = .FixedString(4)
                  Select Case PHP_Ver_Str
                     Case "0y4h", "BrWN", "4+oV", "HR+c", "mdgs"
                        log "IC_Base64_Marker: '" & PHP_Ver_Str & "' Decoding Base64->Bin..."

                       ' Uncommented because '.DisableAutoMove = True'
                       ' .Move -4
                        bFound = True
                        log_verbose ""
                        
                        Base64Init
                        encrypted.Data = Base64DecodeUnicode(.FixedString)
                        Base64ToRaw_SizeShrink = (PHP_Data.Length) - (encrypted.Length - 2)


                     Case Else
                        .Move 1
                  End Select
               Loop Until bFound Or (.Position >= 64)

               If bFound Then
               'okay
               Else
                ' Error
                  myStop
               End If

         End Select

         .DisableAutoMove = False

         
         'DebugOutput
         'And this is the reason
         'well fix it...
         FileSave "UnBase64-Raw.bin", encrypted.Data
'         log_verbose "Data saved to 'UnBase64-Raw.bin'"

         
LogStage 1, "Validating RawData"
      
      Dim ver_enc&
      ver_enc = encrypted.int32
      
      Dim ver&
      ver = ver_enc Xor &H2853CEF2
      log_verbose "ionCubeVersionType: " & H32(ver) & " [Encrypted: " & H32(ver_enc) & "]"
      

      Dim NewFileSize&
      NewFileSize = File.Length - Base64ToRaw_SizeShrink
      
      Dim ExitVal
      
      Select Case ver
         Case &H17EFE671 'v1
            'Dec_v1
            log_verbose ("Oops Unsought_Version_Error: Dec_v1 is not implemented")
            myStop
         
         Case &H2A4496DD 'v2
            log_verbose ("ioncube_Version_Type 2")
            log_verbose ("Warning: this ioncube_Version_Type was not Tested.")
            Dec_v2 encrypted, 0, 0, 1, NewFileSize, encrypted
            
         Case &H3CCC22E1 'v3
            log_verbose ("ioncube_Version_Type 3")
            Dec_v2 encrypted, 0, 1, 1, NewFileSize, encrypted
         
         Case &H4FF571B7, &HB6E5B430 'v4,v6
             log_verbose ("ioncube_Version_Type 4,6 are not supported by 'ioncube_loader_win_4.3.dll'")
             ExitVal = -1
     
         Case &HA0780FF1 'V5 binary
            log_verbose ("ioncube_Version_Type 5")
            log_verbose ("Warning: this ioncube_Version_Type was not Tested.")
            Dec_v2 encrypted, 0, 0, 0, NewFileSize, encrypted
         
         Case &HF6FE0E2C
            'Dec_v7
            log_verbose ("Oops Unsought_Version_Error: Dec_v7 is not implemented yet")
            myStop
            
         Case Else
            log_verbose ("Error: Unknow ioncube_Version_Type")
            ExitVal = 0
            myStop
      End Select
      
      
      End With
      
      Dim PHP_FileLen&
      PHP_FileLen = .Length
      
      
      .CloseFile
   End With

Err_DecodePhp:
   Select Case Err
      Case 0
      Case Else
         log "-> ERROR: " & Err.description
   End Select

' SAVE Data
   SaveSource
   SaveLog
   
Exit Sub

Err_NotA_IC_File:
   log "-> ERROR: This is not a ionCube Encode Php File." & vbCrLf & _
       "   Extentend Error Info: IonCube Start Pattern: '" & IC_START_PATTERN & "' not found!"
   File.CloseFile
   
End Sub

Private Sub Dec_v2(Data As StringReader, f1, f2, f3, RawBinData_FileSize, encrypted As StringReader)
   
LogStage 2, "Processing RawData"
   
   log_verbose "RawBinData_Size: " & H32(RawBinData_FileSize)
   
   Dim FileSizeKey&
   FileSizeKey = (RawBinData_FileSize + 12321) Xor &H23958CDE
   log_verbose "FileSizeKey: " & H32(FileSizeKey) & " (RawBinData_Size + 12321) Xor &H23958CDE"
   
'but whoops should that be 1212 (it's 1213)
'okay mission is now to see why RawBinData_FileSize is 1213 an not 1212

'but here it is -124 so I go back tpo see what happend
'that is a checkpoint
'let's see if data have differed in the 'original'
'but first how to locate this in the asm code?
'Well that const '23958CDE' is very unique and also asm-compatible
'okay got the mission so far?
'k
'nice
   
   With encrypted
   
     '==============================================
     ' Process Header
     ' (Note: Header starts at pos 0x4; @Pos 0x0 is the IC-Version)
     

      Dim Header_Raw As New StringReader
      Header_Raw.Data = .FixedString(&HC + &HC)
      
      
      Dim Header_Bin As New StringReader
      Replace3C_TagStart Header_Raw, &HC, Header_Bin
      With Header_Bin
         .EOS = False
         Dim Header_FileSizeKey&, HeaderSize&, HeaderKey&
         Header_FileSizeKey = .int32 '+0
         HeaderSize = .int32 '+4
         HeaderKey = .int32 '+8
         HeaderSize = AddInt32((HeaderSize Xor 407893395), -203515694) Xor HeaderKey
                                          'xor 0x184FF593   -0x0C21672E
                                          
         'Rest +0C, +10 and +14 are FillData
         ' that will be used more or less in 'real' life
         ' ( for the theoretical case that the Header is fill with only '<'
         '   this FillData will used entirely (since '<' is encode as 0xFF 0x80..0xFF)
         
      End With
      
    
      Header_FileSizeKey = Header_FileSizeKey Xor HeaderKey
      log_verbose "Header_FileSizeKey: " & H32(Header_FileSizeKey)
    ' Test/Validate Key
      If FileSizeKey <> Header_FileSizeKey Then
         log "Header_FileSizeKey and FileSizeKey are different! (Difference: " & H32(Abs(Header_FileSizeKey - FileSizeKey)) & ")"
         
         Dim Header_FileSizeKey_Size&
         Header_FileSizeKey_Size = (Header_FileSizeKey - 12321) Xor &H23958CDE
         log "Note: RawBinData_Size calculate from Header_FileSizeKey is: " & H32(Header_FileSizeKey_Size) & "  (Difference: " & H32(Abs(Header_FileSizeKey_Size - RawBinData_FileSize)) & ")"
      End If
      
'      MT_Init HeaderKey
      random_init HeaderKey
      log_verbose "Header_Size: " & H32(HeaderSize) & " (InitKey for MT: " & H32(HeaderKey) & ")"
      
      If (.Position + HeaderSize + 8) > (.Length - 2) Then
         Err.Raise vbObjectError, , "HeaderSize or Header is corrupt"
      End If
      

     '==============================================
     ' Process FirstChunk
      
     'Get Header & MD5_CheckSum data
      Dim Header As New StringReader, MD5_CheckSum As New StringReader
      With Header
         .Data = ReadHeaderData(encrypted, HeaderSize)
         Debug.Assert Len(.Data) <> 0
         
         .Position = (HeaderSize - &H10)
         MD5_CheckSum.Data = .FixedString(&H10)
      
      
      'ChunkSize=0x40?
      End With
      
     'Test CRC of Header
      Dim CRC_Start&
      CRC_Start = 4
      
      Dim CRC_End&
      CRC_End = .Position
      
      'Create ADLER32 from ICDataPart+4..ICDataPart.EOS
      '<?php //0
      '...
      '?>
      
      .StorePos
      
      
         Dim Adler32Data As New StringReader
         
         .Position = CRC_Start
   
         Adler32Data.Position = 0
         Adler32Data = .FixedString(CRC_End - CRC_Start)
         
         Dim Adler32_generated$
         Adler32_generated = IC_ADLER32(Adler32Data)
      
      .RestorePos
      
     'Read CRC from Header
      Dim CRC_Raw As New StringReader
      CRC_Raw.Data = .FixedString(4 + 4)
      
      Dim CRC As New StringReader
      Replace3C_TagStart CRC_Raw, 4, CRC
      
      Dim Adler32$
      Adler32 = H32(CRC.int32)
      
     'Compare it with generated
      If Adler32 <> Adler32_generated Then
         log_verbose "Adler32_CRC for 'IC_Raw[" & H16(CRC_Start) & "]..[" & H16(CRC_End) & "]' and calculated  DOES NOT MATCH!"
         log_verbose "  from file : " & Adler32
         log_verbose "  calculated: " & Adler32_generated
      Else
         log_verbose "Adler32_CRC for 'IC_Raw[" & H16(CRC_Start) & "]..[" & H16(CRC_End) & "]' and calculated MATCH. CRC: " & Adler32
      End If
   
      
'      log "Adler32(CRC) of ICDataPart+4..ICDataPart.EOS: " & H32(CRC.int32)
            
      'Decrypt ChunkKey(=MD5_CheckSum data)
      ' rotate >> each byte 3bits right.
      With MD5_CheckSum
         .DisableAutoMoveOnRead = True
         .EOS = False
         Do While .EOS = False
            .int8 = ror(.int8, 3)
         Loop
         .DisableAutoMoveOnRead = False
      End With
      
      
      
      'Decrypt Header
LogStage 3, "Extracting Header from RawData (->IC_Header.bin)"
      
      With Header
         .DisableAutoMoveOnRead = True
         MD5_CheckSum.EOS = False
   
         .EOS = False
         Do While .EOS = False
         
            'Xor current byte with byte from MT...
            .int8 = .int8 Xor Mersenne_twister_random(&HFF) Xor MD5_CheckSum.int8
           'Cycle through MD5_CheckSum_Bytes
            If MD5_CheckSum.EOS = True Then MD5_CheckSum.EOS = False
            
         Loop
         
            FileSave "IC_Header.bin", Header.Data
      End With
      
      
LogStage 4, "Interpreting IC_Header"
      
      'ReadChunk
      With Header
         .DisableAutoMoveOnRead = False
         .Position = 0
         
         
         If f3 Then
            Dim Header_v1&
            Header_v1 = .int32
            If Header_v1 > 3 Then
             myStop 'ret -1
            End If
            
            Dim MinimumLoaderVersion&
            MinimumLoaderVersion = .int32
            log "Minimum Loader Version: " & Format(MinimumLoaderVersion, "00\.00\.00") & " (for ex. ioncube_loader_win_4.3.dll requires >0301010)"
'            If MinimumLoaderVersion > 301010 Then
'             myStop 'ret -1
'            End If
            
            log_verbose "VerData" & " 0x" & H32(Header_v1)
            
            Dim ObfuFlags&
            ObfuFlags = .int32
            
            Dim v4_Dummy&
            v4_Dummy = .int32
            
            log_verbose "ObfuFlags " & H32(ObfuFlags) & " " & H32(v4_Dummy) & vbCrLf & _
                        ShowFlags(ObfuFlags, _
                                  "Obfuscate Vars", "Obfuscate Funcs", "0004 ", "0008")

            
         
            Dim Header_CustomLoaderEventMessagesCount&
            Header_CustomLoaderEventMessagesCount = .int32
         
            Dim ObfuHashSeed As New StringReader
            ObfuHashSeed.Data = .FixedString(Header_CustomLoaderEventMessagesCount)
            log "ObfuFuncHashSeed: " & ValuesToHexString(ObfuHashSeed)
            
            'Exec
         
            Bytecode_MT_InitKey = .int32
            
            log "Bytecode_XorKey: " & H32(Bytecode_MT_InitKey)
            If Bytecode_MT_InitKey = BYTECODEKEY_DELETED Then 'Note BYTECODEKEY_DELETED = 0x92A764C5
               log " -> The Encoder uses '0x" & H32(BYTECODEKEY_DELETED) & "' to overwrite/delete the real bytecode key - since there is a LicenseFile(+enforceLicense) from which that value should be retrieved. (However don't worry we'll be able to calcuate that value if there is not LicFile. :)"
            End If
            '... just some profiling:
            '0x92A764C5 -> 2460443845 -> 2460 44 38 45
            '                            ^^^^  D  8  E
            'I wonder what that '246' is - since it's also in the BYTECODEKEY_INIT_246?
            '... on the num keypad 2-4-6  gives a Triangle
            '.... and 2-4-6-0 forms a 'y'
            
            
         Else
            Bytecode_MT_InitKey = BYTECODEKEY_INIT_246 ' &H363432 -> "246."
         End If
         

         IncludeXorKey = .int32
         log "IncludeXorKey[should be 0xE9FC23B1]: " & H32(IncludeXorKey)
         
         If f2 Then
         
            'Read a Struct..
            Dim Header_NumItem&
            Header_NumItem = .int32
            
            '1  EBX+C and EBX+1c
            Dim LicenseFile$
            LicenseFile = LoadLicString(Header)
            If LicenseFile <> "" Then
              log "LicenseFile: " & LicenseFile
' Fill in 'txt_LicFile'
  FrmDeCodeLicense.txt_LicFile = LicenseFile

              
              
            End If
            
            Dim DisableCheckingOfLicenseRestrictions&
            DisableCheckingOfLicenseRestrictions = loadValue(Header)
            log "DisableCheckingOfLicenseRestrictions: " & DisableCheckingOfLicenseRestrictions
            
            
            '2
            Dim LicensePassphrase$
            LicensePassphrase = LoadLicString(Header)
            If LicensePassphrase <> "" Then
               log "LicensePassphrase: " & LicensePassphrase
' Fill in 'LicensePassphrase'
  FrmDeCodeLicense.Txt_PassPhrase = LicensePassphrase
            End If
            
            Dim CustomErrCallbackFile$
            CustomErrCallbackFile = loadString(Header)
            If CustomErrCallbackFile <> "" Then
               log "CustomErrCallbackFile: '" & CustomErrCallbackFile & "'"
            End If
            
            '3 EBX+30  ESP
            Dim CustomErrCallbackHandler$, Enable_auto_prepend_Append_file&
            CustomErrCallbackHandler = LoadLicString(Header)
            log "CustomErrCallbackHandler: '" & CustomErrCallbackHandler & "'"
            
            
            Enable_auto_prepend_Append_file = loadValue(Header)
            log "Enable_auto_prepend_Append_file: " & Enable_auto_prepend_Append_file
            
            'log "LoadData(6): " & Join(Array(LicenseFile, DisableCheckingOfLicenseRestrictions, LicensePassphrase, CustomErrCallbackFile, CustomErrCallbackHandler, Enable_auto_prepend_Append_file), "  ")
            
            Dim i&
            For i = Header_NumItem To 6 - 1
               
               'Read 2 dummy int32 + ??? bytes
               myStop 'Untested
               
               .Move (.int8 + 8)
   
            Next
         End If
                  

         HandleCustomErrorMessages Header

         
'         Header_FileSizeKey = Header_FileSizeKey Xor HeaderKey
         
         HandleIncludeRestrictions Header
         
         
         
         SrvRestrictionsItems = .int8

         Dim SrvRestrictionsItem
         For SrvRestrictionsItem = 1 To SrvRestrictionsItems
      
            Dim SRestr As New clsStrCat
            
            Dim RestrictionRows&
            RestrictionRows = .int8
            
            log "Server restrictions entries: 0x" & H8(RestrictionRows)
      
            Dim RS_row
            For RS_row = 1 To RestrictionRows
               
               SRestr.Concat " #" & RS_row & " "
               HandleServerRestrictions Header, SRestr
            
               log SRestr.value
               SRestr.Clear
         
            Next 'rows
   
         Next 'restrictionItems
         
         
'         Debug.Assert FileSizeKey = Header_FileSizeKey
         
         'Create ADLER32 from PHPLoaderData
         '<?php //0
         '...
         '?>
         'Dim Adler32Data As New StringReader
         File.Position = 0
         Adler32Data = File.FixedString(PHP_Loader_Size + Digits)
         
         'Dim Adler32_generated$
         Adler32_generated = IC_ADLER32(Adler32Data)
         
         
         'Dim Adler32&
         Adler32 = H32(.int32)
        'Compare it with generated
         If Adler32 <> Adler32_generated Then
            log_verbose "Adler32_CRC from Header ('<?php //... ?>') DOES NOT MATCH!"
            log_verbose "  from file : " & Adler32
            log_verbose "  calculated: " & Adler32_generated
         Else
            log_verbose "Adler32_CRC for '<?php //... ?>' and calculated MATCH. CRC: " & Adler32
         End If
         
         
'         Morf_ByteCodeKey
         
         
         log_verbose "IC_HeaderEx start: " & H16(Header.Position)
         
         Dim Header_MoreData As New StringReader
         Header_MoreData.Data = .FixedString(&H28)
         
         log_verbose "IC_HeaderEx end: " & H16(Header.Position) & "   IC_Header HeaderSize: " & H16(Header.Length)
         
         '$ ==>    00040001  ..
         '$+4      00000CA7  §...
         '$+8      09050603  .
         '$+C      000029DD  Ý)..
         '$+10     E9FC23B1  ±#üé
         '$+14     6500A8C0  À¨.e
         '$+18     B1870D00  ..‡±
         '$+1C     00 00 5AC5  ÅZ..
         '$+20     096CD432  2Ôl.
         '$+24     41749CF3  óœtA
         
'         If LicenseFile <> "" Then myStop 'Branch not implemented
         
         log_verbose ""

LogStage 5, "Interpreting IC_HeaderEx"

         With Header_MoreData
            Dim IC_Type_Major&, IC_Type_Minor&, IC_Type_Build&
            IC_Type_Minor = .int16
            log_verbose H16(IC_Type_Minor) & " <-IC_Type_Minor"
            Debug.Assert IC_Type_Minor = 1
            
            
            '"[The file %s was encoded with the ionCube PHP 5 Encoder,and requires PHP 5 to be installed."

            IC_Type_Major = .int16
            log_verbose H16(IC_Type_Major) & " <-IC_Type_Major"
            
            If IC_Type_Major >= 5 Then
               'Err.Raise vbObjectError, ,
               log_verbose "The file [%s] was encoded with the PHP 5 ionCube Encoder, and requires PHP 5 to be installed."
            End If
            

            
            
            Dim PhpFlags&
            PhpFlags = .int32
            log_verbose H32(PhpFlags) & " <-PhpFlags"
            'PhpFlags & 0x20 -> CheckForTrusted Provider "<br><b>%s</b> cannot be processed because an untrusted PHP zend engine extension is installed. <a href="http://www.ioncube.com/untrusted_extensions.php">Read more about this message</a>"
            log ShowFlags(PhpFlags, _
                           "0001", "0002 ", "0004 ", "0008", _
                           "0010", "Allow run with untrusted server extensions", "PHP5_Body", "CmdBytesAreEncrypted", _
 _
                           "0100", "ObfuscateFuncNames", "EncryptStrings[KeyWord.c_Object] in memory during Compile& Execute", "ObfuscateStripLineNumbers", _
                           "ObfuscateVarNames", "some encryflag(used in the middle of 'Interpre')", "4000", "8000") ', _
 _
                           "1 0000", "2 0000", "4 0000", "8 0000", _
                           "10 0000", "20 0000", "40 0000", "80 0000", _
 _
                           "100 0000", "200 0000", "400 0000", "800 0000", _
                           "1000 0000", "2000 0000", "4000 0000", "8000 0000")

            
'67
            Dim IC_SerialData&, IC_MemberID&
            IC_SerialData = .int32
'1
            IC_MemberID = .int32 '0xD431 -> 54321
            log "IC Encoder Registration data(MemberID): 0x" & H32(IC_SerialData) & "  " & IC_MemberID
            
      
          ' Fill in 'Txt_MemberID'
            FrmDeCodeLicense.Txt_MemberID = H32(IC_MemberID)
         
            ' Do there is a LicenseFile and LicenseRestrictions are enforced
            If (LicenseFile <> "") And (DisableCheckingOfLicenseRestrictions <> 1) Then
                  
              ' Set Bytecode_MT_InitKey to some good init value - for the case the licfile is missing or something went wrong
                FrmDeCodeLicense.ByteCode_Key = BYTECODEKEY_INIT_246 + SumOfTheCharsFrom(LicensePassphrase)
                log "Calculate ByteCode_Key ('0x363432' + sum of the chars from LicensePassphrase): " & H32(FrmDeCodeLicense.ByteCode_Key)
              
              ' Click on 'Decrypt'
                FrmDeCodeLicense.Cmd_Decrypt.value = True
            
                Bytecode_MT_InitKey = FrmDeCodeLicense.ByteCode_Key
                log "Overwrite ByteCode_Key to : 0x" & H32(Bytecode_MT_InitKey)
         '   Else
               ' Bytecode_MT_InitKey = BYTECODEKEY_INIT_246 + (EvalTimeMin And &HFFFF&) - 1
               ' ... -1 or -2 or -0 seconds? Well that'll be the question/problem with derivating the key from evalTimeMin
               ' (but please don't let me confuse you - you won't need that - however since I just started writing about it, I'll place it here ;)
               
               'Okay first how the Encoder generates these values
               '   EvalTimeMin  <= _time (Gets actual CTime GMT +0000)
               '   EvalTimeMax  <= ExactBuildTime + DaysUntilExpiry
               '   ExactBuildTime <= _time() is called before EvalTimeMin
               '
               '   Bytecode_MT_InitKey <= 363432 + ( ExactBuildTime and 0xFFFF)
               '
               'The problem EvalTimeMin might be not identical to ExactBuildTime.
               'Since it calls _time two times. When php files are encode on a slow computer
               'the EvalTimeMin might get a little bigger than ExactBuildTime.
               '
               'Assuming that the EvalTimeSpan is only set day-wise I use from
               'EvalTimeMin the date part and from EvalTimeMin the time part....
               '
               'na I just saw there's no need to trace that any further since
               'I got the problem that I need license for the correct byteCodeKey solved.
            
            
            
            End If
              
          ' Modifies ByteCode_Key if there are SrvRestrictionsItems
            Morf_ByteCodeKey
      


            If IC_MemberID = 6666 Then '6666->(1a0a)
               log "The encoded file %s was created with an unauthorised copy of the ionCube PHP Encoder. Please contact legal@ioncube.com with details of the provider of this PHP script."
            End If
            
            Dim Header_MoreData_IncludeXorKey&
            Header_MoreData_IncludeXorKey = .int32 '1 +c                ->+28
            
            
            log "Copy of IncludeXorKey in Header_MoreData: 0x" & H32(Header_MoreData_IncludeXorKey)
            
            Dim Header_MoreData_v5_1&, _
               Header_MoreData_v5_2&, _
               Header_MoreData_v5_3&, _
               SomeFlag&, _
               Header_MoreData_v5_5&
            Header_MoreData_v5_1 = .int32
            Header_MoreData_v5_2 = .int32
'2 2nd-byte only
            Header_MoreData_v5_3 = .int16
            SomeFlag = .int8
            Header_MoreData_v5_5 = .int8
            
            log "unknown Data: " & Join(Array( _
                H32(Header_MoreData_v5_1), H32(Header_MoreData_v5_2), _
                H16(Header_MoreData_v5_3), H8(SomeFlag), _
                H8(Header_MoreData_v5_5)), " ")

'4
            Dim EvalTimeMin&, EvalTimeMin_enc
            EvalTimeMin_enc = .int32           'GetInt32 at +24  and add   83941958 ' 500DA46  -> +3c
            EvalTimeMin = EvalTimeMin_enc + 1023976199
            
'3
            Dim EvalTimeMax&, EvalTimeMax_enc
            EvalTimeMax_enc = .int32          'GetInt32 at +20  and add 1023976199 '3D08A307  -> +40
            EvalTimeMax = EvalTimeMax_enc + 83941958
            
            '4AD7DD0D -> 16.10.2009
            '740A9780 -> 11.09.2031
            Const SecondsPreDay& = 86400 '=24h*60min*60s
            
            If (SomeFlag <> 0) Then
               GoTo skip_IC_MemberID
            Else
               If (IC_MemberID = 0) Then
skip_IC_MemberID:
                  
                  'SecondsPreDay*3 => 259200
                  If (EvalTimeMax - EvalTimeMin) >= 3 * SecondsPreDay Then
                     log "The encoded file %s was created with an unauthorised copy of the ionCube PHP Encoder. Please contact legal@ioncube.com with details of the provider of this PHP script."
                  End If
               End If
            End If
            
            
            '"Your system clock is set more than a day before %s was encoded"
               '~Year: " & 1970 - 1 + EvalTimeMin \ (SecondsPreDay * 360)
               log "EvalTimeMin: " & FormatCTime(EvalTimeMin) & "   " & H32(EvalTimeMin_enc) & " '+0500DA46'=>" & H32(EvalTimeMin) & " seconds since 1.1.1970"
             '"The encoded file %s has expired."
               log "EvalTimeMax: " & FormatCTime(EvalTimeMax) & "   " & H32(EvalTimeMax_enc) & " '+3D08A307'=>" & H32(EvalTimeMax) & " seconds since 1.1.1970"
 
 
           'GetByte at +1e                                 -> +38
           

         End With 'Header_MoreData
         
      End With 'Header
      
      log_verbose "BodyStartOffset: 0x" & H32(.Position)
      
   End With 'Encrypted
      
      
   Select Case IC_Type_Major
   Case 2, 3 '"BASIC"
      myStop 'Not implemented
      'PHP3_ReadBody
   Case 4 '"BASIC"
      PHP4_ReadBody encrypted, PhpFlags, f2
   
   Case 5 '"BASIC"
'      myStop 'Not implemented in ioncube_loader_win_4.3.dll
      
      PHP4_ReadBody encrypted, PhpFlags, f2, True
   
   End Select

End Sub

Private Sub PHP4_ReadBody(encrypted As StringReader, PhpFlags&, f2, Optional IsPhp50 As Boolean = False)
   
LogStage 6, "Extracting & Inflating IC_body from RawData (->IC_body.bin)"
   
   
   With encrypted
      Dim Body As New StringReader

      
      Const MD5LEN& = &H10

      
      Dim my_Seek&
      my_Seek = .Position
      
      
'      Dim Body_MT_Seed&
      Body_MT_Seed = .int32
      log_verbose H32(.Position) & "  Body_MT_Seed: " & H32(Body_MT_Seed)
      random_init Body_MT_Seed
      
      
      If f2 Then
         Bytecode_MT_Seed = .int32
         log_verbose H32(.Position) & "  Bytecode_MT_Seed: " & H32(Bytecode_MT_Seed)
         ' When you encode the same file twice and comparing the results
         ' you may notice that the first part (here '85EE...') of the Bytecode_MT_Seed is the same
         ' Here a note on how the encoder calucates that value and some explanation:
         ' 8<<  003A41D8 EDI (some other value)
         '   =  3A41D800 (provisional result)
         '   +  4B741D8D CurTime
         '   =  85B5F58D (provisional result)
         '   +  0038F1C8 InPhpFile strPointer in Mem(Stack)
         '   =  85EEF23D  <- Example Bytecode_MT_Seed
         '
         ' as you see that some is depended on the time - however that other values may slightly differ
         ' what makes an exact calcuation of CurTime unsure.
         
         
      End If
      

      

     'Inside GetInt32 Stream
     
      Do Until .EOS ' <- "with Encrypted"
      
         Dim Flag&
         Flag = .int8
      If .EOS Then Exit Do
      
         Dim Blocksize&
         Blocksize = .int8
      If .EOS Then Exit Do
         
'         log "Chunk at " & H32(.Position) & " HeaderSize: " & H16(Blocksize) & " Flags: " & H8(Flag) & " => DestOffset: " & H32(Body.Position)
         
         If Flag And &H80 Then
            
            Flag = (Flag And &HE0)
            
            Select Case Flag
               Case &H80
                  'f1 Cont
                  Body.int8 = Blocksize
                  Inc Blocksize
                  'Do  CRC only
            
            
               Case &HA0
                  'f2
'                  myStop
                  .Move -1
               
                  'Special copy
                  Dim CrcFromFile As New StringReader
                  Replace3F_Questionmark encrypted, 4, CrcFromFile

                  
'                  log "CRC for Block: " & H32(CrcFromFile.int32)
                  
               
            
               Case &HC0
                  Body.int8 = &H3C '3C => '<'
                  
               Case Else
                  myStop
                  
              End Select
         
         Else
         
         'GetCRC-HeaderKey
         '      Dim CrcData$
         '      .Move -2
         '      CrcData = .FixedString(CrcSize+2)
         '      Chunk_Alder32 = Adler32(CrcData)
         '     .Move -(CrcSize + 2)
               
            'Get Block
            Dim EncBlock As New StringReader
            EncBlock.Data = .FixedString(Blocksize)
         
            
          ' Decrypt Block
            EncBlock.EOS = False
            Do Until EncBlock.EOS
               Dim K&
               K = Mersenne_twister_random
'               log H32(k) & "  Pos: " & H16(EncBlock.Position)
               
         '      myStop
         '      Body.EOS = False
               Body.int8 = EncBlock.int8 Xor K
            Loop
         
         End If
         

      Loop ' Read Body
   
   
    ' Save Data
      Dim InputFile$
      InputFile = "IC_body_deflated.bin"
      FileSave InputFile, Body.Data
      
      
      Dim OutputFile$
      OutputFile = "IC_body.bin"
      
      With ZlibTool
         .InputFile = InputFile
         .OutputFile = OutputFile
         .Decompress
         
         If (.Status <> "Success") Or FileLen("IC_body.bin") = 0 Then
         
            'Err.Raise vbObjectError, ,
            log "ZlibTool extraction " & _
                     InputFile & " -> " & OutputFile & " failed"
            
         Else

DisableDeleteTmpFile = True
         
            If DisableDeleteTmpFile = True Then
               Kill InputFile
            End If
         
         End If
         
         
      End With
      

      
      
'      log H32(.int32) ' is 80
'      log H32(.int32) ' is 00
   End With

   With File
      .Create OutputFile, False, False, True
      
      If f2 Then
         'NewVersion
         .Position = 4
      Else
         'OldVersion
         .Position = 0
      End If

      
      LogStage 7, "Reading " & OutputFile
      
      
    ' Important for ProcessBody
      random_init Bytecode_MT_Seed
      
      FrmMain.List_Source.AddItem "<?php"
      
    ' MainFunction
      ProcessBody File, PhpFlags, f2, IsPhp50
      
      Dim FunctionsCount&
      FunctionsCount = .intValue
      log "FunctionsCount: " & H16(FunctionsCount)
      
    ' Functions
      Dim FunctionsNr&
      For FunctionsNr = 1 To FunctionsCount
         
         log "----------------- Function #" & FunctionsNr
         ProcessBody File, PhpFlags, f2, IsPhp50
         
      Next
      
      
      'Classes
      Dim Classes&
      Classes = .intValue
      
      log "Classes: " & H16(Classes)
      
      
    ' Classes
      Dim classNr&
      For classNr = 1 To Classes
         
         log_verbose "Class #" & classNr
         
         Dim class_type As Byte
         class_type = .ByteValue
         log "class_type: " & H8(class_type)
         
         Debug.Assert class_type = 2 'So far it was always 2
         
         Dim Class_Name$
         Class_Name = .FixedString(.longValue + 1)
         log "Class_Name: " & Class_Name
'php5ts.zend_initialize_class_data



       ' Note: Value not uses by php5 decoder
         Dim S2_Type As Byte
         S2_Type = .ByteValue
         log "S2_Type: " & H8(S2_Type)
         
         If IsPhp50 Then
         Dim OutArg5&, v0_Size&, v2&, v3&
            OutArg5 = .longValue
            v0_Size = .longValue
            log "v0_Size: " & H32(v0_Size)
            
           'v1 = CurPhpFullFileName
            v2 = .longValue
            log "v2: " & H32(v2)
            
            v3 = .longValue
            log "v3: " & H32(v3)
            
            If v0_Size <> 0 Then
               If class_type = 1 Then
'                  myStop
                  'alloc v0_Size*4
               Else
'                  myStop
                  'alloc v0_Size*4
               End If
               
            End If
            
            Dim ClassComment$
            ClassComment = .FixedString(.longValue + 1)
            log "ClassComment: " & ClassComment
            

            
         End If
         
         
         
         Dim extends As New StringReader
         With extends
            .Data = File.FixedString(File.longValue + 1)
            .DisableAutoMoveOnRead = True
            Select Case .int8
               Case 0
               Case &HD
                  .Data = "[Obfuscated]" & ValuesToHexString(extends)
                  log "extends: " & .Data
                  
               Case Else
                  'extends = LCase(.Data)
                  log "extends: " & .Data
                  
            End Select
         End With


         Dim Class_Name2$
         Class_Name2 = .FixedString(.intValue)
         log "Class_Name2: " & Class_Name2
         
         If Len(Class_Name2) = 1 Then log "implements ?"
         
         
         Dim MethodesCount&
         MethodesCount = .intValue
'         log "MethodesCount: " & MethodesCount
         
         Dim MethodeNr&
         For MethodeNr = 1 To MethodesCount
            log "----------------- Methode #" & MethodeNr
            ProcessBody File, PhpFlags, f2, IsPhp50
            '__construct,__destruct, __call, __clone , __get, __set
         Next
         
         If IsPhp50 Then
            GetClassConsts File
            GetClassConsts File
            GetClassProperties File
         End If
         
         
         DecodeBasicTypes File '3
         
      Next
             
      FrmMain.List_Source.AddItem "?>"
      
      .CloseFile
      
      log "File sucessfully processes!"
      
   End With
   
   

End Sub
'

'Note: 0x3C -> '<'
Private Sub Replace3C_TagStart(ByRef InputData As StringReader, Size&, ByRef OutputData As StringReader)
   ReplaceXX InputData, Size, OutputData, &H3C
End Sub
'Note: 0x3F -> '?'
Private Sub Replace3F_Questionmark(ByRef InputData As StringReader, Size&, ByRef OutputData As StringReader)
   ReplaceXX InputData, Size, OutputData, &H3F
End Sub

   
Private Sub ReplaceXX(ByRef InputData As StringReader, Size&, ByRef OutputData As StringReader, ReplaceByte As Byte)


   
   With InputData
      
      Dim InputData_StartPos&
      InputData_StartPos = InputData.Position
      
      Dim OutputData_StartPos&
      OutputData_StartPos = OutputData.Position
      
    ' Fill the output buffer
      Do Until OutputData.Position >= (OutputData_StartPos + Size)
         Dim i8&
         i8 = .int8
         
         If i8 = &HFF Then
           'if 00..7f -> FF else(7f..ff)-> 3C
            i8 = IIf(.int8 And 128, ReplaceByte, &HFF)
         End If
         
        'Save
         OutputData.int8 = i8
      Loop
      
     'Seek to start of new data
      OutputData.Position = OutputData_StartPos
      
   End With


'   int __cdecl ReplaceCopyFF(BYTE *aIn, int aOut, int aSize)
'{
'  BYTE *In; // eax@1
'  BYTE *Out; // edx@1
'  int i; // esi@2
'  BYTE InByte; // cl@3
'  char InByte2; // cl@4
'
'  In = aIn;
'  Out = (BYTE *)aOut;
'  if ( aSize )
'  {
'    i = aSize;
'    Do
'    {
'      InByte = *In++;
'      if ( InByte == '\xFF' )
'      {
'        InByte2 = *In++ & 0x80;
'        if ( InByte2 )
'          InByte = '<';
'        Else
'          InByte = '\xFF';
'      }
'      *Out++ = InByte;
'      --i;
'    }
'    while ( i );
'  }
'  return (int)In;
'}


End Sub

Function LoadLicString(str As StringReader)
   Dim Size&, Dummy&, Data$
   With str
      Dummy = .int32
      Size = .int32
      'AllocMem
      LoadLicString = .FixedString(Size)
   End With
End Function

Function loadValue&(str As StringReader)
   Dim dummy1&, v1&, Dummy3&
   With str
      dummy1 = .int32
      
      loadValue = .int32
      Debug.Assert loadValue = 4
      
      Dummy3 = .int32
'      log "loadValue_Enable_auto_prepend\Append_file : " & Dummy3
      loadValue = Dummy3
      
   End With

End Function


Function loadString$(str As StringReader)
   Dim Size&, Dummy&, Data$
   With str
      Dummy = .int32
      Size = .int32
      loadString = .FixedString(Size)
   End With

End Function



Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub ListLog_DblClick()
   SaveLog
End Sub

Private Sub SaveLog()
   Dim logFileName As New ClsFilename
   logFileName.mvarFileName = Txt_Filename
   logFileName.Ext = "log"
   
   Dim logFile As New FileStream
   With logFile
      .Create logFileName.FileName, True, False, False
      .Data = GetLogdata(ListLog)
      .CloseFile
   End With
End Sub

Private Sub SaveSource()
   Dim SrcFileName As New ClsFilename
   SrcFileName.mvarFileName = Txt_Filename
   SrcFileName.Name = SrcFileName.Name & "_Decoded"
   
   FileSave SrcFileName.FileName, GetLogdata(List_Source)
   
   SrcFileName.Ext = ".txt"
   FileSave SrcFileName.FileName, Join(Dumper_Data, vbCrLf)

End Sub




Private Sub Timer_OleDrag_Timer()
   Timer_OleDrag.Enabled = False
   Txt_Filename = FilePath_for_Txt
End Sub



Private Sub Txt_Filename_Change()
   On Error GoTo Txt_Filename_err
   If FileExists(Txt_Filename) Then
      
    ' Block any input during DoEvents...
      Txt_Filename.Enabled = False
      
         DecodePhp Txt_Filename
    
    ' ... until processing is finished
      Txt_Filename.Enabled = True
      
   End If

Txt_Filename_err:
Select Case Err
   Case 0
   Case Else
      log "-> ERROR: " & Err.description
End Select
End Sub

Private Sub Txt_Filename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   FilePath_for_Txt = Data.Files(1)
   Timer_OleDrag.Enabled = True
End Sub
Private Sub ListLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   FilePath_for_Txt = Data.Files(1)
   Timer_OleDrag.Enabled = True
End Sub


Private Function ReadHeaderData$(ByRef Data As StringReader, ByVal Size&)
   
'   log "Reading Header:"
   Dim outdata As New clsStrCat
   With Data
      
      Do While Size > 0
         Dim Chunk_Flag&, Chunk_Size&
         Chunk_Flag = .int8
         Chunk_Size = .int8
         If Chunk_Flag And &H80 Then
            
            Dec Size, Chunk_Size
            If Size >= 0 Then
               outdata.Concat .FixedString(Chunk_Size)
            End If
         
            If Chunk_Flag And &H40 Then
   '            log "  Chunk_Flag: " & H8(Chunk_Flag) & _
                   " (and 80)-> ReadSize: " & H8(Chunk_Size) & _
                   " (and 40)-> Appending '<' to data"

               Dec Size
               outdata.Concat "<" '3c
            Else
    '           log "  Chunk_Flag: " & H8(Chunk_Flag) & _
                   " (and 80)-> ReadSize: " & H8(Chunk_Size)

            End If
            
         Else
               
            Dec Size, &HE3
            If Size > 0 Then
               outdata.Concat .FixedString(&HE3)
            ElseIf Size = 0 Then
            Else
               Err.Raise vbObjectError, , "Error Reading Headerdata chunks"
            
            
            End If
            
  '          log "  Chunk_Flag: " & H8(Chunk_Flag) & _
                " ->FixedReadSize=E3)"
            
         End If
       
       Loop
         
      End With
      ReadHeaderData = outdata.value


End Function

Private Sub HandleCustomErrorMessages(ByRef Chunk As StringReader)
   With Chunk
      Dim CustomErrorMessages()
      CustomErrorMessages = Array("corrupt-file", "expired-file", _
         "no-permissions", "clock-skew", "untrusted-extension", _
         "license-not-found", "license-corrupt", "license-expired", _
         "license-property-invalid", "license-header-invalid", _
         "license-server-invalid", "unauth-including-file", _
         "unauth-included-file", "unauth-append-prepend-file")
      
                        
      Dim CustomLoaderEventMessagesCount&
      CustomLoaderEventMessagesCount = .int8
      log "Customised error messages entries: 0x" & H8(CustomLoaderEventMessagesCount)
      
      
      'Alloc CustomLoaderEventMessagesCount*8
      Dim i&
      For i = 1 To CustomLoaderEventMessagesCount
         
         Dim CustomErrorMsgID&
         CustomErrorMsgID = .int8
         
         Dim CustomErrorMsg$
         CustomErrorMsg = .FixedString(.int32)
         
         On Error Resume Next
         log "  #" & H8(CustomErrorMsgID) & _
             "[" & CustomErrorMessages(CustomErrorMsgID - 1) & "] => '" _
             & CustomErrorMsg & "'"
         .Move 1 'wrap over '\0' string terminator
      Next
   End With
End Sub

Private Sub HandleIncludeRestrictions(ByRef Chunk As StringReader)
   With Chunk
      Dim IncludeRestrEntries&
      IncludeRestrEntries = .int8
      
      log "Include file protection entries: 0x" & H8(IncludeRestrEntries)
      Dim i&
      For i = 1 To IncludeRestrEntries
            
            Dim IncludeDummy
            IncludeDummy = .int8

         
      
            Dim IncludeKey1 As New StringReader
            IncludeKey1 = ReadXoredData(Chunk, IncludeXorKey)
            
            Dim IncludeKeyHandler
            IncludeKeyHandler = ReadXoredData(Chunk, IncludeXorKey)
            
            log " #" & i & "[" & IncludeKey1.FixedString(3) & "]: '" & IncludeKey1.FixedString(-1) & _
                "' Handler: '" & IncludeKeyHandler & "'"
            
      Next

   End With
End Sub

Public Sub HandleServerRestrictions(ByRef Chunk As StringReader, SRestr As clsStrCat)
   With Chunk
   
   
      Dim RS_cols ', i
      RS_cols = .int8 ' * &H8
      Dim i
      For i = 1 To RS_cols
      
         Dim dataType
         dataType = .int8
         Select Case dataType
            Case 0 'ip
            
               SRestr.Concat "IPs: "
               Dim rs_entries, entry_i
               rs_entries = .int8 '* &h14
               For entry_i = 1 To rs_entries
               
                  Dim IsNetMask
                  IsNetMask = .int8 '* &h14
                  
                  Dim ip4_1 As Byte, ip4_2 As Byte, ip4_3 As Byte, ip4_4 As Byte
                  ip4_4 = .int8
                  ip4_3 = .int8
                  ip4_2 = .int8
                  ip4_1 = .int8
                  
                  Dim ip4Mask_1 As Byte, ip4Mask_2 As Byte, ip4Mask_3 As Byte, ip4Mask_4 As Byte
                  ip4Mask_4 = .int8
                  ip4Mask_3 = .int8
                  ip4Mask_2 = .int8
                  ip4Mask_1 = .int8
                  
                  
                  SRestr.Concat ip4_1 & "." & ip4_2 & "." & ip4_3 & "." & ip4_4
                  If IsNetMask Then
                     SRestr.Concat "_NetMask(" & ip4Mask_1 & "." & ip4Mask_2 & "." & ip4Mask_3 & "." & ip4Mask_4 & "), "
                  Else
                     SRestr.Concat "-" & ip4Mask_1 & "." & ip4Mask_2 & "." & ip4Mask_3 & "." & ip4Mask_4 & ", "
                  End If
   
                  
               Next
               SRestr.Concat " | "
   
            
            Case 1 'Mac
               SRestr.Concat "MACs: "
               rs_entries = .int8 '* &h14
               For entry_i = 1 To rs_entries
                  Dim Mac As New StringReader
                  Mac = .FixedString(6)
                  SRestr.Concat ValuesToHexString(Mac, ":") & "  "
               Next
               SRestr.Concat " | "
            
            
            Case 2, 4  'Domain
               'Dim rs_entries
               SRestr.Concat "Domains: "
               rs_entries = .int8 '*4
               For entry_i = 1 To rs_entries
                  Dim Domain$
                  Domain = .zeroString
                  SRestr.Concat Domain & " "
               Next
               SRestr.Concat " | "
            
            
            Case 3
            
               SRestr.Concat "IncludeRestriction "
               rs_entries = .int8 '* &hc
               For entry_i = 1 To rs_entries
               
                  Dim IncludeKey1 As New StringReader
                  IncludeKey1 = ReadXoredData(Chunk, IncludeXorKey)
                  SRestr.Concat " [" & IncludeKey1.FixedString(3) & "]: '" & IncludeKey1.FixedString(-1) & "'"
                  
                  Dim IncludeKeyHandler
                  IncludeKeyHandler = ReadXoredData(Chunk, IncludeXorKey)
                  SRestr.Concat " Handler: '" & IncludeKeyHandler & "'"
               
               Next
               
            Case Else
               Err.Raise vbObjectError, , "Serverrestricions: Invalid type for the following dataRecord  @" & H32(.Position)
         End Select
         
      Next ' cols
   
   End With
End Sub

Function ReadXoredData$(ByRef Chunk As StringReader, Key&)
      Dim Size&
      Size = Chunk.int16
      Size = (Size Xor Key) And 65535 '&HFFFF

      Dim Data As New StringReader
      With Data
         .Data = Chunk.FixedString(Size)
         
         .DisableAutoMoveOnRead = True
         
         .EOS = False
         Do Until .EOS
            .int32 = .int32 Xor Key
         Loop
         
         .Position = Size
         .Truncate
         
         ReadXoredData = .Data
         
      End With
      

End Function
Public Function IC_ADLER32$(Data As StringReader)
   IC_ADLER32 = "<not fully implemented yet>"
'   myStop
'Exit Function
   
   Dim StartL&, StartH&
   StartL = 17
   StartH = 0
'   With Data
'     'Note: 5552 = 347 *16
'      Dim MaxPos&
'      MaxPos = Min(.length, 5552)
'
'      Do Until (.Position >= MaxPos)
'         Dim charL As Long
'         charL = .int8
'
''         Dim charH As Long
''         charH = .int8
'
'         StartL = StartL + charL
'         StartH = StartH + StartL ' + charH
'      Loop
      
'   End With
   
   IC_ADLER32 = Adler32(Data, StartL, StartH)
End Function

Private Function Morf_ByteCodeKey()
         
      
'         i 've a problem to implement this formula in VB:
'         Result = 0x92492493 * ByteCodeKey  ;Problem: Overflow error :(
'         Result = 0x92492493 >> 32 ;64Bit part
'
'         Result = Result + ByteCodeKey
'
'         Example Data:
'         0x92492493 * 0x017FCCBD => 0xFF5B83AF 01122487
'         0xFF5B83AF 01122487 >> 32 => 0xFF5B83AF
'
'          0xFF5B83AF+0x017FCCBD => 0x00DB506C
'           0x00DB506C >> 2       =>   0036D41B
'
'         It 's possible to 'transform' it to:
'         Result = (0x92492493 >> 16) * (ByteCodeKey >> 16)
    If SrvRestrictionsItems Then 'And (LicenseFile = "") Then
      
              
      Bytecode_MT_XorKey = (&H92492493 / &H10000) * (Bytecode_MT_InitKey / &H10000) '924 924 93
      Bytecode_MT_XorKey = AddInt32(CDbl(Bytecode_MT_XorKey), CDbl(Bytecode_MT_InitKey))
      Bytecode_MT_XorKey = Bytecode_MT_XorKey \ 4      'SAR     EDX, 2
     ' SHR     EAX, 1F ->   Bytecode_MT_XorKey += Bytecode_MT_XorKey >> 31
      If Bytecode_MT_XorKey < 0 Then
        myStop
        Inc Bytecode_MT_XorKey
      End If
      
      
     'Note that is done by the Check ServerRestriction Function
     '(this as 3 references - however so far I implemented only one/this reference)
      Dim Arg_SkipSub13&, Arg_ServerRestrRecords&
      Arg_SkipSub13 = 0 'Fixed value
      Arg_ServerRestrRecords = SrvRestrictionsItems
   '            Arg_ServerRestrRecords=RestrictionRows
   
     If Arg_SkipSub13 = 0 Then Dec Bytecode_MT_XorKey, 13 * Arg_ServerRestrRecords
      
      log "Bytecode_MT_XorKey: " & H32(Bytecode_MT_XorKey) & " ( ((((Bytecode_MT_InitKey * 0x92492493)>>32) + Bytecode_MT_InitKey) >> 2) - 13 * SrvRestrictionsItems"
      
   Else
   
      Bytecode_MT_XorKey = Bytecode_MT_InitKey
      log "Bytecode_MT_XorKey: " & H32(Bytecode_MT_XorKey)
      
   End If
   
End Function


Private Function SumOfTheCharsFrom&(Data$)
   With New StringReader
      .Data = Data
      Do Until .EOS
         SumOfTheCharsFrom = SumOfTheCharsFrom + .int8
      Loop
   End With
End Function
