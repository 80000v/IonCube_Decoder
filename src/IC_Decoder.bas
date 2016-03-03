Attribute VB_Name = "IC_Decoder"
Option Explicit

Public Const BITMASK_CMD_00000001_HAS_RESULT& = &H1
Public Const BITMASK_CMD_00000010_HAS_OP1& = &H2
Public Const BITMASK_CMD_00000100_HAS_OP2& = &H4
Public Const BITMASK_CMD_00011000_EXTENDED_VALUE& = &H18
Public Const BITMASK_CMD_11100000_UNKNOWNFLAGS& = &HE0


Public icKnownFunctions As Variant
Public icByteCodeNames As Variant

Public Bytecode_MT_Seed&

Public Bytecode_MT_InitKey As Currency
Public Bytecode_MT_XorKey As Currency

Dim StringsData As New StringReader


'/* overloaded elements data types */
'#define OE_IS_ARRAY   (1<<0)
'#define OE_IS_OBJECT   (1<<1)
'#define OE_IS_METHOD   (1<<2)
'
'
'/* Argument passing types */
'#define BYREF_NONE 0
'#define BYREF_FORCE 1
'#define BYREF_ALLOW 2
'#define BYREF_FORCE_REST 3

Enum op_type
   IS_CONST__1& = 1
   IS_TMP_VAR__2& = 2
   IS_VAR__4& = 4
   IS_UNUSED__8& = 8
End Enum

Enum obj_type
   IS_NULL__0 = 0
   IS_LONG__1 = 1
   IS_DOUBLE__2 = 2
   IS_STRING__3 = 3
   IS_ARRAY__4 = 4
   IS_OBJECT__5 = 5
   IS_BOOL__6 = 6
   IS_RESOURCE__7 = 7
   IS_CONSTANT__8 = 8
   IS_CONSTANT_ARRAY__9 = 9
End Enum

Enum is_ref
   ZEND_RETURN_VAL__0 = 0
   ZEND_RETURN_REF__1 = 1
End Enum


'Public Const ZEND_DO_FCALL_BY_NAME& = &H3D
Public Const ZEND_CTOR_CALL& = &H2
Type IC_CmdAttrib
    a_op_type      As op_type

    '_zval_struct {...
    c_Object     As Variant 'As Long
    d_ObjectSize As Long
    e_FlagsPHP5      As Long
    e_Obj_Type      As obj_type 'Byte
    is_ref      As is_ref
    refcount      As Long ' actually its and UINT16 - but VB offers only INT16(integer)
End Type

'struct _zval_struct {
'   /* Variable information */
'   zvalue_value value;      /* value */
'   zend_uchar type;   /* active type */
'   zend_uchar is_ref;
'   zend_ushort refcount;
'};
'typedef union _zvalue_value {
'   long lval;               /* long value */
'   double dval;            /* double value */
'   struct {
'      char *val;
'      int len;
'   } str;
'   HashTable *ht;            /* hash table value */
'   zend_object obj;
'} zvalue_value;
'typedef struct _zend_object {
'   zend_class_entry *ce;
'   HashTable *properties;
'} zend_object;

Type IC_zend_op
    a_opcode    As Long
    b_Reserved1  As Long
    c_result     As IC_CmdAttrib
    c_op1   As IC_CmdAttrib
    c_op2    As IC_CmdAttrib
    f_extended_value     As Variant ' (was Long) change to so I can distinguish between 'empty' and '0'
    g_lineno      As Long
End Type

Public op_extended_value()

Private Sub LogTableHead_Verbose(TableHead$)
   log_verbose TableHead
   log_verbose String(Len(TableHead), "-")
End Sub

Private Sub LogTableHead(TableHead$)
   log TableHead
   log String(Len(TableHead), "-")
End Sub

Private Sub LogStage(Stage&, Title$)
   FrmMain.LogStage Stage, Title
End Sub

Public Function GetNum&(dataStream As StringReader, ByRef FirstNonNumDigit$)
   With dataStream
      Do Until .EOS
         
         Dim Digit$
         Digit = .FixedString(1)
      If IsNumeric(Digit) = False Then Exit Do
      
         GetNum = GetNum * 10 + Digit
      Loop
      
      FirstNonNumDigit = Digit
      
   End With
End Function

Public Function DecodeBasicTypes_rek$(dataStream As StringReader)
   With dataStream
      
      
      Dim cmd$
      cmd = .FixedString(1)
      
      
      Select Case cmd
      Case "n"
         'EDI+8=0
         DecodeBasicTypes_rek = ""
         
      Case "b", "i"
         Dim Seperator$
         DecodeBasicTypes_rek = GetNum(dataStream, Seperator)
         Debug.Assert Seperator = ";"
         
      Case "d"
      'Float
      '";"
         myStop
      
      Case "c", "s"
         
         'c4'null
         ' ^
         Dim value$
         value = .FixedString(GetNum(dataStream, Seperator))
         'c4'null
         '       ^

'         Dim v1&
         If cmd <> "s" Then
          ' Constante
            DecodeBasicTypes_rek = value
'            v1 = -5 + 8
         Else
          ' String
            DecodeBasicTypes_rek = StrQuote(value)
'            v1 = 0 + 8

         End If
         
         
      Case "[", "{"
'         Dim v1&
'         If cmd <> "{" Then
'            v1 = 5 + 4
'         Else
'            v1 = 0 + 4
'         End If
         ' v23 = emalloc(40);
         'zend_hash_init(v23, 10, 0, 0, 0);

       ' Append '{'
         DecodeBasicTypes_rek = DecodeBasicTypes_rek & cmd

        '{1:0s6'KOI8-R1:2s10'ISO-8859-1...
        ' ^
'         StringSize = GetNum(dataStream, Seperator)
        '{1:0s6'KOI8-R1:2s10'ISO-8859-1...
        '   ^
         value = .FixedString(GetNum(dataStream, Seperator))
         
         Do Until (Seperator = "}") Or .EOS
                    
            Dim bIsFirst As Boolean
            bIsFirst = (Len(DecodeBasicTypes_rek) = 1)
            
            If (Seperator = "\") Or (Seperator = "'") Then
             ' DecodeBasicTypes_rek
             ' zend_hash_add_or_update'
               'myStop
            
               DecodeBasicTypes_rek = DecodeBasicTypes_rek & _
                                      IIf(bIsFirst, "", ", ") & _
                                      StrQuote(value) & " => " & _
                                      DecodeBasicTypes_rek(dataStream)
            
            ElseIf (Seperator = ":") Then
'              '{1:0s6'KOI8-R1:2s10'ISO-8859-1...
'              '     ^
'               .Move -1
              '{1:0s6'KOI8-R1:2s10'ISO-8859-1...
              '    ^

               
               DecodeBasicTypes_rek = DecodeBasicTypes_rek & _
                                      IIf(bIsFirst, "", ", ") & _
                                      DecodeBasicTypes_rek(dataStream)
              '{1:0s6'KOI8-R1:2s10'ISO-8859-1...
              '             ^
               ' zend_hash_index_update_or_next_insert(v29, Number, &DataLeft_, 4, 0, 1);
            Else
               myStop 'unexpected seperator
            End If
            
            value = .FixedString(GetNum(dataStream, Seperator))


         Loop
         
       ' Append '}'
         DecodeBasicTypes_rek = DecodeBasicTypes_rek & Seperator
      
      
      Case Else
         myStop
      
      End Select
      
   End With
End Function
Public Sub GetClassConsts(FileStream As FileStream)
   DecodeBasicTypes FileStream, "Const "
End Sub

Public Sub GetClassProperties(FileStream As FileStream)
   With FileStream
      Dim PropertyCount&
      PropertyCount = .longValue
'      FL_verbose "properties: " & H32(PropertyCount)
    
    ' if you ran intop 'stop' here - there is probably something wrong
      If PropertyCount > 10000 Then myStop: PropertyCount = 10000
            
      Dim i&
      For i = 1 To PropertyCount
'         myStop 'Branch not completely implemented yet
         
         Dim Property_Name$
         Property_Name = szNullCut(.FixedString(.longValue + 1))
         
         Dim Property_Type&, Property_NamePtr&, Property_NameLen&, Property_InitValue&
         Property_Type = .longValue
         Property_NamePtr = .longValue
         Property_NameLen = .longValue
         Property_InitValue = .longValue
         
         Dim scrope$
         Select Case (Property_Type And &HFF00)
            Case &H100
               'scrope = "var"
               scrope = "public"
               
            Case &H200
               'zend_mangle_property_name "*"
               scrope = "protected"
               
            Case &H400
               'zend_mangle_property_name ClassName
               scrope = "private"
               
            Case Else
               myStop
         
         End Select
         
         If (Property_Type And &HF) = 1 Then
            scrope = scrope & " " & "static"
         End If
         
         log "property #" & i & " " & scrope & " $" & Property_Name & " - 0x" & H32(Property_InitValue) & " len: " & Property_NameLen
      Next
   End With
End Sub





Public Sub DecodeBasicTypes(FileStream As FileStream, Optional Prefix$ = ".")
   With FileStream
         
      Dim BasicTypes&
      BasicTypes = .longValue
'      FL_verbose "BasicTypes: " & H32(BasicTypes)
    
    ' if you ran intop 'stop' here - there is probably something wrong
      If BasicTypes > 10000 Then myStop: BasicTypes = 10000
            
      Dim BasicType&
      For BasicType = 1 To BasicTypes
'         myStop 'Branch not completely implemented yet
         
         Dim BasicType_Name$
         BasicType_Name = szNullCut(.FixedString(.longValue + 1))
         
         Dim BasicArgs As New StringReader
         BasicArgs.Data = szNullCut(.FixedString(.longValue + 1))
        
         log "BasicType #" & BasicType & " " & BasicType_Name & " - " & BasicArgs
         
         BasicArgs = DecodeBasicTypes_rek(BasicArgs)
'         If BasicArgs.EOS = False Then
'            myStop
'            'ioncube_Warning("Garbage after decoding basic types")
'         End If
         
         log Prefix & BasicType_Name & "  " & BasicArgs

      Next

   
   End With

End Sub


Private Function Read32BitString$(dataStream As StringReader)
   With dataStream
      Read32BitString = .FixedString(.int32)
   End With
End Function


Sub ProcessBody(File As FileStream, Php_Flags&, f2, IsPhp50 As Boolean)


LogStage 8, "Read function body data"
   'DoRunTimeMemCrypt = Php_Flags And &H400
   With File
'      .Create "IC_body.bin"
      
'      If .length = 0 Then
'         .CloseFile
'         Err.Raise vbObjectError, "", "IC_body.bin is 0 byte. Probably decompression error."
'      End If
      
      log "Reading Body"
      
       
'      LanguageOptions = .longValue
'      FL_verbose H32(LanguageOptions)
      
'     'Dirty Quickfix
'      If LanguageOptions = 0 Then
         Dim LanguageOptions&
         LanguageOptions = .longValue
         FL_verbose H32(LanguageOptions) ' & "(Second Try)"
'      End If
'      Select Case LanguageOptions
'         Case 1
'            log "html/txt or other file"
'         Case 2
'            log "php file"
'         Case Else
'            log "unknows LanguageOptions"
'      End Select
      
      
      
      Dim FunctionName$ ' As New StringReader
      FunctionName = szNullCut(.FixedString(.intValue))
      
      Dim s2 As New StringReader
      Dim S2_size&
      S2_size = IIf(IsPhp50, &H70, &H40)
            
      log_verbose "S2[Fixed Size 0x" & H8(S2_size) & "]:"
      s2 = .FixedString(S2_size)
      Dim tmpstr$
      tmpstr = ""
      Do Until s2.EOS
         Dim value&
         value = s2.int32
         tmpstr = tmpstr & " " & H32(value)
         Select Case s2.Position
            
            Case &H8
               tmpstr = tmpstr & vbCrLf & " FuncNameStart -> "
            
            Case &HC
               Dim FuncNameStart&
               FuncNameStart = value
            
            
            Case &H10
               Dim s2_FuncScope$
               s2_FuncScope = value
            
               tmpstr = tmpstr & vbCrLf & " CmdData -> "
               
            Case &H14
               tmpstr = tmpstr & vbCrLf & " Args -> "
            
            
            Case &H18
               tmpstr = tmpstr & vbCrLf & " ArgsMin -> "
               
               Dim s2_0x14_Args&
               s2_0x14_Args = value
               
               
            Case &H1C
               tmpstr = tmpstr & vbCrLf & " php4Extra -> "
               
            Case &H20
               Dim s2_0x1c_php4Extra&
               s2_0x1c_php4Extra = value
               
               
            Case &H34
               tmpstr = tmpstr & vbCrLf & " Skipped1 -> "
   
            Case &H40
               tmpstr = tmpstr & vbCrLf & " PHPFileNamePtr -> "
               
            Case &H8
               tmpstr = tmpstr & vbCrLf & " Skipped2 -> "
               
            Case &H4
               tmpstr = tmpstr & vbCrLf & " Skipped3 -> "
               
            Case &HC
               tmpstr = tmpstr & vbCrLf & " Scope(PHP5 only) -> "
               
            Case &H20
               tmpstr = tmpstr & vbCrLf & " 0x60-Alloc-Byte-with 0x0000002 -> "
           
          ' php5
               
               
            Case &H44
               tmpstr = tmpstr & vbCrLf & " Php5-O1 "
               
            Case &H48
               tmpstr = tmpstr & vbCrLf & " Php5-O1Count "
               
            Case &H4C
               Dim s2_0x48_Php5_O1Count&
               s2_0x48_Php5_O1Count = value
   
               
         End Select
         
      Loop
      FL_verbose tmpstr
      log_verbose " "
      
      
      Dim scopeText$
      If IsPhp50 Then
         
         Select Case (s2_FuncScope And &HF00)
            Case &H0
               scopeText = ""
               
            Case &H100
               'scrope = "var"
               scopeText = "public"
               
            Case &H200
               'zend_mangle_property_name "*"
               scopeText = "protected"
               
            Case &H400
               'zend_mangle_property_name ClassName
               scopeText = "private"
               
            Case Else
               myStop
         
         End Select
         
         If (s2_FuncScope And &H1) Then
            scopeText = scopeText & " static"
         End If
         
         If (s2_FuncScope And &H2) Then
            scopeText = "abstract " & scopeText
         End If
      
      Else
      
       ' There is no 'public', 'protected                                                                              ' or private in PHP4
         scopeText = ""
      End If

      
      
      If (FunctionName = "") Then
         FL "FunctionName: " & scopeText & " function " & "{MainFunction}"
      Else
         FL "FunctionName: " & scopeText & " function " & FunctionName
         
         SetFunctionName FunctionName, scopeText
      End If
      
      log_verbose " "
      
      
      
      DecodeBasicTypes File
      
      
      If IsPhp50 Then
         
         Debug.Assert s2_0x48_Php5_O1Count = 0
         
         Dim s2_0x44 As New StringReader
         s2_0x44 = .FixedString(s2_0x48_Php5_O1Count * 8)
         With s2_0x44
            If .Length Then
               log_verbose "s2_0x44: "
               'LogTableHead_Verbose "StrTblOff StrLen"
               
               Do Until .EOS
   '               Dim tmp
                  log_verbose Join(Array( _
                        H32(.int32), H32(.int32) _
                  ))
               Loop
            End If
         End With
            
         
         Dim s2_0x1c As New StringReader
         s2_0x1c = .FixedString(s2_0x14_Args * &H18)
         With s2_0x1c
            .EOS = False
            If .Length Then
               log_verbose "s2_0x1c_ArgumentData: "
               LogTableHead_Verbose "StrTblOff StrLen"
               
               Do Until .EOS
   '               Dim tmp
                  log_verbose Join(Array( _
                        H32(.int32), H32(.int32), H32(.int32), H32(.int32), H32(.int32), H32(.int32) _
                  ))
               Loop
            End If
            
         End With

                  
      End If




    ' Get Cmd Count
      Dim CmdCount&
      CmdCount = .longValue
      FL_verbose H32(CmdCount) & "<-CmdCount"
      
      
                 
         
  '   Extract Commands
      Dim Cmds_EntrysCount&
      Cmds_EntrysCount = .longValue
      
      Dim IsObfuLineNos As Boolean
      IsObfuLineNos = (Php_Flags And &H800)
      If IsObfuLineNos Then
         Dim byteCmds As New StringReader
         byteCmds.Data = .FixedString(Cmds_EntrysCount * 2)
         
         log_verbose "byteCmds[Elementsize=0x2] Elements: " & H32(Cmds_EntrysCount)
         LogTableHead_Verbose "Command(Encrypted)"

         byteCmds.EOS = False
         Do Until byteCmds.EOS
            log_verbose H16(byteCmds.int16)
         Loop
      
      Else
         byteCmds.Data = .FixedString(Cmds_EntrysCount * 4)
         
         log_verbose "byteCmds[Elementsize=0x4] Elements: " & H32(Cmds_EntrysCount)
         LogTableHead_Verbose "Command(Encrypted)" & "  " & "Line"

         byteCmds.EOS = False
         Do Until byteCmds.EOS
            log_verbose "   " & H16(byteCmds.int16) & "  " & H16(byteCmds.int16)
         Loop
      End If
      
      
      
    ' Generate Key for byteCode
      Dim CmdXorKey As New StringReader
      With CmdXorKey
         
         .Data = Space((CmdCount + 1) * 4)
' Init Done outside Interprete
'      random_init Bytecode_MT_Seed

         .EOS = False
         Do Until .EOS
            Dim a&
            a = Mersenne_twister_random(-1, Bytecode_MT_XorKey)
'            log H32(a)
            .int32 = a
            
         Loop
      
      End With
      
      
      
  '   Extract Operants
      Dim CmdParamsCount&
      CmdParamsCount = .longValue
      
      Dim CmdParams As New StringReader
      Dim CmdParamsSize&
      CmdParamsSize = IIf(IsPhp50, &H14, &H10)
      CmdParams.Data = .FixedString(CmdParamsCount * CmdParamsSize)
      
      log_verbose "CmdParams[Elementsize=0x" & H8(CmdParamsSize) & _
                           "] Elements: " & H32(CmdParamsCount)
      LogTableHead_Verbose "Type      " & _
                           "Op1       " & _
                           "Op2       " & _
              IIf(IsPhp50, "Flags5    ", "") & _
                           "Flags"
      
      CmdParams.EOS = False
      Do Until CmdParams.EOS
      
          Dim index, goto_pos&, len_&, Flags5&, flags&
          index = CmdParams.int32
          goto_pos = CmdParams.int32
          len_ = CmdParams.int32
          If IsPhp50 Then
            Flags5 = CmdParams.int32
          End If
          flags = CmdParams.int32
          
          
         
         log_verbose H32(index) & "   " & H32(goto_pos) & "   " & H32(len_) & _
                     IIf(IsPhp50, "   " & H32(Flags5), "") & _
                     "   " & H32(flags)
      
'         FL_verbose H32(CmdParams.int32) & "   " & H32(CmdParams.int32) & "   " & H32(CmdParams.int32) & "   " & H32(CmdParams.int32)
      Loop
      
      
      
      
   ' PreRead Strings

      Dim StringsDataSize&
      StringsDataSize = .longValue
      StringsData = .FixedString(StringsDataSize)
      
      
      Const MaxLenStringForLog = 2048
      
      log "StringData [" & H32(StringsData.Length) & "]"
      Dim StringsDataShort As New StringReader
      StringsDataShort.Data = Mid(StringsData.Data, 1, MaxLenStringForLog)
      
      FL_verbose ValuesToHexString(StringsDataShort)
      log Replace(StringsDataShort.Data, Chr(0), ".")
      
      If StringsData.Length > MaxLenStringForLog Then
         log "<StringData exceeds " & MaxLenStringForLog & " chars and was cutoff to fit into that log>"
      End If
      
      Dim RemainData_Len&
      RemainData_Len = .Length - .Position
      log "Length of unprocessed data: " & H32(RemainData_Len)
      
'      If RemainData_Len > 4 Then
'
'         Dim overlayFileName As New ClsFilename
'         overlayFileName.FileName = "IC_Header_RemainingUnprocessedData.bin"
'         log "Saving unprocessed data to '" & overlayFileName.FileName & "'"
'
'         Dim overlaydata As New FileStream
'         With overlaydata
'            .Create overlayFileName.FileName, True, False, False
'            .data = File.FixedString(-1)
'            .CloseFile
'         End With
'
'      End If
      
      
 '     .CloseFile
   End With

      
      
      Debug.Assert (Php_Flags And &H2000) = 0
      '...now there is something to do... but -Branch not implemented - yet
      









'Interpreting ByteCode
LogStage 9, "Interpreting IC_body/ByteCode"
   
   CmdParams.EOS = False
   CmdXorKey.EOS = False
   byteCmds.EOS = False

   
 ' create DecompilerObject
 '  Dim IC_Decompile As New IC_Decompile
 ' damn it that used IC_zend_op struct(user defined Type) blocks me from making a class from it
 
 ' you shoould call that before using Decompile on a new function - it resets some value
   DecompilerInit
 
   
   With byteCmds
   'xor byte with rnd

   Dim c As IC_zend_op
   
   Dim cmds As New Collection
   
   Do Until .EOS
      
      
      Dim cmd&, cmd_Enc&, cmd_Flags&, cmd_LineNo&
      cmd_Enc = .int8
      cmd_Flags = .int8
      
      If IsObfuLineNos = False Then
         cmd_LineNo = .int16
      End If
      
     'For php5
      If IsPhp50 And (cmd_Enc = &H95) Then
         CmdXorKey.int8 = 0
         CmdXorKey.Move -1
      End If
      
      
      If (Php_Flags& And &H80) Then
         cmd = cmd_Enc Xor CmdXorKey.int8
      Else
         cmd = cmd_Enc
      End If
      
      c.a_opcode = cmd
      c.g_lineno = cmd_LineNo
      
      On Error Resume Next
      Dim icByteCodeName$
      icByteCodeName = "UNKNOWN Command"
      icByteCodeName = icByteCodeNames(cmd)
      
      log_verbose2 "#" & H16(cmds.Count) & ":" & _
         " 0x" & H8(cmd) & " " & icByteCodeName & _
         "   LineNo: 0x" & H16(cmd_LineNo) & _
         "  Flags: 0x" & H8(cmd_Flags) & _
         "(" & ShowFlagsSimple(cmd_Flags, "|", _
               "Result", "Op1", "Op2", "ExtValue", "ExtValueExtraData") & _
         ")"
      
'         "(Encry: 0x" & H8(cmd_Enc) & ")" & _

      Debug.Assert (cmd_Flags And &HE0) = 0
      '^Unusual/unknown flags set
      

'Type1  to Local.6
      ProcessType CmdParams, c.c_result, _
                  cmd_Flags, BITMASK_CMD_00000001_HAS_RESULT, _
                  IsPhp50, "result: ", True

'Type2
      ProcessType CmdParams, c.c_op1, _
                  cmd_Flags, BITMASK_CMD_00000010_HAS_OP1, _
                  IsPhp50, "op1   : "
'Type4
      ProcessType CmdParams, c.c_op2, _
                  cmd_Flags, BITMASK_CMD_00000100_HAS_OP2, _
                  IsPhp50, "op2   : "
   
  '    Arg_11 , Arg_12
   Dim extended_value As Variant

      Select Case cmd_Flags And BITMASK_CMD_00011000_EXTENDED_VALUE ' Mask these two bit's 00011000
      Case &H0  '00000000
'         extended_value = Empty
         
      Case &H8  '00001000
         extended_value = 1
         log_verbose2 "        extended_value(from Flag): " & H16(extended_value)
         
      Case &H10 '00010000
         extended_value = 60  '<' or 0x3c
         log_verbose2 "        extended_value(from Flag): " & H16(extended_value)
         
      Case &H18 '00011000
         
         extended_value = .int16
         log_verbose2 "        extended_value(from File): " & H16(extended_value)
         
         If IsObfuLineNos Then
            
         Else
            Dim extended_value_LineNr
            extended_value_LineNr = .int16
            If extended_value_LineNr <> 0 Then
               myStop
              '^ It's really unusal that extended_value has a linenumber...
               
               log_verbose2 "extended_value_LineNr: " & H16(extended_value_LineNr)
            End If
         End If
     
      Debug.Assert (cmd_Flags And BITMASK_CMD_11100000_UNKNOWNFLAGS) = 0




      End Select
      
'      If IsPhp50 Then
'
'         On Error Resume Next
'         Dim tmp$
'
'         Select Case cmd
'           'ZEND_JMP
'            Case &H2A '"*"
''               c.c_op1.c_Object = c.c_op1.c_Object
'
'               tmp = c.c_op1.c_Object
'
'             ' On that line is skipped due to 'On Error Resume Next'
'               tmp = H32(tmp)
'
'               log_verbose2 "         CMD_Op1 referes to command[" & tmp & "]"
'           'ZEND_JMPZ,    ZEND_JMPNZ,
'           '(ZEND_JMPZNZ <- &H2D)
'           'ZEND_JMPZ_EX, ZEND_JMPNZ_EX
'            Case &H2B, &H2C, &H2E, &H2F ' ('+'),(','),('.'),('/')
' '              c.c_op2.c_Object = c.c_op2.c_Object
'               tmp = c.c_op2.c_Object
'
'             ' On that line is skipped due to 'On Error Resume Next'
'               tmp = H32(tmp)
'
'               log_verbose2 "         CMD_Op2 referes to command[" & tmp & "]"
'
'
'         End Select
'      End If
      
      
      '  extended_value64 = 0
      ' Special treatment if LongMode (Php_Flags and &h800)
      'if ( IsLong_ )
      '{
      '  *(CmdData + 13) = 0;
      '} else {
      '*(CmdData + 13) = cmdFlag >> 16;
      '  if ( cmdFlag >> 16 == 0xFFFF )
      '  {
      '    CmdCount++;
      '    (_WORD *)CmdTbl_Int16++;
      '    (DWORD *)CmdTbl_Int32++;
      '    *(CmdData + 13) = *CmdTbl_Int32;
      '  }
      ' }


      'ZEND_FE_FETCH, 78
      If cmd = &H4E Then
         extended_value = extended_value Or 2
         log_verbose2 "        Command is 4e - Added Flag '2' to extended_value: " & H16(extended_value)
'         myStop
      End If
      
      c.f_extended_value = extended_value
      
' >>> Feed byteCode data into the decompiler  <<<
      If (cmd_Flags <> 0) Then IC_Decompile.interpret_Cmd c, cmds.Count, CmdCount
    
    ' _____________________________________________________
      log_verbose2 String(70, "_")

    ' Save Command & its data for later
'      cmds.Add c  <- don't work :( stupid VB-limitation
      cmds.Add c.a_opcode '<-well that's not what it should be - but as long I only (mis)use this collection as command counter that is fine
      
      
   Loop

   If Stack.ESP <> 0 Then
      log "ERR: Decompiling this function is finish but there's still data on the decompile stack ->Dumping that data now..."
'      log_decompiled ";==== Remaining data from decompile Stack ======"
      Do While Stack.ESP > 0
         DecompiledOutPut Stack.pop
      Loop
   End If

'tmp = cmdsSize * 780903145
'if Key <> shr32(tmp ,31) + shr32(tmp ,4)
'   "op extract failed"

'Key
End With
   
   
End Sub

Sub ProcessType(ByRef CmdParams As StringReader, ByRef Out_Table As IC_CmdAttrib, _
                cmd_Flags&, TestForFlag&, _
                IsPhp5 As Boolean, Optional OutPutAdd$, Optional DisableStringLookup As Boolean)
                
   With Out_Table
      If cmd_Flags And TestForFlag Then
         'CmdParams[1] CmdParams[2]  CmdParams[3]  CmdParams[4]
         'index     Start      Length     Flags
          .a_op_type = CmdParams.int32
          .c_Object = CmdParams.int32
          .d_ObjectSize = CmdParams.int32
         If IsPhp5 Then
            .e_FlagsPHP5 = CmdParams.int32

'Other values are unusal
Debug.Assert ((.e_FlagsPHP5 = 0) Or _
             (.e_FlagsPHP5 = 2))
            
            
            
            
         End If
          .e_Obj_Type = CmdParams.int8
          .is_ref = CmdParams.int8
          .refcount = CmdParams.int16

'          If (.is_ref <> 0) And (.refcount <> 0) Then
'          .is_ref = 0
'          .refcount = 1
 '         End If
          
         On Error Resume Next
         Dim op_type$
         op_type = Switch(.a_op_type = IS_CONST__1, "CONST  ", _
                          .a_op_type = IS_TMP_VAR__2, "TMP_VAR", _
                          .a_op_type = IS_VAR__4, "VAR    ", _
                          .a_op_type = IS_UNUSED__8, "UNUSED ")
         If Err Then log "Unexpected op_type!"
                          
         On Error Resume Next
         Dim obj_type$
'        obj_type = Switch(.e_Obj_Type = IS_NULL__0,           "NULL  ", _
                          .e_Obj_Type = IS_LONG__1,           "LONG  ", _
                          .e_Obj_Type = IS_DOUBLE__2,         "DOUBLE", _
                          .e_Obj_Type = IS_STRING__3,         "STRING", _
                          .e_Obj_Type = IS_ARRAY__4,          "ARRAY ", _
                          .e_Obj_Type = IS_OBJECT__5,         "OBJECT", _
                          .e_Obj_Type = IS_BOOL__6,           "BOOL  ", _
                          .e_Obj_Type = IS_RESOURCE__7,       "RESOURCE", _
                          .e_Obj_Type = IS_CONSTANT__8,       "CONST ", _
                          .e_Obj_Type = IS_CONSTANT_ARRAY__9, "CONST_ARRAY")
         
         obj_type = Switch(.e_Obj_Type = IS_NULL__0, "NULL  ", _
                          .e_Obj_Type = IS_LONG__1, "LONG  ", _
                          .e_Obj_Type = IS_DOUBLE__2, "DOUBLE", _
                          .e_Obj_Type = IS_STRING__3, "STRING", _
                          .e_Obj_Type = IS_ARRAY__4, "ARRAY ", _
                          .e_Obj_Type = IS_OBJECT__5, "OBJECT", _
                          .e_Obj_Type = IS_BOOL__6, "BOOL  ", _
                          .e_Obj_Type = IS_RESOURCE__7, "RESOURCE", _
                          .e_Obj_Type = IS_CONSTANT__8, "CONST ", _
                          .e_Obj_Type = IS_CONSTANT_ARRAY__9, "CONST_ARRAY")
         If Err Then log "Unexpected obj_type!"

         
         log_verbose2 OutPutAdd & _
                     H8(.a_op_type) & "->" & op_type & " " & _
                     H32(.c_Object) & " " & H32(.d_ObjectSize) & "  " & _
         IIf(IsPhp5, H8(.e_FlagsPHP5) & "  ", "") & _
                     H8(.e_Obj_Type) & "->" & obj_type & " " & _
                     H8(.is_ref) & "->" & IIf(.is_ref = ZEND_RETURN_VAL__0, "val", "ref") & " " & _
                     H16(.refcount)
                     
         If (.a_op_type = IS_CONST__1) And (DisableStringLookup = False) Then
            
            Dim KeyWord$
            KeyWord = DeSerialiseString(.c_Object, .d_ObjectSize, .e_Obj_Type)
            If KeyWord <> "" Then
               log_verbose2 "        KeyWord: " & KeyWord
               
              'Overwrite Offset with String
               .c_Object = KeyWord

            End If
            
            
         End If
      
      Else
         'Move this from IC.dll '.data' to EBP
         '0x00000008 (0x00000000)
         '0x00000000 0x00000000 0x00000000 (0x00000000)
         .a_op_type = IS_UNUSED__8
      End If
      
   End With

End Sub

 


'Serializing Data Serializing means the transformation of variables
'to a byte-code representation that can be stored anywhere as a normal string.
Private Function DeSerialiseString$(goto_pos, len_, obj_type As obj_type)
  
   
'T2 [Elementsize=0x10]:
'index      Start        Length   Flags
'00000001   00000002   00000051   00020103
'00000001   00000001   BFFF8F18   00020101
   'T2[1]      T2[2]       T2[3]     T2[4]
   
   
   Dim StringOut As New StringReader
   With StringOut
   
      Select Case obj_type
        'Case 0, 1, 2, 6
            'Do Nothing
        
        Case IS_NULL__0
        Case IS_LONG__1
        Case IS_DOUBLE__2
         'convert Float to String

            Dim FloatConv As New StringReader
            With FloatConv

               .DisableAutoMove = False
               .EOS = False
               .int32 = goto_pos
               .int32 = len_
               
               .EOS = False
               .DisableAutoMove = True
               StringOut.Data = Replace(.GetFloat64, ",", ".")
            End With
            
        
        
        Case IS_STRING__3, IS_CONSTANT__8
           
         If len_ = 0 Then
             
             .Data = "" 'EmptyStr
             
         Else
         
      ' Is custom string, current FileName or internal command
        If goto_pos >= 0 Then
         ' Get from customStringTable
               
             .Data = Mid(StringsData.Data, goto_pos + 1, len_)
             If .int8 = &HD Then
                 .Data = "[Obfuscated]" & ValuesToHexString(StringOut)
             Else
                 .Move -1
             End If
             
        ElseIf goto_pos = -1 Then
            .Data = "[FullPath To CurrentPhpFile]  (Initial Len was: " & len_ & ")"
        
        Else
         ' GetString from KnownFunctionTable
           .Data = icKnownFunctions(-goto_pos)
           
           If Len(StringOut) <> len_ Then
              MsgBox "Error in icKnownFunctions: expected Len: " & len_ & " - RealLen: " & Len(StringOut)
                         End If
           
        End If
         End If
        
        Case IS_CONSTANT_ARRAY__9
           
'           myStop
         ' TODO: Not implemented ('cause not used so far)
          .Data = Mid(StringsData.Data, goto_pos + 1, len_)
          .Data = DecodeBasicTypes_rek(StringOut)
           
        Case Else
            Err.Raise vbObjectError, , "ERROR!!! Can't deserialise zval type " & obj_type
           myStop
   
   
      End Select
   
      DeSerialiseString = .Data
      
   End With
   
End Function



Sub log(Text$)
   FrmMain.log Text$
End Sub

Sub FL(Text$)
   FrmMain.FL Text$
End Sub

Sub FL_verbose(Text$)
   FrmMain.FL_verbose Text$
End Sub

Sub FL_verbose2(Text$)
   FrmMain.FL Text$
End Sub

