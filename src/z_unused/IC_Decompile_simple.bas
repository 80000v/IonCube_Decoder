Attribute VB_Name = "IC_DecompileFuncs"

Option Explicit
Dim Stack As Stack

   Dim Prefix$, InFix$, PostFix$
   
   Dim DecompileLine As New clsStrCat
      
   Private c As IC_zend_op
   Private cmd As Byte
   Private Result As IC_CmdAttrib
   Private op1 As IC_CmdAttrib
   Private op2 As IC_CmdAttrib

Sub PushResult(DataToStore)
   Select Case Result.a_op_type
      
      Case IS_VAR__4
'quickHack1
IS_VAR:
         If Result.d_ObjectSize = 1 Then
            DecompileLine.Concat CStr(DataToStore)
         Else
            Stack.push CStr(DataToStore), Result.c_Object
         End If
               
      Case IS_TMP_VAR__2

'quickHack1
If tmp_valueType = 0 Then GoTo IS_VAR
         

         tmp_values(CStr(Result.c_Object)).Add DataToStore
         
      Case Else
         Stop
   End Select
End Sub
Private Function ProcessData()
With Stack
   Select Case Result.a_op_type
      Case IS_CONST__1
         Stop 'Probably a Error
         
      Case IS_TMP_VAR__2
         
         Dim TmpVarID
         TmpVarID = Result.c_Object
         
         Select Case op1.a_op_type
            Case IS_CONST__1
            Case IS_TMP_VAR__2
               .push .pop, TmpVarID
               
            Case IS_VAR__4
            Case IS_UNUSED__8
         End Select
      
         Select Case op2.a_op_type
            Case IS_CONST__1
            Case IS_TMP_VAR__2
            Case IS_VAR__4
            Case IS_UNUSED__8
         End Select
      
      
      Case IS_VAR__4
         'Normal processing
         
      Case IS_UNUSED__8
         'No return var Output
         
   End Select
End With
End Function
   

Private Function GetObjData_simple(c As IC_CmdAttrib, _
                  Optional isCommand As Boolean, _
                  Optional isVar As Boolean) As Variant
   
   Select Case c.a_op_type
      
      
      Case IS_CONST__1
         
         'Stack=0
         Select Case c.e_Obj_Type
            Case IS_LONG__1, IS_DOUBLE__2
               If isCommand Then
                  Err.Raise vbObjectError, , "Logic Error keyword expected - a literal(Number) found."
                  myStop
               Else
                  GetObjData = c.c_Object
               End If
            
            Case IS_STRING__3
               If isCommand Then
                  GetObjData = c.c_Object
                  
               ElseIf isVar Then
                  GetObjData = "$" & c.c_Object
               Else
                  GetObjData = MakeStr(c.c_Object)
               End If
               
            Case IS_NULL__0
               GetObjData = Null
               
               
            Case IS_CONSTANT__8
'               myStop
               GetObjData = c.c_Object
            
            Case IS_CONSTANT_ARRAY__9
'               myStop
               GetObjData = c.c_Object
               
            Case Else
               myStop
               
               
            End Select
            
         
            
      Case IS_TMP_VAR__2
         'concat TmpData according to type
         Select Case tmp_valueType
            
           'Dirty Quick Hack1 - Till theres a better solution
           'Makes TMP_VAR->Out  to behave like IS_CONST
            Case 0
               GoTo IS_VAR_TmpHack
               
            Case IS_ARRAY__4
               GetObjData = "array(" & joinCol(tmp_values(c.c_Object), ", ") & ")"
               tmp_valueType = 0 'Dirty Quick Hack1
               
            Case IS_STRING__3
               GetObjData = MakeStr(joinCol(tmp_values(c.c_Object), ""))
               tmp_valueType = 0 'Dirty Quick Hack1
               
            Case Else
               myStop
         End Select
      
      Case IS_VAR__4
IS_VAR_TmpHack: 'Dirty Quick Hack1

         'Stack=1
         GetObjData = Stack.pop(c.c_Object)
         
      Case IS_UNUSED__8
         GetObjData = c.c_Object

'         GetObjData = "<Unused>"
         
         
      Case Else
         myStop
         
   End Select
End Function



Sub ic_interpret_simple(cmdObj As IC_zend_op, CmdIndex&, CmdCount&)

 ' Do Init
   c = cmdObj
   
   cmd = c.a_opcode And &HFF
   Result = c.c_result
   op1 = c.c_op1
   op2 = c.c_op2


   
   Select Case cmd

      Case ZEND_NOP                       ' &h00  0
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ADD                       ' &h01  1
         Prefix = "": InFix = "+": PostFix = ""
         myStop

      Case ZEND_SUB                       ' &h02  2
         Prefix = "": InFix = "-": PostFix = ""
         myStop

      Case ZEND_MUL                       ' &h03  3
         Prefix = "": InFix = "*": PostFix = ""
         myStop

      Case ZEND_DIV                       ' &h04  4
         Prefix = "": InFix = "/": PostFix = ""
         myStop

      Case ZEND_MOD                       ' &h05  5
         Prefix = "": InFix = "%": PostFix = ""
         myStop

      Case ZEND_SL                        ' &h06  6
         Prefix = "": InFix = "<<": PostFix = ""
         myStop

      Case ZEND_SR                        ' &h07  7
         Prefix = "": InFix = ">>": PostFix = ""
         myStop

      Case ZEND_CONCAT                    ' &h08  8
         Prefix = "": InFix = ".": PostFix = ""
         myStop

      Case ZEND_BW_OR                     ' &h09  9
         Prefix = "": InFix = "|": PostFix = ""
         myStop

      Case ZEND_BW_AND                    ' &h0A  10
         Prefix = "": InFix = "%": PostFix = ""
         myStop

      Case ZEND_BW_XOR                    ' &h0B  11
         Prefix = "": InFix = "%": PostFix = ""
         myStop

      Case ZEND_BW_NOT                    ' &h0C  12
         Prefix = "": InFix = "!": PostFix = ""
         myStop

      Case ZEND_BOOL_NOT                  ' &h0D  13
         Prefix = "": InFix = "!!": PostFix = ""
         myStop

      Case ZEND_BOOL_XOR                  ' &h0E  14
         Prefix = "": InFix = "^": PostFix = ""
         myStop

      Case ZEND_IS_IDENTICAL              ' &h0F  15
         Prefix = "": InFix = "===": PostFix = ""
         myStop

      Case ZEND_IS_NOT_IDENTICAL          ' &h10  16
         Prefix = "": InFix = "!==": PostFix = ""
         myStop

      Case ZEND_IS_EQUAL                  ' &h11  17
         Prefix = "": InFix = "==": PostFix = ""
         myStop

      Case ZEND_IS_NOT_EQUAL              ' &h12  18
         Prefix = "": InFix = "!=": PostFix = ""
         myStop

      Case ZEND_IS_SMALLER                ' &h13  19
         Prefix = "": InFix = "<": PostFix = ""
         myStop

      Case ZEND_IS_SMALLER_OR_EQUAL       ' &h14  20
         Prefix = "": InFix = "<=": PostFix = ""
         myStop

      Case ZEND_CAST                      ' &h15  21
         Prefix = "": InFix = "?CAST?": PostFix = ""
         myStop

      Case ZEND_QM_ASSIGN                 ' &h16  22
         Prefix = "": InFix = ":": PostFix = ""
         myStop

      Case ZEND_ASSIGN_ADD                ' &h17  23
         Prefix = "": InFix = "+=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_SUB                ' &h18  24
         Prefix = "": InFix = "-=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_MUL                ' &h19  25
         Prefix = "": InFix = "*=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_DIV                ' &h1A  26
         Prefix = "": InFix = "\=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_MOD                ' &h1B  27
         Prefix = "": InFix = "%=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_SL                 ' &h1C  28
         Prefix = "": InFix = "<<=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_SR                 ' &h1D  29
         Prefix = "": InFix = ">>=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_CONCAT             ' &h1E  30
         Prefix = "": InFix = ".=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_BW_OR              ' &h1F  31
         Prefix = "": InFix = "|=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_BW_AND             ' &h20  32
         Prefix = "": InFix = "%=": PostFix = ""
         myStop

      Case ZEND_ASSIGN_BW_XOR             ' &h21  33
         Prefix = "": InFix = "^=": PostFix = ""
         myStop

      Case ZEND_PRE_INC                   ' &h22  34
         Prefix = "": InFix = "++": PostFix = ""
         myStop

      Case ZEND_PRE_DEC                   ' &h23  35
         Prefix = "": InFix = "--": PostFix = ""
         myStop

      Case ZEND_POST_INC                  ' &h24  36
         Prefix = "": InFix = "++": PostFix = ""
         myStop

      Case ZEND_POST_DEC                  ' &h25  37
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN                    ' &h26  38
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_REF                ' &h27  39
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ECHO                      ' &h28  40
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_PRINT                     ' &h29  41
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMP                       ' &h2A  42
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMPZ                      ' &h2B  43
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMPNZ                     ' &h2C  44
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMPZNZ                    ' &h2D  45
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMPZ_EX                   ' &h2E  46
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMPNZ_EX                  ' &h2F  47
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_CASE                      ' &h30  48
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_SWITCH_FREE               ' &h31  49
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_BRK                       ' &h32  50
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_CONT                      ' &h33  51
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_BOOL                      ' &h34  52
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INIT_STRING               ' &h35  53
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ADD_CHAR                  ' &h36  54
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ADD_STRING                ' &h37  55
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ADD_VAR                   ' &h38  56
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_BEGIN_SILENCE             ' &h39  57
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_END_SILENCE               ' &h3A  58
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INIT_FCALL_BY_NAME        ' &h3B  59
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DO_FCALL                  ' &h3C  60
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DO_FCALL_BY_NAME          ' &h3D  61
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_RETURN                    ' &h3E  62
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_RECV                      ' &h3F  63
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_RECV_INIT                 ' &h40  64
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_SEND_VAL                  ' &h41  65
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_SEND_VAR                  ' &h42  66
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_SEND_REF                  ' &h43  67
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_NEW                       ' &h44  68
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMP_NO_CTOR               ' &h45  69
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FREE                      ' &h46  70
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INIT_ARRAY                ' &h47  71
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ADD_ARRAY_ELEMENT         ' &h48  72
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INCLUDE_OR_EVAL           ' &h49  73
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_UNSET_VAR                 ' &h4A  74
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_UNSET_DIM                 ' &h4B  75
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_UNSET_OBJ                 ' &h4C  76
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FE_RESET                  ' &h4D  77
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FE_FETCH                  ' &h4E  78
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_EXIT                      ' &h4F  79
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_R                   ' &h50  80
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_R               ' &h51  81
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_OBJ_R               ' &h52  82
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_W                   ' &h53  83
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_W               ' &h54  84
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_OBJ_W               ' &h55  85
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_RW                  ' &h56  86
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_RW              ' &h57  87
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_OBJ_RW              ' &h58  88
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_IS                  ' &h59  89
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_IS              ' &h5A  90
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_OBJ_IS              ' &h5B  91
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_FUNC_ARG            ' &h5C  92
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_FUNC_ARG        ' &h5D  93
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_OBJ_FUNC_ARG        ' &h5E  94
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_UNSET               ' &h5F  95
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_UNSET           ' &h60  96
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_OBJ_UNSET           ' &h61  97
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_DIM_TMP_VAR         ' &h62  98
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_CONSTANT            ' &h63  99
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DECLARE_FUNCTION_OR_CLASS ' &h64  100
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_EXT_STMT                  ' &h65  101
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_EXT_FCALL_BEGIN           ' &h66  102
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_EXT_FCALL_END             ' &h67  103
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_EXT_NOP                   ' &h68  104
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_TICKS                     ' &h69  105
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_SEND_VAR_NO_REF           ' &h6A  106
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_CATCH                     ' &h6B  107
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_THROW                     ' &h6C  108
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_FETCH_CLASS               ' &h6D  109
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_CLONE                     ' &h6E  110
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INIT_CTOR_CALL            ' &h6F  111
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INIT_METHOD_CALL          ' &h70  112
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INIT_STATIC_METHOD_CALL   ' &h71  113
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ISSET_ISEMPTY_VAR         ' &h72  114
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ISSET_ISEMPTY_DIM_OBJ     ' &h73  115
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_IMPORT_FUNCTION           ' &h74  116
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_IMPORT_CLASS              ' &h75  117
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_IMPORT_CONST              ' &h76  118
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_OP_119                    ' &h77  119
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_OP_120                    ' &h78  120
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_ADD_OBJ            ' &h79  121
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_SUB_OBJ            ' &h7A  122
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_MUL_OBJ            ' &h7B  123
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_DIV_OBJ            ' &h7C  124
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_MOD_OBJ            ' &h7D  125
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_SL_OBJ             ' &h7E  126
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_SR_OBJ             ' &h7F  127
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_CONCAT_OBJ         ' &h80  128
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_BW_OR_OBJ          ' &h81  129
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_BW_AND_OBJ         ' &h82  130
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_BW_XOR_OBJ         ' &h83  131
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_PRE_INC_OBJ               ' &h84  132
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_PRE_DEC_OBJ               ' &h85  133
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_POST_INC_OBJ              ' &h86  134
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_POST_DEC_OBJ              ' &h87  135
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_OBJ                ' &h88  136
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_OP_DATA                   ' &h89  137
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_INSTANCEOF                ' &h8A  138
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DECLARE_CLASS             ' &h8B  139
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DECLARE_INHERITED_CLASS   ' &h8C  140
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DECLARE_FUNCTION          ' &h8D  141
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_RAISE_ABSTRACT_ERROR      ' &h8E  142
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_START_NAMESPACE           ' &h8F  143
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ADD_INTERFACE             ' &h90  144
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_VERIFY_INSTANCEOF         ' &h91  145
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_VERIFY_ABSTRACT_CLASS     ' &h92  146
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ASSIGN_DIM                ' &h93  147
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_ISSET_ISEMPTY_PROP_OBJ    ' &h94  148
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_HANDLE_EXCEPTION          ' &h95  149
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_USER_OPCODE               ' &h96  150
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_U_NORMALIZE               ' &h97  151
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_JMP_SET                   ' &h98  152
         Prefix = "": InFix = "": PostFix = ""
         myStop

      Case ZEND_DECLARE_LAMBDA_FUNCTION   ' &h99  153
         Prefix = "": InFix = "": PostFix = ""
         myStop

      log_verbose2 "ESP:" & .ESP & "   '" & .PreviewPop & "'"
      log_verbose2 "DecompBuff: " & DecompileLine.value

      If Stack.ESP = 0 Then
         DecompiledOutPut DecompileLine
      End If
   End Select
End Sub
