Attribute VB_Name = "IC_Decompile"
Option Explicit

Const UNKNOWN_OPCODE$ = "Unknown Opcode"

Public Dumper_Data


Private DecompileLine As New clsStrCat

Private FuncName$

Private TmpFuncArgsIsRef As New Stack

Private FuncArgsCall As New Stack

Private FuncArgsIsRef As New Collection
Private FuncArgsNames As New Collection

Private FuncDefNames As New Collection
Private FuncDefIsRef As New Collection


Public Stack As New Stack

Private BranchesStack As New Stack

'Private CommandHistory As New Collection

 ' Note: These modulvars are intialise by interpret_Cmd()
   Private c As IC_zend_op
   Private cmd As Byte
   Private Result As IC_CmdAttrib
   Private op1 As IC_CmdAttrib
   Private op2 As IC_CmdAttrib

   Dim DisamTxt$
   Dim DisamTxtComment$
   
   Dim NestLevel&

Private isInFunctionArgInit As Boolean

Public tmp_values As New Collection
Public tmp_valueType

Private Dumper_Line As New clsStrCat
      Private LValueWithIgnoreFlag As Boolean
  
Sub DecompilerInit()
   tmp_valueType = Empty
   
   NestLevel = 0
   Set TmpFuncArgsIsRef = New Stack
   
   Set FuncArgsCall = New Stack
   
   Set FuncArgsIsRef = New Collection
   Set FuncArgsNames = New Collection
   
   Set FuncDefNames = New Collection
   Set FuncDefIsRef = New Collection
   
   
   Set Stack = New Stack
   
   Set BranchesStack = New Stack
      
   
   Set Dumper_Line = New clsStrCat
   LValueWithIgnoreFlag = True
'   isInFunctionArgInit = False
   
End Sub


Sub Tmp_Prepare(tmp_Var_Type&)
     tmp_valueType = tmp_Var_Type
   
   ' clear old
     Set tmp_values = New Collection
     tmp_values.Add New Collection, CStr(Result.c_Object)
   
End Sub

Sub SelectFunction(myFuncName$)
   On Error Resume Next
   Set FuncArgsNames = FuncDefNames(myFuncName)
   If Err Then
    
    ' Create New
      Set FuncArgsNames = New Collection
      FuncDefNames.Add FuncArgsNames, myFuncName
   
   End If
   
   FuncName = myFuncName
End Sub

'Function Arg_PureName$(Arg$)
'   If Arg_IsRef(Arg) Then
'    ' cur '&' at the beginning
'      Arg_PureName = Mid(1, Arg)
'   Else
'      Arg_PureName = Arg
'   End If
'End Function


Sub CallFunctionArg(ArgName$, ByVal ArgsCount&)
   SelectFunction ArgName
   
   Debug.Assert ArgsCount = FuncArgsIsRef.Count
   Do While TmpFuncArgsIsRef.ESP > 0
      Dim IsRef As Boolean
      TmpFuncArgsIsRef.pop
'      SetFunctionArg ArgsCount, Arg_MakeRef
   Loop
   
End Sub



Sub SetFunctionArg(ArgNr&, isPassedByRef As Boolean, Optional ArgName$ = "")
  
   On Error Resume Next
   
   Dim isPassedByRef_Cur As Boolean
   isPassedByRef_Cur = FuncArgsIsRef(ArgNr)
   If Err Then
    ' Okay Arg doesn't exists create new Arg
      FuncArgsIsRef.Add isPassedByRef, ArgNr
   Else
   
   ' Compared Passed byRef or Passed byVal
     If isPassedByRef_Cur Then
        If isPassedByRef Then
           'Okay both are Ref
        Else
           'Okay 'Upgrade' Arg_Cur to Ref
            Stop
            'FuncArgs(ArgNr) = ArgName_New
        End If
        
     Else
        If isPassedByRef Then
           'Error 'Downgrade' from Arg_Ref to Arg_Val
           Stop
        Else
           'Okay both are not Ref
        End If
        
     End If
   
   
   If ArgName = "" Then Exit Sub
   
   Dim ArgName_Cur$
   ArgName_Cur = FuncArgsNames(ArgNr)
   If Err Then
    ' Okay Arg doesn't exists create new Arg
      FuncArgsNames.Add ArgName, ArgNr
   Else
    
    ' Check current Arg Names
      If ArgName <> ArgName_Cur Then
         Stop ' - different Names
         Err.Raise vbObjectError, , "Error: ArgumentNames different"
   
      End If
    
   End If
      
   
   End If
End Sub

Function GetFunctionArgList$(Optional ListSeperator$ = ",")
   
'   On Error Resume Next
   Dim tmp As New Collection
   Dim i&
'   For i = LBound(FuncArgsIsRef) To UBound(FuncArgsIsRef)
'      Tmp.Add GetFunctionArg(i)
'   Next
'   GetFunctionArgList = Join(Tmp, ListSeperator)
   
   
End Function


Function GetFunctionArg$(ArgNr&)
   
   On Error Resume Next
   GetFunctionArg = FuncArgsNames(ArgNr)
   
   If FuncArgsIsRef(ArgNr) Then
      GetFunctionArg = Arg_MakeRef(GetFunctionArg)
   End If
   
End Function

Sub SetFunctionName(FuncName$, Optional Scope$ = "")
    
 ' Select Function ArgList - if there is none create a new one
'   SelectFunction FuncName

   
   isInFunctionArgInit = True
   
   DecompileLine.Concat Join(Array(Scope, "function", FuncName))

End Sub



Function Arg_MakeRef$(Varname$)
   Arg_MakeRef = "&" & Varname
End Function


   '======== Operators  opcodes 0x1 .. 0x25 ==================
Private Sub ProcessOperator()
   
   Select Case cmd

      Case ZEND_ADD                       ' &h01  1
         Dim OpSymbol$
         OpSymbol = "+"
         
binaryInFixOperator:
          PushResult str(op1) & OpSymbol & str(op2)


      Case ZEND_SUB                       ' &h02  2
         OpSymbol = "-": GoTo binaryInFixOperator

      Case ZEND_MUL                       ' &h03  3
         OpSymbol = "*": GoTo binaryInFixOperator

      Case ZEND_DIV                       ' &h04  4
         OpSymbol = "/": GoTo binaryInFixOperator

      Case ZEND_MOD                       ' &h05  5
         OpSymbol = "%": GoTo binaryInFixOperator

      Case ZEND_SL                        ' &h06  6
         OpSymbol = "<<": GoTo binaryInFixOperator

      Case ZEND_SR                        ' &h07  7
         OpSymbol = ">>": GoTo binaryInFixOperator

      Case ZEND_CONCAT                    ' &h08  8
         OpSymbol = ".": GoTo binaryInFixOperator

      Case ZEND_BW_OR                     ' &h09  9
         OpSymbol = "|": GoTo binaryInFixOperator

      Case ZEND_BW_AND                    ' &h0A  10
         OpSymbol = "%": GoTo binaryInFixOperator

      Case ZEND_BW_XOR                    ' &h0B  11
         OpSymbol = "%": GoTo binaryInFixOperator

      Case ZEND_BW_NOT                    ' &h0C  12
         OpSymbol = "!": GoTo UnaryPreOperator
myStop
      Case ZEND_BOOL_NOT                  ' &h0D  13
         OpSymbol = "!!": GoTo UnaryPreOperator
myStop
      Case ZEND_BOOL_XOR                  ' &h0E  14
         OpSymbol = "^": GoTo binaryInFixOperator

      Case ZEND_IS_IDENTICAL              ' &h0F  15
         OpSymbol = "===": GoTo binaryInFixOperator

      Case ZEND_IS_NOT_IDENTICAL          ' &h10  16
         OpSymbol = "!==": GoTo binaryInFixOperator

      Case ZEND_IS_EQUAL                  ' &h11  17
         OpSymbol = "==": GoTo binaryInFixOperator

      Case ZEND_IS_NOT_EQUAL              ' &h12  18
         OpSymbol = "!=": GoTo binaryInFixOperator

      Case ZEND_IS_SMALLER                ' &h13  19
         OpSymbol = "<": GoTo binaryInFixOperator

      Case ZEND_IS_SMALLER_OR_EQUAL       ' &h14  20
         OpSymbol = "<=": GoTo binaryInFixOperator

      Case ZEND_CAST                      ' &h15  21
'         OpSymbol = "?CAST?": goto binaryInFixOperator
         Select Case c.f_extended_value
            Case IS_NULL__0:
               OpSymbol = "null"
            Case IS_LONG__1:
               OpSymbol = "long"
            Case IS_DOUBLE__2:
               OpSymbol = "float"
            Case IS_STRING__3:
               OpSymbol = "string"
            Case IS_ARRAY__4:
               OpSymbol = "array"
            Case IS_OBJECT__5:
               OpSymbol = "object"
            Case IS_BOOL__6:
               OpSymbol = "boolean"
            Case Else
               Stop
         End Select
         
         PushResult "(" & OpSymbol & ")" & Var(op1)

myStop

      Case ZEND_QM_ASSIGN                 ' &h16  22
'         Tmp_Prepare IS_BOOL__6
         OpSymbol = ":" ': GoTo UnaryPostOperator
         PushResult str(op1) & OpSymbol

myStop

      Case ZEND_ASSIGN_ADD                ' &h17  23
         OpSymbol = "+=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_SUB                ' &h18  24
         OpSymbol = "-=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_MUL                ' &h19  25
         OpSymbol = "*=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_DIV                ' &h1A  26
         OpSymbol = "\=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_MOD                ' &h1B  27
         OpSymbol = "%=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_SL                 ' &h1C  28
         OpSymbol = "<<=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_SR                 ' &h1D  29
         OpSymbol = ">>=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_CONCAT             ' &h1E  30
         OpSymbol = ".=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_BW_OR              ' &h1F  31
         OpSymbol = "|=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_BW_AND             ' &h20  32
         OpSymbol = "%=": GoTo binaryInFixOperator

      Case ZEND_ASSIGN_BW_XOR             ' &h21  33
         OpSymbol = "^=": GoTo binaryInFixOperator



      Case ZEND_PRE_INC                   ' &h22  34
         OpSymbol = "++"

UnaryPreOperator:
         PushResult OpSymbol & Var(op1)
         
      Case ZEND_PRE_DEC                   ' &h23  35
         OpSymbol = "--": GoTo UnaryPreOperator

      Case ZEND_POST_INC                  ' &h24  36
         OpSymbol = "++"
      
UnaryPostOperator:
         PushResult Var(op1) & OpSymbol

      Case ZEND_POST_DEC                  ' &h25  37
         OpSymbol = "--": GoTo UnaryPostOperator

   End Select
   
   
End Sub

Sub dataDumper_addLine(lineNr&, Data)
   ArrayEnsureBounds Dumper_Data
   Do While lineNr >= UBound(Dumper_Data)
      ArrayAdd Dumper_Data
   Loop
   
   Dumper_Data(lineNr) = Dumper_Data(lineNr) & " | " & Data
   
End Sub

Sub dataDumper()
   
   
   If op1.is_ref Then
      Dumper_Line.Concat "  " & str(op1)
   End If
   
   If op2.is_ref Then
      Dumper_Line.Concat "  " & str(op2)
   End If
   
   If (c.g_lineno <> 0) Then
      dataDumper_addLine c.g_lineno, Dumper_Line.value
      
   Else
 '   Debug.Assert LValueWithIgnoreFlag = LineNrChanged
   
      If LValueWithIgnoreFlag Then
         ArrayAdd Dumper_Data, Dumper_Line.value
      End If
   
    ' Memorise for next run
      
      LValueWithIgnoreFlag = ((Result.a_op_type = IS_VAR__4) And Result.d_ObjectSize = 1)
      LValueWithIgnoreFlag = LValueWithIgnoreFlag Or (Result.a_op_type = IS_UNUSED__8)
   End If
   
   Dumper_Line.Clear

End Sub

Sub interpret_Cmd(cmdObj As IC_zend_op, CmdIndex&, CmdCount&)

   On Error GoTo interpret_Cmd_err
   
 ' Do Init
   c = cmdObj
   
   cmd = c.a_opcode And &HFF
   Result = c.c_result
   op1 = c.c_op1
   op2 = c.c_op2
   
   
   dataDumper
  
  ' try to be in sync with original LineNr
   Do While c.g_lineno > FrmMain.List_Source.ListCount
      DecompiledOutPut ""
   Loop
   
   
   
'   .push "a5", 2
'   Debug.Print .pop(2)
   
'   RangeCheck cmd, &HCC, 0, "cmd out of range"
   
   
'   DisamTxt = UNKNOWN_OPCODE
   DisamTxtComment = ""
   
   Dim tmpstr$
   
ReDoLoop:
   
   With Stack
      Select Case cmd
      

         Case ZEND_NOP                       ' &h00  0
'         Case &H0, &H65 To &H68 '    00'   .
            myStop


   
   
         Case ZEND_ADD To ZEND_POST_DEC '&H1 To &H25
            ProcessOperator
  
   '=========================================================================

   
         Case ZEND_ASSIGN                    ' &h26  38
            PushResult Var(op1) & "=" & str(op2)
   
         Case ZEND_ASSIGN_REF                ' &h27  39
            myStop
   
         Case ZEND_ECHO                      ' &h28  40
      'http://www.google.com/codesearch/p?hl=de&sa=N&cd=4&ct=rc#OdM6HPk4UX4/php-5.1.4/Zend/zend_vm_def.h&q=zend_print_variable
      'ZEND_VM_HANDLER(40, ZEND_ECHO, CONST|TMP|VAR|CV, ANY)
      '{
      '        zend_op *opline = EX(opline);
      '        zend_free_op free_op1;
      '        zval z_copy;
      '        zval *z = GET_OP1_ZVAL_PTR(BP_VAR_R);
      '
      '        if (Z_TYPE_P(z) == IS_OBJECT && Z_OBJ_HT_P(z)->get_method != NULL &&
      '                zend_std_cast_object_tostring(z, &z_copy, IS_STRING, 0 TSRMLS_CC) == SUCCESS) {
      '                zend_print_variable(&z_copy);
      '                zval_dtor(&z_copy);
      '        } else {
      '                zend_print_variable(z);
      '        }
      '
      '        FREE_OP1();
      '        ZEND_VM_NEXT_OPCODE();
      '}
               
      '            DisamTxt = "ZEND_ECHO" 'zend_print_variable
         
               DecompileLine.Concat "echo " & str(op1)
   
         Case ZEND_PRINT                     ' &h29  41
            .push "print " & str(op1)
   
         Case ZEND_JMP                       ' &h2A  42
             DisamTxtComment = "GoTo Command #" & H16(Num(op1))
               
   '            Dim CmdNr_EndBody$
   '            CmdNr_EndBody = BranchesStack.pop
   '            If CmdNr_EndBody <> (CmdIndex + 1) Then
   '              'Nest down error - that is probably not the correct body
   '               myStop
   '            End If
               
   myStop
   
   
         Case ZEND_JMPZ                      ' &h2B  43
             DisamTxt = "ZEND_JMPZ"
            DisamTxtComment = "Conditional jmp to #" & H32(Num(op2))
            
          ' in case that is a 'For' the 'extended_value'
          ' showing were the body ends
            If IsEmpty(c.f_extended_value) Then
               DecompileLine.Concat "if (" & Key(op1) & ") "
               
            Else
             '... but as long 'for' is not correctly implemented a need to use 'while'
               DecompileLine.Concat "while (" & Key(op1) & ") "
               
             ' --- Unecessary Debug code
               Dim JmpAdr_absolute&
               JmpAdr_absolute = Num(op2)
               
               Dim JmpAdr_relative&
               JmpAdr_relative = c.f_extended_value
   
               
               If ((JmpAdr_relative) + CmdIndex) <> JmpAdr_absolute Then
                ' Unusal 'f_extended_value' - however that is uncritical
                ' since it's I don't use that value - however that may indicate an error
'                  myStop
               End If
              ' ----
   
               
            End If
            
            
            BranchesStack.push Num(op2)
            
               
   
         Case ZEND_JMPNZ                     ' &h2C  44
            myStop
   
         Case ZEND_JMPZNZ                    ' &h2D  45
            myStop
   
         Case ZEND_JMPZ_EX                   ' &h2E  46
            myStop
   
         Case ZEND_JMPNZ_EX                  ' &h2F  47
            myStop
   
         Case ZEND_CASE                      ' &h30  48
            myStop
   
         Case ZEND_SWITCH_FREE               ' &h31  49
            myStop
   
         Case ZEND_BRK                       ' &h32  50
            myStop
   
         Case ZEND_CONT                      ' &h33  51
            myStop
   
         Case ZEND_BOOL                      ' &h34  52
            myStop
   
         Case ZEND_INIT_STRING               ' &h35  53
            Tmp_Prepare IS_STRING__3
'            GoTo ZEND_ADD_STRING
'            myStop
   
         Case ZEND_ADD_CHAR                  ' &h36  54
            myStop
   
         Case ZEND_ADD_STRING                ' &h37  55
'ZEND_ADD_STRING:
            PushResult Key(op2)
            myStop
   
         Case ZEND_ADD_VAR                   ' &h38  56
           
           
           
            PushResult Var(op2)
   
         Case ZEND_BEGIN_SILENCE             ' &h39  57
            myStop
   
         Case ZEND_END_SILENCE               ' &h3A  58
            myStop
   
'=== Functions Calls
   
         Case ZEND_INIT_FCALL_BY_NAME        ' &h3B  59
            If 0 = (c.f_extended_value And ZEND_CTOR_CALL) Then
            'Class with op1 and op2
            myStop
            End If
            
            .push Key(op2)
            
         Case ZEND_DO_FCALL                  ' &h3C  60
            
           'Test num of args
            Debug.Assert FuncArgsCall.ESP = c.f_extended_value
            PushResult Key(op1) & "(" & _
               Join(FuncArgsCall.popArray(FuncArgsCall.ESP), ",") _
               & ")"
               
               
         Case ZEND_DO_FCALL_BY_NAME          ' &h3D  61
DoFunctionCall:
'            If Result.d_ObjectSize = 1 Then
'            DecompileLine.Concat .pop & "(" & _
'               Join(FuncArgsCall.popArray(FuncArgsCall.ESP), ",") _
'               & ")"
'            Else
               PushResult .pop & "(" & _
                  Join(FuncArgsCall.popArray(FuncArgsCall.ESP), ",") _
                  & ")"
'            End If
   
         Case ZEND_RETURN                    ' &h3E  62
            Dim tmp
            tmp = Var(op1)
            If IsNull(tmp) = False Then
               DecompileLine.Concat "Return (" & tmp & ")"
            End If
         
      '      https://www.codeblog.org/viewsrc/php-4.4.1/Zend/zend_execute.c
   '
   '       case ZEND_RETURN: {
   ' 1759:                                         zval *retval_ptr;
   ' 1760:                                         zval **retval_ptr_ptr;
   '1761:
   ' 1762:                                         if (EG(active_op_array)->return_reference == ZEND_RETURN_REF) {
   ' 1763:                                                 if (EX(opline)->op1.op_type == IS_CONST || EX(opline)->op1.op_type == IS_TMP_VAR) {
   ' 1764:                                                         /* Not supposed to happen, but we'll allow it */
   ' 1765:                                                         zend_error(E_NOTICE, "Only variable references should be returned by reference");
   ' 1766:                                                         goto return_by_value;
   ' 1767:                                                 }
   '1768:
   ' 1769:                                                 retval_ptr_ptr = get_zval_ptr_ptr(&EX(opline)->op1, EX(Ts), BP_VAR_W);
   '1770:
   ' 1771:                                                 if (!retval_ptr_ptr) {
   ' 1772:                                                         zend_error(E_ERROR, "Cannot return overloaded elements or string offsets by reference");
   ' 1773:                                                 }
   '1774:
   ' 1775:                                                 if (!(*retval_ptr_ptr)->is_ref) {
   ' 1776:                                                         if (EX(opline)->extended_value == ZEND_RETURNS_FUNCTION &&
   ' 1777:                                                                 EX(Ts)[EX(opline)->op1.u.var].var.fcall_returned_reference) {
   ' 1778:                                                                 /* intentionally left empty */
   ' 1779:                                                         } else if (EX(Ts)[EX(opline)->op1.u.var].var.ptr_ptr == &EX(Ts)[EX(opline)->op1.u.var].var.ptr) {
   ' 1780:                                                                 PZVAL_LOCK(*retval_ptr_ptr); /* undo the effect of get_zval_ptr_ptr() */
   ' 1781:                                                                 zend_error(E_NOTICE, "Only variable references should be returned by reference");
   ' 1782:                                                                 goto return_by_value;
   ' 1783:                                                         }
   ' 1784:                                                 }
   '1785:
   ' 1786:                                                 SEPARATE_ZVAL_TO_MAKE_IS_REF(retval_ptr_ptr);
         'https://www.codeblog.org/viewsrc/php-4.4.1/Zend/zend_execute.c
         
   
         Case ZEND_RECV                      ' &h3F  63
            
            FuncArgsCall.push Var(Result), Num(op1)

            
            
'            myStop
   
         Case ZEND_RECV_INIT                 ' &h40  64
            FuncArgsCall.push .pop & " = " & str(op2), Result.c_Object
'            myStop
   
         
         Case ZEND_SEND_VAL                  ' &h41  65
            
'            If c.f_extended_value = ZEND_DO_FCALL_BY_NAME Then
'
'              'ByRef
'              'A fixed value shouldn't be passed as Ref
'               myStop
'            ElseIf c.f_extended_value = ZEND_DO_FCALL Then
              
              'ByVal
               TmpFuncArgsIsRef.push False
               FuncArgsCall.push str(op1)
'            Else
'               myStop
'            End If
            DisamTxtComment = "Arg[" & Num(op2) & "] = " & FuncArgsCall.PreviewPop
   
         Case ZEND_SEND_VAR                  ' &h42  66
            If c.f_extended_value = ZEND_DO_FCALL_BY_NAME Then
               DisamTxtComment = "SHOULD_BE_SENT_BY_REF"
               GoTo send_by_ref
            
            ElseIf c.f_extended_value = ZEND_DO_FCALL Then
               
              'ByVal
               TmpFuncArgsIsRef.push False

            Else
              'Unusal branch
               myStop
            End If
            
            FuncArgsCall.push Var(op1)

   
         Case ZEND_SEND_REF                  ' &h43  67
send_by_ref:
            TmpFuncArgsIsRef.push True
            FuncArgsCall.push Var(op1)
'            myStop
   
         Case ZEND_NEW                       ' &h44  68
            myStop
   
         Case ZEND_JMP_NO_CTOR               ' &h45  69
            myStop
   
         Case ZEND_FREE                      ' &h46  70
            myStop
   
         Case ZEND_INIT_ARRAY                ' &h47  71
            Tmp_Prepare IS_ARRAY__4
            GoTo ZEND_ADD_ARRAY_ELEMENT

'            myStop
   
         Case ZEND_ADD_ARRAY_ELEMENT         ' &h48  72
ZEND_ADD_ARRAY_ELEMENT:

            Debug.Assert tmp_valueType = IS_ARRAY__4
            PushResult str(op1)
'            myStop
   
         Case ZEND_INCLUDE_OR_EVAL           ' &h49  73
            myStop
   
         Case ZEND_UNSET_VAR                 ' &h4A  74
            myStop
   
         Case ZEND_UNSET_DIM                 ' &h4B  75
            myStop
   
         Case ZEND_UNSET_OBJ                 ' &h4C  76
            myStop
   
         Case ZEND_FE_RESET                  ' &h4D  77
            myStop
   
         Case ZEND_FE_FETCH                  ' &h4E  78
            myStop
   
         Case ZEND_EXIT                      ' &h4F  79
            myStop
   
         Case ZEND_FETCH_R                   ' &h50  80
ZEND_FETCH:
'            If op1.a_op_type = IS_CONST__1 Then
'               .push MakeVar(op1), Result
               PushResult Var(op1)
'            Else
 '              myStop
 '           End If
   
         Case ZEND_FETCH_DIM_R               ' &h51  81
            myStop
   
         Case ZEND_FETCH_OBJ_R               ' &h52  82
            myStop
   
         Case ZEND_FETCH_W                   ' &h53  83
            GoTo ZEND_FETCH
   
         Case ZEND_FETCH_DIM_W               ' &h54  84
            myStop
   
         Case ZEND_FETCH_OBJ_W               ' &h55  85
            myStop
   
         Case ZEND_FETCH_RW                  ' &h56  86
            GoTo ZEND_FETCH
   
         Case ZEND_FETCH_DIM_RW              ' &h57  87
            myStop
   
         Case ZEND_FETCH_OBJ_RW              ' &h58  88
            myStop
   
         Case ZEND_FETCH_IS                  ' &h59  89
            myStop
   
         Case ZEND_FETCH_DIM_IS              ' &h5A  90
            myStop
   
         Case ZEND_FETCH_OBJ_IS              ' &h5B  91
            myStop
   
         Case ZEND_FETCH_FUNC_ARG            ' &h5C  92
            PushResult Var(op1)
'            myStop
   
         Case ZEND_FETCH_DIM_FUNC_ARG        ' &h5D  93
            myStop
   
         Case ZEND_FETCH_OBJ_FUNC_ARG        ' &h5E  94
            myStop
   
         Case ZEND_FETCH_UNSET               ' &h5F  95
            myStop
   
         Case ZEND_FETCH_DIM_UNSET           ' &h60  96
            myStop
   
         Case ZEND_FETCH_OBJ_UNSET           ' &h61  97
            myStop
   
         Case ZEND_FETCH_DIM_TMP_VAR         ' &h62  98
            myStop
   
         Case ZEND_FETCH_CONSTANT            ' &h63  99
            PushResult Key(op1)
   
         Case ZEND_DECLARE_FUNCTION_OR_CLASS ' &h64  100
            myStop
   
         Case ZEND_EXT_STMT                  ' &h65  101
            myStop
   
         Case ZEND_EXT_FCALL_BEGIN           ' &h66  102
            myStop
   
         Case ZEND_EXT_FCALL_END             ' &h67  103
            myStop
   
         Case ZEND_EXT_NOP                   ' &h68  104
            myStop
   
         Case ZEND_TICKS                     ' &h69  105
            myStop
   
         Case ZEND_SEND_VAR_NO_REF           ' &h6A  106
            myStop

   'added by me
         Case &HCC '   65'   e
            DisamTxtComment = "Common 'UNKNOWN OPCODE' "
            myStop


'PHP5 OpCodes
         Case &H6B To &HC7 '   65'   e
   
         Case ZEND_CATCH                     ' &h6B  107
            myStop
   
         Case ZEND_THROW                     ' &h6C  108
            myStop
   
         Case ZEND_FETCH_CLASS               ' &h6D  109
            myStop
   
         Case ZEND_CLONE                     ' &h6E  110
            myStop
   
         Case ZEND_INIT_CTOR_CALL            ' &h6F  111
            myStop
   
         Case ZEND_INIT_METHOD_CALL          ' &h70  112
            myStop
   
         Case ZEND_INIT_STATIC_METHOD_CALL   ' &h71  113
            myStop
   
         Case ZEND_ISSET_ISEMPTY_VAR         ' &h72  114
            myStop
   
         Case ZEND_ISSET_ISEMPTY_DIM_OBJ     ' &h73  115
            myStop
   
         Case ZEND_IMPORT_FUNCTION           ' &h74  116
            myStop
   
         Case ZEND_IMPORT_CLASS              ' &h75  117
            myStop
   
         Case ZEND_IMPORT_CONST              ' &h76  118
            myStop
   
         Case ZEND_OP_119                    ' &h77  119
            myStop
   
         Case ZEND_OP_120                    ' &h78  120
            myStop
   
         Case ZEND_ASSIGN_ADD_OBJ            ' &h79  121
            myStop
   
         Case ZEND_ASSIGN_SUB_OBJ            ' &h7A  122
            myStop
   
         Case ZEND_ASSIGN_MUL_OBJ            ' &h7B  123
            myStop
   
         Case ZEND_ASSIGN_DIV_OBJ            ' &h7C  124
            myStop
   
         Case ZEND_ASSIGN_MOD_OBJ            ' &h7D  125
            myStop
   
         Case ZEND_ASSIGN_SL_OBJ             ' &h7E  126
            myStop
   
         Case ZEND_ASSIGN_SR_OBJ             ' &h7F  127
            myStop
   
         Case ZEND_ASSIGN_CONCAT_OBJ         ' &h80  128
            myStop
   
         Case ZEND_ASSIGN_BW_OR_OBJ          ' &h81  129
            myStop
   
         Case ZEND_ASSIGN_BW_AND_OBJ         ' &h82  130
            myStop
   
         Case ZEND_ASSIGN_BW_XOR_OBJ         ' &h83  131
            myStop
   
         Case ZEND_PRE_INC_OBJ               ' &h84  132
            myStop
   
         Case ZEND_PRE_DEC_OBJ               ' &h85  133
            myStop
   
         Case ZEND_POST_INC_OBJ              ' &h86  134
            myStop
   
         Case ZEND_POST_DEC_OBJ              ' &h87  135
            myStop
   
         Case ZEND_ASSIGN_OBJ                ' &h88  136
            myStop
   
         Case ZEND_OP_DATA                   ' &h89  137
            myStop
   
         Case ZEND_INSTANCEOF                ' &h8A  138
            myStop
   
         Case ZEND_DECLARE_CLASS             ' &h8B  139
            myStop
   
         Case ZEND_DECLARE_INHERITED_CLASS   ' &h8C  140
            myStop
   
         Case ZEND_DECLARE_FUNCTION          ' &h8D  141
            myStop
   
         Case ZEND_RAISE_ABSTRACT_ERROR      ' &h8E  142
            myStop
   
         Case ZEND_START_NAMESPACE           ' &h8F  143
            myStop
   
         Case ZEND_ADD_INTERFACE             ' &h90  144
            myStop
   
         Case ZEND_VERIFY_INSTANCEOF         ' &h91  145
            myStop
   
         Case ZEND_VERIFY_ABSTRACT_CLASS     ' &h92  146
            myStop
   
         Case ZEND_ASSIGN_DIM                ' &h93  147
            myStop
   
         Case ZEND_ISSET_ISEMPTY_PROP_OBJ    ' &h94  148
            myStop
   
         Case ZEND_HANDLE_EXCEPTION          ' &h95  149
            myStop
   
         Case ZEND_USER_OPCODE               ' &h96  150
            myStop
   
         Case ZEND_U_NORMALIZE               ' &h97  151
            myStop
   
         Case ZEND_JMP_SET                   ' &h98  152
            myStop
   
         Case ZEND_DECLARE_LAMBDA_FUNCTION   ' &h99  153
            myStop
   
      End Select

interpret_Cmd_err:
'      log_verbose2 "                                                 ESP: " & .ESP
      
      If .ESP Then log_verbose2 "----------------------------------------------------------------"
      Dim StackItem
      For StackItem = 1 To .Storage.Count
         log_verbose2 "[" & StackItem & "] '" & .Storage(StackItem)
      Next
      If DisamTxtComment <> "" Then
         log_verbose2 "...................................................................."
         log_verbose2 "Comment: " & DisamTxtComment
      End If
      
   ' Create Function ArgList
   
   Dim isFunctionArgInitFinished As Boolean
   isFunctionArgInitFinished = False
   If isInFunctionArgInit Then
      Select Case cmd
         Case ZEND_FETCH_W, ZEND_RECV, ZEND_RECV_INIT
         Case Else
            isInFunctionArgInit = False
               
            DecompileLine.Concat "(" & _
                  Join(FuncArgsCall.popArray(FuncArgsCall.ESP), ",") _
                  & ")"
                  
              'Set Branchbody to end of function
               BranchesStack.push CmdCount
               isFunctionArgInitFinished = True
            
      End Select
   End If
      
      
'      #If isDebug Then
'         If DisamTxt = UNKNOWN_OPCODE Then
'            myStop
'            GoTo ReDoLoop
'         End If
'      #End If
   ' Handle Indepent \ 'Bodies
      If BranchesStack.ESP > 0 Then
         Dim CmdNr_EndBody&
         CmdNr_EndBody = BranchesStack.PreviewPop
         If CmdNr_EndBody = (CmdIndex + 1) Then
            BranchesStack.pop
            If .ESP > 0 Then DecompileLine.Concat .pop
         End If
         
      End If

'Stack Check
      With Result
      If (Stack.ESP <> 0) And ( _
           ((.a_op_type = IS_VAR__4) And (.d_ObjectSize = 1)) _
         ) Then
         log_verbose2 "Correcting Stack error..."
'         DecompileLine.Concat "\*Stack error detected and corrected*\"
         Do While Stack.ESP > 0
            DecompileLine.Concat Stack.pop
         Loop
      End If
      End With



      Dim IsNestUp As Boolean
      IsNestUp = (NestLevel < BranchesStack.ESP)
      
      Dim IsNestDown As Boolean
      IsNestDown = (NestLevel > BranchesStack.ESP)
            
      
      Dim IsCommandLineFinish As Boolean
      IsCommandLineFinish = (.ESP = 0) And _
                (FuncArgsCall.ESP = 0) And _
                (DecompileLine.value <> "")
      
      ' Or IsNestUp Or IsNestDown) Then '
      If IsCommandLineFinish Or isFunctionArgInitFinished Then
         
         Dim IsSkipReturn As Boolean
         IsSkipReturn = ((DecompileLine.value = "") And _
                         (CmdIndex + 1) = CmdCount)

         
         DecompileLine.Concat Switch( _
            IsNestUp, "{", _
            IsSkipReturn, "", _
            True, ";")
         
         log_verbose2 "===================================================================="
         log_verbose DecompileLine.value
         DecompiledOutPut DecompileLine.value, NestLevel
         
         If IsNestDown Then
            log_verbose "}"
            DecompiledOutPut "}", NestLevel - 1
         End If
         
       ' Reset LineBuffer
         DecompileLine.Clear
      
      End If
      
      If DecompileLine.value <> "" Then
         log_verbose2 "...................................................................."
         log_verbose2 "Buf: " & DecompileLine.value
      End If

      
      NestLevel = BranchesStack.ESP
      
      
'    ' Add command to history (<- this data is used to recognise a 'For' )
'      If CommandHistory.Count = 0 Then
'         CommandHistory.Add cmd
'      Else
'         CommandHistory.Add cmd, , 1
'      End If
      
   End With

End Sub










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
         
      Case IS_UNUSED__8
         Stop
         
      Case Else
         Stop
   End Select
End Sub


