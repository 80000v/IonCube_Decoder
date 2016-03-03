Attribute VB_Name = "IC_Decompile_Helper"
Option Explicit

Const DeCompiledSource_CodeIndentation& = 3

Public Function MakeStr$(Text)
   MakeStr = StrQuote(Replace(Text, vbLf, "\n"))
End Function
Private Function GetObjData(c As IC_CmdAttrib, _
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


Public Function Var(c As IC_CmdAttrib)
   Var = GetObjData(c, False, True)
End Function

Public Function Key(c As IC_CmdAttrib)
   Key = GetObjData(c, True)
End Function

Public Function Str(c As IC_CmdAttrib)
   Str = GetObjData(c)
End Function

Public Function Num(c As IC_CmdAttrib)
   Num = GetObjData(c)
End Function


Public Sub DecompiledOutPut(TextLine$, Optional IndentLevel = 0)
   Dim CodeIndent$
   CodeIndent = String(IndentLevel * DeCompiledSource_CodeIndentation, " ")

   log_decompiled CodeIndent & TextLine
   
End Sub

Private Sub log_decompiled(Text$)
   FrmMain.List_Source.AddItem Text
End Sub
