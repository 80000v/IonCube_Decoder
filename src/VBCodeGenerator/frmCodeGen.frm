VERSION 5.00
Begin VB.Form frmCodeGen 
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   Icon            =   "frmCodeGen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Txt_data 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmCodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const IndentCount& = 3
Private Sub Form_Load()
   
   Dim Operators
   Operators = Array("+", "-", "*", "/", "%", "<<", ">>", ".", "|", "%", "%", "!", "!!", "^", "===", "!==", "==", "!=", "<", "<=", "?CAST?", ":", "+=", "-=", "*=", "\=", "%=", "<<=", ">>=", ".=", "|=", "%=", "^=", "++", "--", "++")
   
   
   Dim LineMaxLength&
   LineMaxLength = 0
   
   
   Dim NameMaxLength&
   NameMaxLength = 0
   
   Dim output
   ArrayAdd output, "Attribute VB_Name = ""IC_DecompileConsts"""
   ArrayAdd output, "Option Explicit"
   
   Dim ConstStartLine&
   ConstStartLine = UBound(output) + 1
   
   Dim lineNr&
   For lineNr = LBound(icByteCodeNames) To UBound(icByteCodeNames)
      Dim Text$
      Text = icByteCodeNames(lineNr)
      
      
      Dim Operator
      Operator = ""
      On Error Resume Next
      Operator = Operators(lineNr - 1)
      
      Dim outputLine$
      outputLine = Join(Array( _
         "Public Const ", Text, " as Long = &h", H8(lineNr) _
         ), "")
         
      MaxProc LineMaxLength, Len(outputLine)
      MaxProc NameMaxLength, Len(Text)
      
      ArrayAdd output, outputLine
   Next
   
   Inc LineMaxLength, 2
   Inc NameMaxLength, 2
   
   For lineNr = ConstStartLine To UBound(output)
      output(lineNr) = BlockAlign_r(output(lineNr), LineMaxLength) & "  ' " & lineNr - ConstStartLine & "  " & Operator
   Next
   Txt_data = Join(output, vbCrLf)
   FileSave App.Path & "\IC_Decompile_Const.bas", Txt_data
   
   ReDim output(0)
   ArrayAdd output, "Attribute VB_Name = ""IC_DecompileFuncs"""
   ArrayAdd output, "Option Explicit"
   
   
'-------Select case -------------

   ArrayAdd output, ""
   ArrayAdd output, ""
   
   ArrayAdd output, Indent(0) & "sub ic_interpret(cmd&)"
   ArrayAdd output, Indent(1) & ""
   ArrayAdd output, Indent(1) & "Dim Prefix$, InFix$, PostFix$"
   ArrayAdd output, Indent(1) & ""
   ArrayAdd output, Indent(1) & "select case cmd"
   
   ArrayAdd output, ""

   Dim Functions
   ArrayAdd Functions, Indent(0) & "'"
   ArrayAdd Functions, Indent(0) & "'"
   ArrayAdd Functions, Indent(0) & "'"
   ArrayAdd Functions, Indent(0) & "'" & String(70, "=")
   ArrayAdd Functions, Indent(0) & ""
   ArrayAdd Functions, Indent(0) & "'        F U N C T I O N S     " & String(70, "=")
   ArrayAdd Functions, Indent(0) & ""
   ArrayAdd Functions, Indent(0) & "'" & String(70, "=")
   
   For lineNr = LBound(icByteCodeNames) To UBound(icByteCodeNames)
      Text = icByteCodeNames(lineNr)
      
      Operator = ""
      On Error Resume Next
      Operator = Operators(lineNr - 1)
      
      outputLine = Join(Array( _
      Indent(2) & "case ", BlockAlign_r(Text, NameMaxLength), _
      " ' &h", H8(lineNr), "  ", lineNr _
      ), "")
      
      ArrayAdd output, outputLine
      
      Dim FuncName$
      FuncName = "Do" & Text
      
      ArrayAdd Functions, Indent(0) & "'" & String(70, "/")
      ArrayAdd Functions, Indent(0) & "'//  " & FuncName
      ArrayAdd Functions, Indent(0) & "'//  "
      ArrayAdd Functions, Indent(0) & "Private Sub " & FuncName & "()"
      ArrayAdd Functions, Indent(1) & "On error goto " & FuncName & "_err"
      ArrayAdd Functions, Indent(0) & ""
      ArrayAdd Functions, Indent(2) & "PreFix= """" : InFix = """": PostFix = """""
      ArrayAdd Functions, Indent(2) & ""
      ArrayAdd Functions, Indent(2) & ""
      ArrayAdd Functions, Indent(2) & ""
      ArrayAdd Functions, Indent(0) & Text & "_err:"
      ArrayAdd Functions, Indent(0) & "Select case err"
      ArrayAdd Functions, Indent(1) & "Case 0"
      ArrayAdd Functions, Indent(2) & ""
      ArrayAdd Functions, Indent(1) & "Case else"
      ArrayAdd Functions, Indent(2) & "Err.raise vbobjecterror , , ""Error in Decomile!" & FuncName & """ - """""
      ArrayAdd Functions, Indent(1) & "end select"
      ArrayAdd Functions, Indent(0) & "End Sub"
      ArrayAdd Functions, Indent(0) & ""
      ArrayAdd Functions, Indent(0) & ""
'      ArrayAdd Functions, Indent(0) & "'" & String(70, "_")
      ArrayAdd Functions, Indent(0) & ""
      ArrayAdd Functions, Indent(0) & ""
      
      
'      ArrayAdd output, Indent(3) & FuncName
      ArrayAdd output, Indent(3) & "PreFix= """" : InFix = """ & Operator & """: PostFix = """""
      ArrayAdd output, Indent(3) & "myStop"
      ArrayAdd output, ""
   Next
      
   
   ArrayAdd output, Indent(1) & "end select"
   ArrayAdd output, Indent(0) & "end sub"

   Txt_data = Join(output, vbCrLf) ' & _
              Join(Functions, vbCrLf)

   FileSave App.Path & "\IC_Decompile.bas", Txt_data
   MsgBox "Saved to :" & App.Path
End Sub

Function Indent$(Level)
   Indent = String(Level * IndentCount, " ")
End Function

