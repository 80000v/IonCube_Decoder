Attribute VB_Name = "RegExp"
Option Explicit

Dim myRegExp As New RegExp

Public Const RE_Anchor_LineBegin$ = "^"
Public Const RE_Anchor_LineEnd$ = "$"

Public Const RE_Anchor_WordBoarder$ = "\b"
Public Const RE_Anchor_NoWordBoarder$ = "\B"

Public Const RE_AnyChar$ = "."
Public Const RE_AnyChars$ = ".*"

Public Const RE_AnyCharNL$ = "[\S\s]"
Public Const RE_AnyCharsNL$ = "[\S\s]*?"

Public Const RE_NewLine$ = "\r?\n"


Public Function RE_WSpace(ParamArray Elements()) As String
   Dim WS$ ' WhiteSpace
   WS = "\s*"
   
   RE_WSpace = Join(Elements, WS)
End Function



Public Function RE_LookHead_positive(ExpressionThatShouldBeFound$) As String
   RE_LookHead_positive = "(?=" & ExpressionThatShouldBeFound & ")"
End Function

Public Function RE_LookHead_negative(ExpressionThatShouldNOTBeFound$) As String
   RE_LookHead_negative = "(?!" & ExpressionThatShouldNOTBeFound & ")"
End Function

Public Function RE_Repeat(Optional MinRepeat& = 0, Optional MaxRepeat = "") As String
   If (MinRepeat = MaxRepeat) Then
      RE_Repeat = "{" & MinRepeat & "}"
   Else
      RE_Repeat = "{" & MinRepeat & "," & MaxRepeat & "}"
   End If
   
End Function


Public Function RE_AnyCharRepeat(Optional MinRepeat& = 0, Optional MaxRepeat = "") As String
   RE_AnyCharRepeat = "." & RE_Repeat(MinRepeat, MaxRepeat)
End Function

Public Function RE_Group(RegExpForTheGroup$) As String
   RE_Group = "(" & RegExpForTheGroup & ")"
End Function

Public Function RE_Group_NonCaptured(RegExpForTheNonCapturedGroup$) As String
   RE_Group_NonCaptured = "(?:" & RegExpForTheNonCapturedGroup & ")"
End Function

Public Function RE_Literal(TextWithLiterals) As String
   'Mask metachars
   RE_Literal = RE_Mask(TextWithLiterals, "][{}()*+?.\\^$|")
                                           
End Function


Public Function RE_Replace_Literal(TextWithLiterals) As String
  'Mask Replace metachars
   ' $0-9   Back reference
   ' $+     Last reference
   
   ' $&     MatchText
   
   ' $`     Text left from subject
   ' $'     Text right from subject
   ' $_     Whole subject
   
   RE_Replace_Literal = RE_Mask(TextWithLiterals, "0-9+`'_", "\$", "$$")


End Function
Private Sub RE_Mask_Whitespace(Text)
   ReplaceDo Text, vbCr, "\r"
   ReplaceDo Text, vbLf, "\n"
   ReplaceDo Text, vbTab, "\t"
End Sub

Private Function RE_Mask(Text, CharsToMask$, _
   Optional CharMaskSearch$ = "", _
   Optional CharMaskReplace$ = "\") As String
   With myRegExp
      .Global = True
      
     ' Mask MetaChars like with a preciding '\'
      .Pattern = CharMaskSearch & "[" & CharsToMask & "]"
      
     'Attention Text is passed byref - so don use Text =...!
      RE_Mask = .Replace(Text, CharMaskReplace & "$&")
   
   
   End With

'   RE_Mask_Whitespace Text
   
'   RE_Mask = Text

End Function

Public Function RE_CharCls(Chars$) As String
   ' mask ']' and '-'
   RE_CharCls = "[" & RE_Mask(Chars, "]\\-") & "]"
End Function

Public Function RE_CharCls_Excluded(Chars$) As String
   ' mask ']' and '-'
   RE_CharCls_Excluded = "[^" & RE_Mask(Chars, "]\\-") & "]"

End Function

