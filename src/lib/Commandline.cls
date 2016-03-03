VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Commandline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CommandLine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetCommandLineRaw Lib "kernel32" Alias "GetCommandLineA" () As String
Private argsWithQuotes$()
Private Args$()
Private EmptyArray$()

Public NumberOfCommandLineArgs
Public commandline As String


'///////////////////////////////////////////////////
'// getArgs - Get CommandLineArguments with Quotes
Public Function getArgs() As String()
   getArgs = Args
End Function

Public Function getArgsWithQuotes()
   getArgsWithQuotes = argsWithQuotes
End Function


'//////////////////////////////////////////////////////////////////////////
'// Constructor - Sets args, args and NumberOfCommandLineArgs
Private Sub Class_Initialize()
   
  GetCommandLine

  If commandline = Empty Then Exit Sub
  Call commandLineSplitRek(commandline)

End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
'// commandLineSplitRek -  Split the commandLine and store chunks in args() and argsWithQuotes()
'//                        should only used by the Constructor
Private Sub commandLineSplitRek(lCommandLine As String, Optional oldpos% = 1, Optional levelCounter = 0)
'On Error GoTo commandLineSplitRek_err
   
   Dim endPos%                ' Endposition for cut (...e.exe"<-)
   Dim startPos%           ' Startposition for "new" commandline (->"C:\p...")
   Dim bIsQuoted As Boolean
   Dim bEndLine As Boolean
   Dim bNoQuotedAtEnd As Boolean
   
  'Filter Leading Space
   For startPos = 1 To Len(lCommandLine)
     If Mid(lCommandLine, startPos, 1) > " " Then Exit For
   Next
   
  'Trivial case - we are already at the end
   If (startPos >= Len(lCommandLine)) Then
           
    ' create array for saving commandline arguments
      NumberOfCommandLineArgs = levelCounter
      If NumberOfCommandLineArgs > 0 Then
         ReDim argsWithQuotes(NumberOfCommandLineArgs - 1)
         ReDim Args(NumberOfCommandLineArgs - 1)
      Else
       'Clear Arrays
         argsWithQuotes = EmptyArray
         Args = EmptyArray
      End If



  'Recusion case - Look for next element
   Else
   
   ' does Commandline starts with "C:\P...
      bIsQuoted = Mid(lCommandLine, startPos, 1) = """"
      If bIsQuoted Then
       ' find next "(DoubleQuote)
         endPos = InStr(startPos + 1, lCommandLine, """") + 1
         bEndLine = (endPos <= 1)
      Else
       ' find next ' '(Space)
         endPos = InStr(startPos, lCommandLine, " ")
         bEndLine = (endPos <= 0)
      End If
      
      If bEndLine Then
        endPos = Len(lCommandLine)
      End If
      
      ' Call commandLineSplitRek recursiv with "new" commandline 6 increase levelCounter
        commandLineSplitRek Mid(lCommandLine, endPos), startPos, levelCounter + 1
      

'      bNoQuotedAtEnd = bIsQuoted And bEndLine And (endPos = 0)
   
     Dim length&
     length = endPos - startPos
     
    'Save Data in Array
     argsWithQuotes(levelCounter) = Mid(lCommandLine, startPos, length + 1)
     If bIsQuoted Then
        Args(levelCounter) = Mid(argsWithQuotes(levelCounter), 2, length - 2)
        
     Else
        Args(levelCounter) = argsWithQuotes(levelCounter)
        
     End If
     
   End If
   
'commandLineSplitRek_err:
End Sub
Private Function GetCommandLine() As String

  commandline = Command

  
End Function

