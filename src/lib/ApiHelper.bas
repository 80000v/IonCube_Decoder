Attribute VB_Name = "ApiHelper"
Option Explicit
 
Private Const STATUS_PENDING As Long = &H103
Private Const STILL_ACTIVE As Long = STATUS_PENDING
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const WAIT_TIMEOUT As Long = 258&
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_DUP_HANDLE As Long = (&H40)
Private Const PROCESS_TERMINATE As Long = (&H1)
Private Const PROCESS_VM_OPERATION As Long = (&H8)
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long



Public Const SW_HIDE As Long = 0
Public Const SW_MAXIMIZE As Long = 3
Public Const SW_MINIMIZE As Long = 6
Public Const SW_NORMAL As Long = 1
Public Const SW_RESTORE As Long = 9
Public Const SW_SHOW As Long = 5
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByVal lpLastAccessTime As Long, ByRef lpLastWriteTime As FILETIME) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByVal lpLastAccessTime As Long, ByRef lpLastWriteTime As FILETIME) As Long

Public Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type


Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long




'The LB_GETHORIZONTALEXTENT message is useful to retrieve the current value of the horizontal extent:
Const LB_GETHORIZONTALEXTENT = &H193
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

' Set the horizontal extent of the control (in pixel).
' If this value is greater than the current control's width
' an horizontal scrollbar appears.

Sub Listbox_SetHorizontalExtent(LB As Listbox, ByVal newWidth As Long)
    SendMessage LB.hwnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&
End Sub


' Return the horizontal extent of the control (in pixel).
Function Listbox_GetHorizontalExtent(LB As Listbox) As Long
    Listbox_GetHorizontalExtent = SendMessage(LB.hwnd, LB_GETHORIZONTALEXTENT, 0, ByVal 0&)
End Function

Public Function FormatCTime(ByVal CTimeStamp&) As String
   FormatCTime = DateAdd("s", CTimeStamp, "01/01/1970")
End Function



Public Function FormatFileTime(TimeStamp As FILETIME) As String
   Dim SysTime As SYSTEMTIME
   With SysTime
      FileTimeToSystemTime TimeStamp, SysTime
      FormatFileTime = Format(.wDay & "." & .wMonth & "." & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond, "dd.mm.yyyy hh:mm:ss") & " [" & .wMilliseconds & "]"
   End With
End Function


Function ShellEx&(FileName$, Params$)
      
Dim Retval
'   On Error Resume Next
'  RetVal = ShellExecute(Me.hwnd, "open", """" & App.Path & "/" & "lzss.exe""", "-d """ & dbgFile.FileName & """ """ & outFileName & """", "", SW_NORMAL)
   Retval = Shell("""" & FileName & """ " & Params)
   
   If Retval Then
    Dim hProcess&, ExitCode&
    
    Retval = OpenProcess(PROCESS_QUERY_INFORMATION, 0, Retval)
    hProcess = Retval
    If hProcess Then
       Do
          Retval = GetExitCodeProcess(hProcess, ExitCode)
'               RetVal = WaitForSingleObject(hProcess, 100)
          DoEvents
       Loop While Retval And (ExitCode = STILL_ACTIVE)
       
       ShellEx = ExitCode
    End If

     
   End If

End Function


'Private Sub FileRename(SourceFileName$, destinationFileName$)
'         On Error Resume Next
'         log_verbose "Copying: " & SourceFileName & " -> " & destinationFileName
'
'         VBA.FileCopy SourceFileName, destinationFileName
'
'         If Err Then log_verbose "=> FAILED - " & Err.Description
'
'End Sub
Public Function FileRename(SourceFileName$, destinationFileName$) As Boolean

      Dim Retval&
'      log_verbose "Renaming: " & SourceFileName & " -> " & destinationFileName
      Retval = MoveFile(SourceFileName$ & vbNullChar, destinationFileName$ & vbNullChar)
      
      If Retval = 0 Then
         On Error Resume Next
         GetAttr SourceFileName
         If Err Then
            log_verbose "=> FAILED - Can't open source file!"
         Else
            GetAttr destinationFileName
            If Err = 0 Then
               log_verbose "=> FAILED - destination file already exists!"
            Else
               log_verbose "=> FAILED - source file is in use!"
            End If
         End If
      Else
         FileRename = True
      End If

End Function

Public Sub FileDelete(SourceFileName$)
   
   On Error Resume Next
   log_verbose "Deleting: " & SourceFileName
   
   Kill SourceFileName
   
   If Err Then log_verbose "=> FAILED - " & Err.description
  
End Sub

'Private Sub createBackup()
'   With FileName
'      On Error Resume Next
'      log_verbose " Creating Backup..."
'
'     'Prepare FileNames
'      Dim FileExe$, FileBak$
'      FileExe = .NameWithExt
'      FileBak = .Name & ".vEx"
'
'     'Set Workingdir
'      ChDrive .Path
'      ChDir .Path
'
' '    'Delete .bak
''      FileDelete FileBak
'
'      On Error GoTo 0
'
'     'Better we close the file before renaming...
'     'in short what later will cause problems:
'     'the openfilehandle which refered to winlogon.exe will
'     'refered to winlogon.bak after renaming, but the File.-objekt
'     'still thinks the openfilehandle belongs to winlogon.exe
'      File.CloseFile
'
'     'Rename .exe to .bak
'      FileRename FileExe, FileBak
'
'     'copy .bak to .exe
'      FileCopy FileBak, FileExe
'
'     'Remove readonly attrib & Test if .FileName exists => raise.Err 53
'      SetAttr .FileName, vbNormal
'
'   End With
'End Sub


Sub log_verbose2(Text$)
   FrmMain.log Text
End Sub


Sub log_verbose(Text$)
   FrmMain.log_verbose Text
End Sub

