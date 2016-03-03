VERSION 5.00
Begin VB.Form frmBase64 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB-Tec / Base64 - Jost Schwider, Soest"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTrimWS 
      Caption         =   "Trim WhiteSpaces"
      Height          =   315
      Left            =   3420
      TabIndex        =   7
      Top             =   60
      Width           =   1575
   End
   Begin VB.FileListBox lstFile 
      Height          =   2430
      Left            =   60
      TabIndex        =   2
      Top             =   3660
      Width           =   1995
   End
   Begin VB.DirListBox lstDir 
      Height          =   2790
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   1995
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   1995
   End
   Begin VB.CommandButton cmdFromBase64 
      Caption         =   "From Base64"
      Height          =   315
      Left            =   5040
      TabIndex        =   8
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdToBase64 
      Caption         =   "To Base64"
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txt 
      Height          =   5595
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "frmBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFromBase64_Click()
  MousePointer = vbHourglass
  txt.Text = Base64DecodeUnicode(txt.Text)
  MousePointer = vbDefault
End Sub

Private Sub cmdLoad_Click()
  Dim s As String
  
  s = ReadFile(lstFile.Path & "\" & lstFile.FileName)
  txt.Text = s
  DoEvents
  txt.Tag = s
End Sub

Private Sub cmdSave_Click()
  Dim FileName As String
  
  FileName = lstFile.Path & "\" & lstFile.FileName
  FileName = InputBox("Datei speichern unter...", , FileName)
  If Len(FileName) Then WriteFile FileName, txt.Text
End Sub

Private Sub cmdToBase64_Click()
  MousePointer = vbHourglass
  txt.Text = Base64EncodeUnicode(txt.Tag)
  MousePointer = vbDefault
End Sub

Private Sub cmdTrimWS_Click()
  txt.Tag = TrimWS2(txt.Tag)
  txt.Text = txt.Tag
End Sub

Private Sub lstDir_Change()
  lstFile.Path = lstDir.Path
End Sub

Private Sub lstDrive_Change()
  lstDir.Path = lstDrive.Drive
End Sub

Private Sub lstFile_DblClick()
  cmdLoad_Click
End Sub


Private Function TrimWS2(ByVal s As String) As String
  Dim i As Long
  Dim j As Long
  
  For i = 1 To Len(s)
    If Asc(Mid$(s, i, 1)) > 32 Then
      j = j + 1
      Mid$(s, j) = Mid$(s, i, 1)
    End If
  Next i
  TrimWS2 = Left$(s, j)
End Function

Private Function ReadFile(ByRef Path As String) As String
  Dim f As Integer
  
  f = FreeFile
  Open Path For Binary As #f
    ReadFile = Space$(LOF(f))
    Get #f, , ReadFile
  Close #f
End Function

Private Sub WriteFile(ByRef Path As String, ByRef Text As String)
  Dim FileNr As Long
  
  FileNr = FreeFile
    Open Path For Output As #FileNr
    Print #FileNr, Text;
  Close #FileNr
End Sub

Private Sub txt_Change()
  txt.Tag = txt.Text
End Sub
