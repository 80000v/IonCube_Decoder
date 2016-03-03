VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PE_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private PEStart&
'Private Const CHECKSUM_OPEN_FAILURE As Long = 1
'Private Const CHECKSUM_MAP_FAILURE As Long = 2
'Private Const CHECKSUM_MAPVIEW_FAILURE As Long = 3
'Private Const CHECKSUM_UNICODE_FAILURE As Long = 4
Private Declare Function MapFileAndCheckSum Lib "E:\Programmierung\Projekte\vir_clean\Imagehlp.dll" Alias "MapFileAndCheckSumA" (ByVal FileName As String, ByRef HeaderSum As Long, ByRef CheckSum As Long) As Long

Public Sub Create()
'     '--- find PE-signature ---
'     'Get First 0x400 Bytes
'      Dim tmpstr$
'      file.Position = 0
'      tmpstr = file.FixedString(&H400)
'
'     'Locate start of PE-header
'      PEStart = InStr(1, tmpstr, "PE" & vbNullChar & vbNullChar, vbBinaryCompare)
'      If PEStart = 0 Then err.Raise vbObjectError + 1, , "No PE-Header Found"
    
     '--- find PE-signature ---
     'Check DOS Header
      Dim tmpstr$
      file.Position = 0
     
     'MZ DOS-Header->e_magic
      If file.intValue <> &H5A4D Then Err.Raise vbObjectError + 1, , "No ExeFile DOS-Header.e_magic<>""MZ"""

     
     'Locate & Validate PE-header
      file.Position = &H3C '   DOS-Header->e_lfanew
      PEStart = file.longValue
      file.Position = PEStart
      PEStart = PEStart + 1
      
      If file.longValue <> &H4550 Then Err.Raise vbObjectError + 2, , "No ExeFile 'PE-Header.Signature<>""PE"""
    
    '  --- get PE_Header  ---
      Dim hFile&
      hFile = FreeFile
      Open file.FileName For Binary Access Read As #hFile
        Get hFile, PEStart, PE_Header
      Close hFile
      
      
    ' Validate Machine Type
      If PE_Header.Machine <> &H14C Then
         If PE_Header.Machine = &H8664 Then
            Err.Raise vbObjectError + 4, , "PE-Header.Signature=HDR64_MAGIC!"
         Else
           Err.Raise vbObjectError + 3, , "Unsupported PE-Header.Signature<>I386(0x14C)."
         End If
      End If
     
'    ' Validate Optionaldata Type
'      If PE_Header.Magic <> &H10B Then
'         err.Raise vbObjectError + 5, , "PE_Header.Magic <> PE32!"
'      End If
     
      If PE_Header.OptionalHeaderSize <> &HE0 Then
         Err.Raise vbObjectError + 5, , "PE_Header.OptionalHeaderSize = E0 expected"
      End If
      

     
'      If PE_Header.NumberofDataDirectories > &HE Then
'         err.Raise vbObjectError + 5, , "PE_Header.NumberofDataDirectories must be greater than 14"
'      End If
'imagever 3..5 <-winXP / 6 vista
     
End Sub

Public Sub WriteHeader()
   If PEStart = 0 Then Err.Raise vbObjectError, , "You must call PE_info::Create first!"
    
    Dim tmpstr$
    tmpstr = Space(Len(PE_Header) - (UBound(PE_Header.Sections) - PE_Header.NumberofSections) * &H28)
'    Stop
    MemCopyAnyToStr tmpstr, PE_Header, Len(tmpstr)
    file.Position = PEStart - 1
    file.FixedString(-1) = tmpstr
        
'    '  --- get PE_Header  ---
'      Dim hFile&
'      hFile = FreeFile
'      Open file.FileName For Binary Access Write As #hFile
'        Put hFile, PEStart, PE_Header
'      Close hFile

End Sub

Public Function UpdateChecksum&() '(Optional ByRef Output_ChecksumCurrent&, Optional ByRef Output_ChecksumNew&)
   If PEStart = 0 Then Err.Raise vbObjectError, , "You must call PE_info::Create first!"
   UpdateChecksum = MapFileAndCheckSum(file.FileName, 0, PE_Header.FileChecksum)
   WriteHeader
End Function


Public Function GetChecksum&() '(Optional ByRef Output_ChecksumCurrent&, Optional ByRef Output_ChecksumNew&)
   If PEStart = 0 Then Err.Raise vbObjectError, , "You must call PE_info::Create first!"
   Dim tmp&, RetVal&
   RetVal = MapFileAndCheckSum(file.FileName, 0, tmp)
   GetChecksum = tmp
End Function

Public Function ComputeSizeOfHeader&()
   ComputeSizeOfHeader = PEStart + &H18 + PE_Header.OptionalHeaderSize + PE_Header.NumberofSections * &H28
End Function


Public Sub PE_GetSectionData()
   'ReDim PE_SectionData(PE_Header.NumberofSections - 1)
   
   Set PE_SectionData = New Collection
   
   Dim i&, filedata As StringReader
      For i = 0 To PE_Header.NumberofSections - 1
         With PE_Header.Sections(i)
            
            Set filedata = New StringReader
            
            file.Position = .PointertoRawData
            filedata = file.FixedString(.RawDataSize)
            PE_SectionData.Add filedata
         End With
      Next

End Sub

Public Sub PE_SetSectionData()
   Dim i&
      For i = 0 To PE_Header.NumberofSections - 1
         With PE_Header.Sections(i)
            file.Position = .PointertoRawData
            file.FixedString(.RawDataSize) = PE_SectionData(i)
         End With
      Next
'   ReDim PE_SectionData(0)
End Sub

