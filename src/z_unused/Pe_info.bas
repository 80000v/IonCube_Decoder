Attribute VB_Name = "Pe_info_bas"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Type JOBOBJECT_BASIC_LIMIT_INFORMATION
  PerProcessUserTimeLimit As Long
  PerJobUserTimeLimit As Long
  LimitFlags As Long
  MinimumWorkingSetSize As Long
  MaximumWorkingSetSize As Long
  ActiveProcessLimit As Long
  Affinity As Long
  PriorityClass As Long
  SchedulingClass As Long
End Type


Public Type PE_Section
    SectionName          As String * 8
    VirtualSize          As Long
    RVAOffset            As Long
    RawDataSize          As Long
    PointertoRawData     As Long
    PointertoRelocs      As Long
    PointertoLineNumbers As Long
    NumberofRelocs       As Integer
    NumberofLineNumbers  As Integer
    SectionFlags         As Long
End Type

Public Type PE_Header
  PESignature                    As Long
  Machine                        As Integer
  NumberofSections               As Integer
  TimeDateStamp                  As Long
  PointertoSymbolTable           As Long
  NumberofSymbols                As Long
  OptionalHeaderSize             As Integer
  Characteristics                As Integer
  Magic                          As Integer
  MajorVersionNumber             As Byte
  MinorVersionNumber             As Byte
  SizeofCodeSection              As Long
  InitializedDataSize            As Long
  UninitializedDataSize          As Long
  EntryPointRVA                  As Long
  BaseofCode                     As Long
  BaseofData                     As Long

' extra NT stuff
  ImageBase                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  OSMajorVersion                 As Integer
  OSMinorVersion                 As Integer
  UserMajorVersion               As Integer
  UserMinorVersion               As Integer
  SubSysMajorVersion             As Integer
  SubSysMinorVersion             As Integer
  RESERVED                       As Long
  ImageSize                      As Long
  HeaderSize                     As Long
  FileChecksum                   As Long
  SubSystem                      As Integer
  DLLFlags                       As Integer
  StackReservedSize              As Long
  StackCommitSize                As Long
  HeapReserveSize                As Long
  HeapCommitSize                 As Long
  LoaderFlags                    As Long
  NumberofDataDirectories        As Long
'end of NTOPT Header
  ExportTableAddress             As Long
  ExportTableAddressSize         As Long
  ImportTableAddress             As Long
  ImportTableAddressSize         As Long
  ResourceTableAddress           As Long
  ResourceTableAddressSize       As Long
  ExceptionTableAddress          As Long
  ExceptionTableAddressSize      As Long
  SecurityTableAddress           As Long
  SecurityTableAddressSize       As Long
  BaseRelocationTableAddress     As Long
  BaseRelocationTableAddressSize As Long
  DebugDataAddress               As Long
  DebugDataAddressSize           As Long
  CopyrightDataAddress           As Long
  CopyrightDataAddressSize       As Long
  GlobalPtr                      As Long
  GlobalPtrSize                  As Long
  TLSTableAddress                As Long
  TLSTableAddressSize            As Long
  LoadConfigTableAddress         As Long
  LoadConfigTableAddressSize     As Long
  
  BoundImportsAddress            As Long
  BoundImportsAddressSize        As Long
  IATAddress                     As Long
  IATAddressSize                 As Long

  DelayImportAddress             As Long
  DelayImportAddressSize         As Long
  COMDescriptorAddress           As Long
  COMDescriptorAddressSize       As Long
  
  ReservedAddress                As Long
  ReservedAddressSize            As Long
  
'  Gap                            As String * &H28&
  Sections(64)                   As PE_Section
End Type


'assumption the .text Sections ist the first and .data Section the second in pe_header
Public Const TEXT_SECTION& = 0
Public Const DATA_SECTION& = 1

Public PE_info As New PE_info
Public PE_Header As PE_Header
Public File As New FileStream
Public file_readonly As New FileStream
Public FileName As New ClsFilename
Public PE_SectionData As Collection
