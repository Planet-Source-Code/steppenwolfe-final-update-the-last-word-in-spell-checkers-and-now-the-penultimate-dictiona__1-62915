Attribute VB_Name = "mCompress"
Option Explicit

'** Hats off to Ion Alex Ionescu for the api compression routines..
'** Alex is easily one of the best coders on PSC

'//file constants//
Private Const GENERIC_WRITE                As Long = &H40000000
Private Const GENERIC_READ                 As Long = &H80000000
Private Const FILE_SHARE_WRITE             As Long = &H2
Private Const FILE_SHARE_READ              As Long = &H1
Private Const OPEN_ALWAYS                  As Long = 4
Private Const PAGE_READWRITE               As Long = &H4
Private Const SEC_COMMIT                   As Long = &H8000000
Private Const STANDARD_RIGHTS_REQUIRED     As Long = &HF0000
Private Const SECTION_EXTEND_SIZE          As Long = &H10
Private Const SECTION_MAP_EXECUTE          As Long = &H8
Private Const SECTION_MAP_READ             As Long = &H4
Private Const SECTION_MAP_WRITE            As Long = &H2
Private Const SECTION_QUERY                As Long = &H1
Private Const SECTION_ALL_ACCESS           As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

'//compression constants//
Private Const LZNT1 = &H2

'//Compression Engine Constants
Public Enum CompressionEngines
    COMPRESSION_ENGINE_STANDARD = &H0
    COMPRESSION_ENGINE_MAXIMUM = &H100
    COMPRESSION_ENGINE_HIBER = &H200
End Enum

#If False Then
Private COMPRESSION_ENGINE_STANDARD, COMPRESSION_ENGINE_MAXIMUM, COMPRESSION_ENGINE_HIBER
#End If

'//Buffer Manipulation Constants
Private Const MEM_COMMIT                 As Long = &H1000
Private Const MEM_DECOMMIT               As Long = &H4000
Private Const PAGE_EXECUTE_READWRITE     As Long = &H40

'//file api//
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, _
                                                                            ByVal dwDesiredAccess As Long, _
                                                                            ByVal dwShareMode As Long, _
                                                                            ByVal lpSecurityAttributes As Long, _
                                                                            ByVal dwCreationDisposition As Long, _
                                                                            ByVal dwFlagsAndAttributes As Long, _
                                                                            ByVal hTemplateFile As Long) As Long

Private Declare Function NtCreateSection Lib "ntdll.dll" (Handle As Long, _
                                                          ByVal DesiredAcess As Long, _
                                                          ObjectAttributes As Any, _
                                                          SectionSize As Any, _
                                                          ByVal Protect As Long, _
                                                          ByVal Attributes As Long, _
                                                          ByVal FileHandle As Long) As Long

Private Declare Function NtMapViewOfSection Lib "ntdll.dll" (ByVal Handle As Long, _
                                                             ByVal ProcessHandle As Long, _
                                                             BaseAddress As Long, _
                                                             ByVal ZeroBits As Long, _
                                                             ByVal CommitSize As Long, _
                                                             SectionOffset As Any, _
                                                             ViewSize As Long, _
                                                             ByVal InheritDisposition As Long, _
                                                             ByVal AllocaitonType As Long, _
                                                             ByVal Protect As Long) As Long

Public Declare Function NtUnmapViewOfSection Lib "ntdll.dll" (ByVal ProcessHandle As Long, _
                                                              ByVal Handle As Long) As Long
                                                              
Public Declare Function NtClose Lib "ntdll.dll" (ByVal hObject As Long) As Long

Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, _
                                                         lpFileSizeHigh As Long) As Long

Public Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long

Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, _
                                                           ByVal liDistanceToMove As Long, _
                                                           ByVal lpNewFilePointer As Long, _
                                                           ByVal dwMoveMethod As Long) As Long

'//compression api//
Private Declare Function NtFreeVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, _
                                                              BaseAddress As Long, _
                                                              regionsize As Long, _
                                                              ByVal FreeType As Long) As Long

Private Declare Function RtlCompressBuffer Lib "NTDLL" (ByVal CompressionFormatAndEngine As Integer, _
                                                        ByVal UnCompressedBuffer As Long, _
                                                        ByVal UnCompressedBufferSize As Long, _
                                                        ByVal CompressedBuffer As Long, _
                                                        ByVal CompressedBufferSize As Long, _
                                                        ByVal UncompressedChunkSize As Long, _
                                                        FinalCompressedSize As Long, _
                                                        ByVal Workspace As Long) As Long

Private Declare Function RtlDecompressBuffer Lib "NTDLL" (ByVal CompressionFormat As Integer, _
                                                          ByVal UnCompressedBufferPtr As Long, _
                                                          ByVal UnCompressedBufferSize As Long, _
                                                          ByVal CompressedBuffer As Long, _
                                                          ByVal CompressedBufferSize As Long, _
                                                          FinalCompressedSize As Long) As Long

Private Declare Function RtlGetCompressionWorkSpaceSize Lib "NTDLL" (ByVal CompressionFormatAndEngine As Integer, _
                                                                     CompressBufferWorkSpaceSize As Long, _
                                                                     CompressFragmentWorkSpaceSize As Long) As Long

Private Declare Function NtAllocateVirtualMemory Lib "ntdll.dll" (ByVal ProcessHandle As Long, _
                                                                  BaseAddress As Long, _
                                                                  ByVal ZeroBits As Long, _
                                                                  regionsize As Long, _
                                                                  ByVal AllocationType As Long, _
                                                                  ByVal Protect As Long) As Long


Public Enum eRatio
    cLow = 0
    cHigh = 1
End Enum

Private Workspace   As Long

Public Function Compress_File(ByVal sSrce As String, _
                              ByVal sDest As String, _
                              ByVal Ratio As eRatio) As Long
'//select compression options and build compressed file
Dim CompressionLevel  As CompressionEngines
Dim pFile             As Long
Dim FileSize          As Long
Dim FileHandle        As Long
Dim MemoryFileHandle  As Long
Dim FinalSize         As Long

On Error GoTo Handler

    Workspace = 0

    Select Case Ratio
        Case 0
            CompressionLevel = COMPRESSION_ENGINE_STANDARD
        Case 1
            CompressionLevel = COMPRESSION_ENGINE_MAXIMUM
    End Select

    Workspace = CreateWorkSpace(CompressionLevel)
    pFile = OpenFile(sSrce, FileSize, FileHandle, MemoryFileHandle)
    Compress CompressionLevel, pFile, FileSize, FinalSize, sDest, Workspace
    NtUnmapViewOfSection -1, pFile
    NtClose FileHandle
    NtClose MemoryFileHandle
    
    'dCRatio = Format(Int(FileSize / FinalSize), "##.##")
    Compress_File = 1
    Exit Function
    
Handler:
Compress_File = 0

End Function

Public Function Decompress_File(ByVal sSrce As String, _
                                ByVal sDest As String) As Long
'//decompress file
Dim pFile             As Long
Dim FileSize          As Long
Dim FileHandle        As Long
Dim MemoryFileHandle  As Long
Dim FinalSize         As Long
Dim lStart            As Long

On Error GoTo Handler

    pFile = OpenFile(sSrce, FileSize, FileHandle, MemoryFileHandle)
    lStart = GetFileSize(FileHandle, 0)
    DeCompress pFile, FileSize, FinalSize, sDest
    NtUnmapViewOfSection -1, pFile
    NtClose FileHandle
    NtClose MemoryFileHandle

    'dCRatio = Format(Int(FinalSize / lStart), "##.##")
    Decompress_File = 1
    Exit Function
    
Handler:
Decompress_File = 0

End Function

'//end of exposed functions//

Private Function OpenFile(File As String, _
                          SIZE As Long, _
                          FileHandle As Long, _
                          MemoryHandle As Long) As Long

Dim BaseAddress As Long

    FileHandle = CreateFile(File, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_ALWAYS, 0&, 0&)
    '//Open the file from disk
    '//Get size of the file if not there already
    If SIZE = 0 Then
        SIZE = GetFileSize(FileHandle, 0)
    End If
    NtCreateSection MemoryHandle, SECTION_ALL_ACCESS, ByVal 0&, SIZE, PAGE_READWRITE, SEC_COMMIT, FileHandle
    '//Load file to memory
    NtMapViewOfSection MemoryHandle, -1, BaseAddress, 0&, SIZE, 0&, SIZE, 1, 0, PAGE_READWRITE
    '//Map it into memory
    OpenFile = BaseAddress

End Function

Private Function CreateWorkSpace(Engine As CompressionEngines) As Long

'//IN: Format and Engine
'//OUT: Pointer to Buffer
'//About CompressionFormatAnd Engine: This is an integer, which means 4 bytes, or in hex: 0xYYYY where Y can be 0 to F. The first two bytes are the format
'//and the last two bytes are the engine. Therefore, LZNT1 which is 0x0002 and high compression which is 0x0100 would become 0x0102, basically we OR the two

Dim CompressionType       As Integer    '//Holds our two ORed values
Dim WorkSpaceSize         As Long       '//Return value from API call
Dim FragmentWorkSpaceSize As Long       '//We don't care about this one

    '//Create the Workspace
    CompressionType = LZNT1 Or Engine
    '//Calculate the Format+Engine Value
    RtlGetCompressionWorkSpaceSize CompressionType, WorkSpaceSize, FragmentWorkSpaceSize
    '//Call the API to get our Workspace Size
    NtAllocateVirtualMemory -1, CreateWorkSpace, 0, WorkSpaceSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE
    '//Return a pointer to the WorkSpace Buffer

End Function

Private Function Compress(Engine As CompressionEngines, _
                          UnCompressedBuffer As Long, _
                          ByVal UnCompressedBufferSize As Long, _
                          FinalSize As Long, _
                          NewFile As String, _
                          ByVal Workspace As Long) As Long

'//IN: Compression Format and Engine, Data to compress as a Byte Array, Pointer to Workspace Buffer,
'//and ChunkSize (&H1000 is recommended, but can be 0)
'//OUT: Compressed Data, Final Size of the Compressed Data

Dim CompressionType         As Integer      '//Holds our two ORed values
Dim CompressedBuffer        As Long         '//Holds the buffer to receive
Dim CompressedBufferSize    As Long         '//Holds the size of the buffer to receive
Dim CompressedBufferHandle  As Long
Dim CompressedBufferHandle2 As Long

    '//Open Destination File
    CompressedBufferSize = UnCompressedBufferSize * 1.13 + 4
    '//Size of compressed buffer can never be bigger
    CompressedBuffer = OpenFile(NewFile, CompressedBufferSize, CompressedBufferHandle, CompressedBufferHandle2)
    '// Allocate it
    CompressionType = LZNT1 Or Engine
    '//Calculate the Format+Engine Value
    Compress = RtlCompressBuffer(CompressionType, UnCompressedBuffer, UnCompressedBufferSize, CompressedBuffer, CompressedBufferSize, 0&, FinalSize, Workspace)
    '//Write the new file
    NtUnmapViewOfSection -1, CompressedBuffer
    NtClose CompressedBufferHandle2
    SetFilePointer CompressedBufferHandle, FinalSize, 0, 0
    SetEndOfFile CompressedBufferHandle
    NtClose CompressedBufferHandle

    '//Empty the Workspace Buffer
    NtFreeVirtualMemory -1, Workspace, 0, MEM_DECOMMIT

End Function

Private Function DeCompress(CompressedBuffer As Long, _
                           CompressedBufferSize As Long, _
                           FinalSize As Long, _
                           NewFile As String) As Long

'//IN: Compression Format, Data to decompress as a Byte Array
'//OUT: Decompressed Data, Final Size of the Compressed Buffer
'//Variable Declarations

Dim UnCompressedBuffer     As Long   '//Holds the buffer to send
Dim UnCompressedBufferSize As Long   '//Holds the size of the buffer to send
Dim OriginalBufferHandle   As Long
Dim OriginalBufferHandle2  As Long

    '//Calculations needed for the API Call
    UnCompressedBufferSize = CompressedBufferSize * 12.5
    '//Max compression possible (92%)
    UnCompressedBuffer = OpenFile(NewFile, UnCompressedBufferSize, OriginalBufferHandle, OriginalBufferHandle2)
    '//Pointer to the compressed buffer
    DeCompress = RtlDecompressBuffer(LZNT1, UnCompressedBuffer, UnCompressedBufferSize, CompressedBuffer, CompressedBufferSize, FinalSize)
    '//Write the new file
    NtUnmapViewOfSection -1, UnCompressedBuffer
    NtClose OriginalBufferHandle2
    SetFilePointer OriginalBufferHandle, FinalSize, 0, 0
    SetEndOfFile OriginalBufferHandle
    NtClose OriginalBufferHandle

End Function

