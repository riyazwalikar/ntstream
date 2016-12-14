Attribute VB_Name = "Functions"

Public FileName As String
Public isObjFile As Long
Public ListObj As ListItem
Public StreamBuffer As String
Public quitjob As Long
Public formstat As Long

Public ScanFiles As Long
Public ScanFolders As Long
Public FileWithADS As Long
Public DirWithADS As Long
Public BigFSO As New FileSystemObject

Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40

Public Const BIF_RETURNONLYFSDIRS As Long = &H1
Public Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Public Const BIF_RETURNFSANCESTORS As Long = &H8
Public Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Public Const BIF_BROWSEFORPRINTER As Long = &H2000
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Public Const MAX_PATH As Long = 260

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function BackupRead Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Long) As Long
Public Declare Function BackupSeek Lib "kernel32" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, ByRef lpdwLowByteSeeked As Long, ByRef lpdwHighByteSeeked As Long, ByRef lpContext As Long) As Long
Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function QueryFileInfo Lib "NTDLL.DLL" Alias "NtQueryInformationFile" (ByVal FileHandle As Long, IoStatusBlock_Out As IO_STATUS_BLOCK, lpFileInformation_Out As Long, ByVal Length As Long, ByVal FileInformationClass As FILE_INFORMATION_CLASS) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Const DefBufferSize As Long = 128& * 1024&

Private Type BrowseInfo
  hWndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Public Enum FileAttributes
    FILE_ATTRIBUTE_DIRECTORY = &H10
End Enum

Public Enum FileAccess
    FRead = &H80000000
    FWrite = &H40000000
    FReadWrite = &H80000000 Or &H40000000
    FDelete = &H10000
    FReadControl = &H20000
    FWriteDac = &H40000
    FWriteOwner = &H80000
    FSynchronize = &H100000
    FStandardRightsRequired = &HF0000
    FStandardRightsAll = &H1F0000
    FSystemSecurity = &H1000000
End Enum

Public Enum FileShare
    FNone = 0
    FRead
    FWrite
    FReadWrite
    FDelete
End Enum

Public Enum FileMode
    CreateNew = 1
    CreateAlways
    OpenExisting
    OpenOrCreate
    Truncate
    Append
End Enum

Public Enum OpenFile
    NoFlags = 0
    POSIXSemantics = &H1000000
    BackupSemantics = &H2000000 'will be used by CreateFile to Open file Or a folder..
    DeleteOnClose = &H4000000
    SequentialScan = &H8000000
    RandomAccess = &H10000000
    NoBuffering = &H20000000
    OverlappedIO = &H40000000
    WriteThrough = &H80000000
End Enum

Public Enum FilePointerOptions   ' Options for the SetFilePointer API
    BeginOfFile = 0
    FileCurrentPosition = 1
    EndOfFile = 2
End Enum

Public Type WIN32_STREAM_ID
    dwStreamID As Long
    dwStreamAttributes As Long
    Size As Long
    SizeHi As Long
    dwStreamNameSize As Long
End Type

Public Enum FILE_INFORMATION_CLASS
    FileDirectoryInformation = 1
    FileFullDirectoryInformation = 2
    FileBothDirectoryInformation = 3
    FileBasicInformation = 4
    FileStandardInformation = 5
    FileInternalInformation = 6
    FileEaInformation = 74
    FileAccessInformation = 8
    FileNameInformation = 9
    FileRenameInformation = 10
    FileLinkInformation = 11
    FileNamesInformation = 12
    FileDispositionInformation = 13
    FilePositionInformation = 14
    FileFullEaInformation = 15
    FileModeInformation = 16
    FileAlignmentInformation = 17
    FileAllInformation = 18
    FileAllocationInformation = 19
    FileEndOfFileInformation = 20
    FileAlternateNameInformation = 21
    FileStreamInformation = 22
    FilePipeInformation = 23
    FilePipeLocalInformation = 24
    FilePipeRemoteInformation = 25
    FileMailslotQueryInformation = 26
    FileMailslotSetInformation = 27
    FileCompressionInformation = 28
    FileObjectIdInformation = 29
    FileCompletionInformation = 30
    FileMoveClusterInformation = 31
    FileQuotaInformation = 32
    FileReparsePointInformation = 33
    FileNetworkOpenInformation = 34
    FileAttributeTagInformation = 35
    FileTrackingInformation = 36
    FileMaximumInformation = 37
End Enum

Public Type IO_STATUS_BLOCK
    IoStatus As Long
    Info As FILE_INFORMATION_CLASS
End Type
    
Public Type FILE_STREAM_INFO
    NextEntryOffset As Long
    StreamNameLength As Long
    StreamSize As Long
    StreamSizeHi As Long
    StreamAllocationSize As Long
    StreamAllocationSizeHi As Long
    streamname(259) As Byte
End Type

Public Enum FileStreamTypes
  BACKUP_INVALID = 0
  BACKUP_DATA = 1                     ' Standard data stream (NTFS names "::DATA$")
  BACKUP_EA_DATA = 2                  ' Extended attribute data
  BACKUP_SECURITY_DATA = 3            ' Contains ACL's, etc.
  BACKUP_ALTERNATE_DATA = 4           ' Alternative data stream
  BACKUP_LINK = 5                     ' Posix style hard link
  BACKUP_PROPRETY_DATA = 6            ' Property data
  BACKUP_OBJECT_ID = 7                ' Uniquely identifies a file in the file system
  BACKUP_REPARSE_DATA = 8             ' Stream uses reparse points
  BACKUP_SPARSE_BLOCK = 9             ' Stream is a sparse file.
End Enum

Public Enum FileStreamAttributes
  BACKUP_NORMAL_ATTRIBUTE = &H0
  BACKUP_MODIFIED_WHEN_READ = &H1
  BACKUP_CONTAINS_SECURITY = &H2
  BACKUP_CONTAINS_PROPRETIES = &H4
  BACKUP_SPARSE_ATTRIBUTE = &H8
End Enum


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 261
    cAlternate As String * 14
End Type

Public fhandle As Long
    
Public Sub StreamHunter(ByVal file As String)
If quitjob = 0 Then
    Dim i As Long
    Dim Size As Long
    Dim buffer() As Byte
    Dim StatusBlock As IO_STATUS_BLOCK
    Dim streamname As String, StreamSz As Long, StreamData As String
    Dim RetVal As Integer
    Dim lpStreamInfo, cbStreamInfo As Long
    Dim streamInfo As FILE_STREAM_INFO
    Dim temp As String
    
    temp = file
        
    Main.StatFilelbl.Caption = temp
           
    fhandle = CreateFileW(StrPtr(file), 0&, 0&, 0&, FileMode.OpenExisting, 0& Or OpenFile.BackupSemantics, 0&)
    If fhandle = -1 Then Exit Sub
    DoEvents
    DoEvents
    Size = 4096
    RetVal = 234
        
    ReDim buffer(1 To Size)
    
    Do While RetVal = 234
        Dim ptr As Long
        ptr = VarPtr(buffer(1))
        RetVal = QueryFileInfo(fhandle, StatusBlock, ByVal ptr, Size, ByVal FileStreamInformation)
        DoEvents
        DoEvents
        If RetVal = 234 Then
            Size = Size + 4096
            ReDim buffer(Size)
        End If
    Loop
    i = 1
    Set ListObj = Main.StreamList.ListItems.Add(, , temp)
      
    If BigFSO.FileExists(temp) Then
        ScanFiles = ScanFiles + 1
    Else
        ScanFolders = ScanFolders + 1
    End If
    Do
        CopyMemory ByVal VarPtr(streamInfo.NextEntryOffset), ByVal ptr, 24
        CopyMemory ByVal VarPtr(streamInfo.streamname(0)), ByVal ptr + 24, streamInfo.StreamNameLength
        DoEvents
        DoEvents
        streamname = Left$(streamInfo.streamname, streamInfo.StreamNameLength / 2)
        If InStr(1, streamname, "::$DATA", vbTextCompare) = 0 Then 'And InStr(1, streamname, ":encryptable:$DATA", vbTextCompare) = 0 Then                                   ' Add the stream to our stream list, except if it's the default one or encryption, also take off : and :$DATA
            If streamname <> "" Then
            DoEvents
            DoEvents
            If BigFSO.FileExists(temp) Then
                FileWithADS = FileWithADS + 1
                Main.FilesWithADSlbl.Caption = CStr(FileWithADS)
                 DoEvents
                DoEvents
            ElseIf BigFSO.FolderExists(temp) Then
                DirWithADS = DirWithADS + 1
                Main.DirWithADSlbl.Caption = CStr(DirWithADS)
                 DoEvents
            DoEvents
            End If
            
            temp = Mid$(streamname, 2, Len(streamname) - 7)
            Streams = temp & ", " & Streams
            End If
        End If
        If streamInfo.NextEntryOffset Then
            ptr = ptr + streamInfo.NextEntryOffset
            i = i + 1
            DoEvents
            DoEvents
        Else
            If Len(Streams) > 2 Then
                Streams = Left(Streams, Len(Streams) - 2)
            End If
            ListObj.SubItems(1) = Streams
            DoEvents
            DoEvents
            Exit Do
        End If
    Loop
    Main.StreamList.Enabled = True
    'Main.StatFilelbl.Caption = "Files Scanned: " & CStr(ScanFiles) & "                       " & "With ADS :" & CStr(FileWithADS) & vbCrLf & "Folders Scanned :" & CStr(ScanFolders) & "                       " & "With ADS :" & CStr(DirWithADS)
ReDim buffer(0)
CloseHandle fhandle
End If
End Sub

Public Function StreamEnlist(file As String)
Dim streamname() As Byte, buffer() As Byte
Dim W32 As WIN32_STREAM_ID
Dim LObj As ListItem
Dim cbRead As Long, lpContext As Long, LoBytes As Long
Dim HiBytes As Long, cbToSeek As Long, BufferLength As Long, StreamItem As Long

fhandle = CreateFileW(StrPtr(file), FileAccess.FStandardRightsRequired, FileShare.FRead, 0&, FileMode.OpenExisting, 0& Or OpenFile.BackupSemantics, 0&)
ReDim buffer(DefBufferSize - 1)

cbRead = 1
Do While cbRead
    LoBytes = SetFilePointer(fhandle, 0&, 0&, BeginOfFile)
    BackupRead fhandle, VarPtr(buffer(0)), LenB(W32), cbRead, 0&, 0&, lpContext
    If cbRead = 0 Then Exit Do
    CopyMemory ByVal VarPtr(W32), ByVal VarPtr(buffer(0)), LenB(W32)
    With W32
        If .dwStreamNameSize Then
            ReDim streamname(.dwStreamNameSize - 1)
            cbRead = 0
            BackupRead fhandle, VarPtr(buffer(0)) + BufferLength, .dwStreamNameSize, cbRead, 0&, 0&, lpContext
            CopyMemory ByVal VarPtr(streamname(0)), ByVal VarPtr(buffer(0)) + BufferLength, .dwStreamNameSize
            Set LObj = SubStreams.SubStreamList.ListItems.Add(, , Left$(Left$(streamname, .dwStreamNameSize / 2), .dwStreamNameSize / 2 - 6))
            LObj.SubItems(1) = StreamIDToString(.dwStreamID)
            LObj.SubItems(2) = .Size & " Bytes"
            LObj.SubItems(3) = StreamAttributeToString(.dwStreamAttributes)
        End If
        cbToSeek = W32.Size
        BackupSeek fhandle, cbToSeek, 0&, LoBytes, HiBytes, lpContext
        'If LoBytes = 0 Then Exit Do
    End With
Loop

BackupRead fhandle, 0&, 0&, 0&, 1&, 0&, lpContext
CloseHandle fhandle
End Function

Public Function StreamIDToString(ByVal StreamId As FileStreamTypes) As String
Select Case StreamId
    Case BACKUP_EA_DATA
        StreamIDToString = "Extended Data"
    Case BACKUP_ALTERNATE_DATA
        StreamIDToString = "Alternate Data"
    Case BACKUP_ALTERNATE_DATA
        StreamIDToString = "Hard Link"
    Case BACKUP_SECURITY_DATA
        StreamIDToString = "Security Data"
    Case BACKUP_PROPRETY_DATA
        StreamIDToString = "Proprety Data"
    Case BACKUP_OBJECT_ID
        StreamIDToString = "Object ID"
    Case BACKUP_REPARSE_DATA
        StreamIDToString = "Reparse Data"
    Case BACKUP_SPARSE_BLOCK
        StreamIDToString = "Sparse Block"
End Select
End Function

Public Function StreamAttributeToString(ByVal StreamAttribute As FileStreamAttributes) As String
Select Case StreamAttribute
    Case BACKUP_NORMAL_ATTRIBUTE
        StreamAttributeToString = "Normal"
    Case BACKUP_MODIFIED_WHEN_READ
        StreamAttributeToString = "Modified when Read"
    Case BACKUP_CONTAINS_SECURITY
        StreamAttributeToString = "Contains Security Stuff"
    Case BACKUP_CONTAINS_PROPRETIES
        StreamAttributeToString = "Additional Properties"
    Case BACKUP_SPARSE_ATTRIBUTE
        StreamAttributeToString = "Sparse Attribute"
End Select
End Function

Public Function Attacher(origfile As String, newfile As String, NewStreamName As String) As Boolean
On Error GoTo errfound
Dim srcsz As Long
Dim fileno As Integer
Dim finalstream As String
Dim newbuffer() As Byte

finalstream = origfile & ":" & NewStreamName
fileno = FreeFile

With NewStream
.P.Max = 100
Open newfile For Binary As fileno
srcsz = LOF(fileno)
.P.Value = 20
ReDim newbuffer(srcsz)
.P.Value = 25
Get fileno, 1, newbuffer
Close fileno

fileno = FreeFile
Open finalstream For Binary As fileno
Put fileno, , newbuffer
.P.Value = 85
Close fileno
Attacher = True
.P.Value = 100
End With

Exit Function
errfound:
    MsgBox "The new stream could not be created. Try a smaller file size or check your disk for errors using chkdsk.", vbCritical, "Error"
End Function

Public Sub EnumDir(ByVal RootDir As String)
On Error Resume Next
If quitjob = 0 Then
StreamHunter (RootDir)
Dim fso As New FileSystemObject
Dim newfile As file
Dim newfolder As Scripting.Folder
Dim Dir As Scripting.Folder
If fso.FolderExists(RootDir) Then
        
    Set newfolder = fso.GetFolder(RootDir)
        For Each newfile In newfolder.Files
        If quitjob = 1 Then Exit Sub
        DoEvents
        DoEvents
        ScanFiles = ScanFiles + 1
        Main.Filenumlbl.Caption = CStr(ScanFiles)
        StreamHunter (newfile.Path)
        DoEvents
     
        Next
        DoEvents
        For Each Dir In newfolder.SubFolders
            DoEvents
            DoEvents
        If quitjob = 1 Then Exit Sub
            DoEvents
            DoEvents
            ScanFolders = ScanFolders + 1
            Main.DirNumlbl = CStr(ScanFolders)
            RootDir = Dir
            DoEvents
            DoEvents
            EnumDir (Dir)
       
        Next
End If
End If
End Sub

Public Function BrowseFolder(owner As Form) As String
  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  
  szTitle = "Select a folder or drive to scan..."
  With tBrowseInfo
    .hWndOwner = owner.hwnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With

  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Trim(sBuffer)
    sBuffer = Left(sBuffer, Len(sBuffer) - 1)
    
    BrowseFolder = sBuffer
  Else
    BrowseFolder = ""
  End If
  
End Function

Function GetWin32ErrorDescription(ErrorCode As Long) As String

Dim lngRet As Long
Dim strAPIError As String
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

' Preallocate the buffer.
strAPIError = String$(2048, " ")

' Now get the formatted message.
lngRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrorCode, 0, strAPIError, Len(strAPIError), 0)

' Reformat the error string.
strAPIError = Left$(strAPIError, lngRet)

' Return the error string.
GetWin32ErrorDescription = strAPIError

End Function
