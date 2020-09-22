Attribute VB_Name = "modDod"
' Created by Michael Bos april 2000
' For info mail at mb@dds.nl
'
'
' The source may be distributed on every possible way thinkable.
' For Win nt some security constants are required i think ;)
'
' This function on my computer gets about 22 Kb/s for an 8 passer

Option Explicit ' no loose vars

Public Const GENERIC_READ = &H80000000 'Allow the program to read data from the file.
Public Const GENERIC_WRITE = &H40000000 'Allow the program to write data to the file.
Public Const FILE_SHARE_READ = &H1 'Allow other programs to read data from the file.
Public Const FILE_SHARE_WRITE = &H2 'Allow other programs to write data to the file.
Public Const CREATE_ALWAYS = 2 'Create a new file. Overwrite the file (i.e., delete the old one first) if it already exists.
Public Const CREATE_NEW = 1 'Create a new file. The function fails if the file already exists.
Public Const OPEN_ALWAYS = 4 'Open an existing file. If the file does not exist, it will be created.
Public Const OPEN_EXISTING = 3 'Open an existing file. The function fails if the file does not exist.
Public Const TRUNCATE_EXISTING = 5 'Open an existing file and delete its contents. The function fails if the file does not exist.
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20 'An archive file (which most files are).
Public Const FILE_ATTRIBUTE_HIDDEN = &H2 'A hidden file, not normally visible to the user.
Public Const FILE_ATTRIBUTE_NORMAL = &H80 'An attribute-less file (cannot be combined with other attributes).
Public Const FILE_ATTRIBUTE_READONLY = &H1 'A read-only file.
Public Const FILE_ATTRIBUTE_SYSTEM = &H4 'A system file, used exclusively by the operating system.
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000 'Delete the file once it is closed.
Public Const FILE_FLAG_NO_BUFFERING = &H20000000 'Do not use any buffers or caches. If used, the following things must be done: access to the file must begin at whole number multiples of the disk's sector size; the amounts of data accessed must be a whole number multiple of the disk's sector size; and buffer addresses for I/O operations must be aligned on whole number multiples of the disk's sector size.
Public Const FILE_FLAG_OVERLAPPED = &H40000000 'Allow asynchronous I/O; i.e., allow the file to be read from and written to simultaneously. If used, functions that read and write to the file must specify the OVERLAPPED structure identifying the file pointer. Windows 95 does not support overlapped disk files, although Windows NT does.
Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000 'Allow file names to be case-sensitive.
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000 'Optimize the file cache for random access (skipping around to various parts of the file).
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000 'Optimize the file cache for sequential access (starting at the beginning and continuing to the end of the file).
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000 'Bypass any disk cache and instead read and write directly to the

Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hfile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Function DodWipeFile(sPath As String, iPasses As Integer)
Dim lVal, hfile, lNumWritten, i, j As Long
Dim dCount As Double
Dim dRest As Double
Dim sTemp As String
Dim lFileLength As Long
Dim sTemp0, sTemp1 As String
Dim RetVal As Variant
Const CWR_BUFFER = 32768 'write 32Kb blocks

On Error Resume Next

' Calculate some var's
sTemp0 = String(CWR_BUFFER, Chr$(0))
sTemp1 = String(CWR_BUFFER, Chr$(255))
lFileLength = FileLen(sPath)
dCount = Int((lFileLength) / CWR_BUFFER)
dRest = lFileLength - (dCount * CWR_BUFFER)

hfile = CreateFile(sPath, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE Or FILE_FLAG_SEQUENTIAL_SCAN Or FILE_FLAG_DELETE_ON_CLOSE, 0)
If hfile = -1 Then  ' the file could not be opened
GoTo ErrSub
End If
' Do it the Dod way
For lVal = 1 To iPasses ' 8 is the official supersecure number (i think)
    If lFileLength > CWR_BUFFER Then
        For i = 1 To dCount
            RetVal = WriteFile(hfile, ByVal sTemp0, CWR_BUFFER, lNumWritten, 0)
        Next i
    End If
    sTemp = String(dRest, Chr$(0))
    RetVal = WriteFile(hfile, ByVal sTemp, Len(sTemp), lNumWritten, 0)
    sTemp = Empty
        
    RetVal = SetFilePointer(hfile, 0, 0, 0) ' return to the begin of the file
            
    dRest = lFileLength
    If lFileLength > CWR_BUFFER Then
        For i = 1 To dCount
            RetVal = WriteFile(hfile, ByVal sTemp1, CWR_BUFFER, lNumWritten, 0)
        Next i
    End If
    sTemp = String(dRest, Chr$(255))
    RetVal = WriteFile(hfile, ByVal sTemp, Len(sTemp), lNumWritten, 0)
    sTemp = Empty
    
    RetVal = SetFilePointer(hfile, 0, 0, 0) ' return to the begin of the file
            
    dRest = lFileLength
    If lFileLength > CWR_BUFFER Then
        For i = 1 To dCount
            RetVal = WriteFile(hfile, ByVal sTemp1, CWR_BUFFER, lNumWritten, 0)
        Next i
    End If
    sTemp = String(dRest, Chr$(0))
    RetVal = WriteFile(hfile, ByVal sTemp, Len(sTemp), lNumWritten, 0)
    sTemp = Empty
    
    RetVal = SetFilePointer(hfile, 0, 0, 0) ' return to the begin of the file
Next lVal

' Random shit to the file so its really really gone
Randomize
For j = 1 To iPasses

    frmMain.PB1.Value = 0
    frmMain.PB1.Max = iPasses
    frmMain.PB1.Value = j

    dRest = lFileLength
    If lFileLength > CWR_BUFFER Then
        For i = 1 To dCount
            sTemp = String(CWR_BUFFER, Chr(Int(255 * Rnd) + 1))
            RetVal = WriteFile(hfile, ByVal sTemp, CWR_BUFFER, lNumWritten, 0)
            sTemp = Empty
        Next i
    End If
    sTemp = String(dRest, Chr(Int(255 * Rnd) + 1))
    RetVal = WriteFile(hfile, ByVal sTemp, Len(sTemp), lNumWritten, 0)
    sTemp = Empty
    RetVal = SetFilePointer(hfile, 0, 0, 0) ' return to the begin of the file
Next

    If Rename = True Then
    
    Dim RandomName, OldName, NewName As String
    Dim a, b As Long
    
    a = Rnd * 1000 'a & b are random numbers which we
    b = Rnd * 1000 'use to generate a new file name
    
    RandomName = a & "." & b 'the final random name
    'in the format of 111.111.111

    OldName = FileTemp
    NewName = GetPath(FileTemp) & RandomName

    Name OldName As NewName

    End If

RetVal = CloseHandle(NewName) ' close and delete the file


ErrSub:

    If Err.Number = 0 Then

    If Err.Number = 9 Then 'probably hex corrupt generated a figure of ""
    Resume Next

    If Err.Number = 55 Then
    Close #1 'if file is already open, close it
    
    Else
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error"

    End If: End If: End If

End Function
