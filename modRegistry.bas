Attribute VB_Name = "modRegistry"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

    Function CreateContextMenu() As Boolean
    'add an option to 'Delete with TFS' to all files when
    'you right click on them

    On Error GoTo ErrSub
    Dim Path As String
    Dim Ret&
    Dim lphKey&
    
    Path = App.Path
    If Right(Path, 1) <> "\" Then Path = Path & "\"

    Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, "*\shell\Delete With TFS", lphKey&)
    If Ret& <> ERROR_SUCCESS Then Exit Function

    Ret& = RegSetValue&(lphKey&, "", REG_SZ, "Delete With TFS", MAX_PATH)
    If Ret& <> ERROR_SUCCESS Then Exit Function

    Ret& = RegCreateKey&(HKEY_CLASSES_ROOT, "*\shell\Delete With TFS\command", lphKey&)
    If Ret& <> ERROR_SUCCESS Then Exit Function
    
    Ret& = RegSetValue&(lphKey&, "", REG_SZ, Path & App.EXEName & ".exe" & " %1", MAX_PATH)
    If Ret& <> ERROR_SUCCESS Then Exit Function

    CreateContextMenu = True
    Exit Function

ErrSub:

End Function

    Function DeleteContextMenu() As Boolean
    'deletes the option to 'delete with tfs' from files
    
    On Error GoTo ErrSub
    Dim Ret&

    Ret& = RegDeleteKey&(HKEY_CLASSES_ROOT, "*\shell\Delete With TFS\command")
    If Ret& <> ERROR_SUCCESS Then Exit Function

    Ret& = RegDeleteKey&(HKEY_CLASSES_ROOT, "*\shell\Delete With TFS")
    If Ret& <> ERROR_SUCCESS Then Exit Function

    DeleteContextMenu = True
    Exit Function

ErrSub:

End Function
