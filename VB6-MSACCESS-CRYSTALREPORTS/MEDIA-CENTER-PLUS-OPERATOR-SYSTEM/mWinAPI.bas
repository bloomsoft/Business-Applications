Attribute VB_Name = "mWinAPI"
'***********************************************************************************
'The CreateRoundRectRgn function creates a rectangular region with rounded corners.
'***********************************************************************************
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
ByVal Y3 As Long) As Long

'***********************************************************************************
'The SetWindowRgn function sets the window region of a window.
'The system does not display any portion of a window that lies outside of the window region.
'***********************************************************************************
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'***********************************************************************************
'API used to retrieve the ComputerName of the current local system.-----------------
'***********************************************************************************
Public Declare Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'**********************************************************************************
'The Find***File function searches a directory for a file whose name matches the
'specified filename or folder. Find***File examines subdirectory names as well as filenames.
'**********************************************************************************
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
(ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260
Private Const AMAX_PATH = 260

Public Type FILETIME                        'Structure FILETIME.
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA                  'Structure WIN32_FIND_DATA.
    dwFileAttributes As Long                 'Specifies the file attributes of the file found.
    ftCreationTime As FILETIME               'Specifies a FILETIME structure containing the time the file was created.
    ftLastAccessTime As FILETIME             'Specifies a FILETIME structure containing the time that the file was last accessed.
    ftLastWriteTime As FILETIME              'Specifies a FILETIME structure containing the time that the file was last written to.
    nFileSizeHigh As Long                    'Specifies the high-order DWORD value of the file size, in bytes.
    nFileSizeLow As Long                     'Specifies the low-order DWORD value of the file size, in bytes.
    dwReserved0 As Long                      'If the dwFileAttributes member includes the FILE_ATTRIBUTE_REPARSE_POINT attribute, this member specifies the reparse tag. Otherwise, this value is undefined and should not be used.
    dwReserved1 As Long                      'Reserved for future use.
    cFileName As String * MAX_PATH           'A null-terminated string that is the name of the file.
    cAlternateFileName As String * AMAX_PATH 'A null-terminated string that is an alternative name for the file.
End Type

Public Enum FILE_ATTRIBUTES           'Self explanitary
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum

'**********************************************************************************
'The mciSendString function sends a command string to an MCI device.
'The device that the command is sent to is specified in the command string.
'**********************************************************************************
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'**********************************************************************************
'The mciSendCommand function sends a command message to the specified MCI device.
'**********************************************************************************
Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

'**********************************************************************************
'The mciGetErrorString function retrieves a string that describes the specified MCI error code.
'**********************************************************************************
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

'**********************************************************************************
'The GetCursorPos function retrieves the cursor's position, in screen coordinates.
'**********************************************************************************
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Public Type POINTAPI    'Structure for POINTAPI.
    lxPos As Long
    lyPos As Long
End Type

'**********************************************************************************
'The GetShortPathName function obtains the short path form of a specified input path.
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'**********************************************************************************

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 21
End Type
    
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205
Public SysIcon As NOTIFYICONDATA
Public sysIcon2 As NOTIFYICONDATA

