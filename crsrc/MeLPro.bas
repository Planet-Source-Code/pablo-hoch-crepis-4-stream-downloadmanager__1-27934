Attribute VB_Name = "MeLPro"
Const MeLProModuleVersion = "0.1a"
' -~+=|[ MeLProModule ]|=+~-
' Dieses Modul stammt von Pablo Hoch aka -=MeL=-
' http://www.melaxis.de Support: mel@melaxis.de
' Dieses Modul darf nur mit Erlaubnis des Autors verwendet werden.
' Bei Fragen wenden Sie sich bitte per E-Mail an mel@melaxis.de
' Nur für >>Microsoft Visual Basic 6.0<<
' Ältere Versionen von VB können mit einigen verwendeten Funktionen nicht
' umgehen. Laufzeitfehler werden um jeden Preis vermieden!!


Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)


Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FileTime) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, lpcbSecurityDescriptor As Long) As Long

Public Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
Public Const REG_EXPAND_SZ = 2                  ' Null-terminierte Unicode-Zeichenfolge
Public Const REG_DWORD = 4                      ' 32-Bit-Zahl
Public Const REG_OPTION_NON_VOLATILE = 0       ' Schlüssel bleibt beim Neustart erhalten
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0
Public Const ERROR_NO_MORE_ITEMS = 259&
Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As ACL
        Dacl As ACL
End Type
Global Const SCHWARZ = vbBlack
Global Const BLAU = vbBlue
Global Const CYAN = vbCyan
Global Const GRÜN = vbGreen
Global Const ROT = vbRed
Global Const WEISS = vbWhite
Global Const GELB = vbYellow
Global Const GRAU = &H8000000F


Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd _
As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Global Const AW_HOR_POSITIVE = &O1 ' Animate the window from
'left to right. This flag can be used with roll or slide
'animation It is ignored when used with the AW_CENTER flag.
Global Const AW_HOR_NEGATIVE = &H2 ' Animate the window from
'right to left. This flag can be used with roll or slide
'animation. It is ignored when used with the AW_CENTER flag.
Global Const AW_VER_POSITIVE = &H4 ' Animate the window from
'top to bottom. This flag can be used with roll or slide
'animation. It is ignored when used with the AW_CENTER flag.
Global Const AW_VER_NEGATIVE = &H8 ' Animate the window from
'bottom to top. This flag can be used with roll or slide
'animation. It is ignored when used with the AW_CENTER flag.
Global Const AW_CENTER = &H10 ' Makes the window appear to
'collapse inward if the AW_HIDE flag is used or expand outward
'if the AW_HIDE flag is not used.
Global Const AW_HIDE = &H10000 ' Hides the window. By default,
'the window is shown.
Global Const AW_ACTIVATE = &H20000 ' Activates the window. Do
'not use this flag with AW_HIDE.
Global Const AW_SLIDE = &H40000 ' Uses slide animation. By
'default, roll animation is used. This flag is ignored when used
'with the AW_CENTER flag.
Global Const AW_BLEND = &H80000 ' Uses a fade effect. This flag
'can be used only if hwnd is a top-level window.

Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_MEMORY = &H4       '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000    '  name is a WIN.INI [sounds] entry
Public Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0    '  must be > 4096 to keep strings in same section of resource file
Public Const SND_APPLICATION = &H80 '  look for application specific association
Public Const SND_FILENAME = &H20000 '  name is a file name
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NODEFAULT = &H2    '  silence not default, if sound not found
Public Const SND_NOSTOP = &H10      '  don't stop any currently playing sound
Public Const SND_NOWAIT = &H2000    '  don't wait if the driver is busy
Public Const SND_PURGE = &H40        '  purge non-static events for task
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_TYPE_MASK = &H170007   ' ?
Public Const SND_VALIDFLAGS = &H17201F '  Set of valid flag bits.  Anything outside this range will raise an error
Public Const SND_VALID = &H1F       '  valid flags          / ;Internal /

Public Type melAppInfo
    dateiEndung As String
    programmName As String
    ContentType As String
    dateiBeschreibung As String
    fileIcon As String
    fileCommand As String
    prevApp As String
End Type

'Public fso As New FileSystemObject ' siehe Verweise: MS Scripting
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    ' ShellExecute Me.hWnd, "Open", "mailto:mel@melaxis.de", "", "", 1

Public Type PointAPI
        x As Long
        y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const WM_DESTROY = &H2
Public Const WM_NCDESTROY = &H82

Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Const ALTERNATE = 1
Public Const WINDING = 2

'Public Declare Function mciSendString Lib "winmm" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnStr As Any, ByVal wReturnLen As Long, ByVal _
hCallBack As Long) As Long
Public Declare Function mciExecute Lib "WINMM.DLL" (ByVal lpstrCommand As String) As Long
Public Const SW_NORMAL = 1

Public Declare Function MapVirtualKey Lib "user32" _
  Alias "MapVirtualKeyA" (ByVal wCode As Long, _
  ByVal wMapType As Long) As Long
Public Declare Sub keybd_event Lib "user32" _
  (ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_MENU = &H12
Public Const VK_SNAPSHOT = &H2C
Public Const KEYEVENTF_KEYUP = &H2

Public Declare Function ExtractIcon Lib "shell32.dll" _
    Alias "ExtractIconA" (ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long
    
Public Declare Function DrawIcon Lib "user32" _
   (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal hIcon As Long) As Long

Public Declare Function DestroyIcon Lib "user32" _
    (ByVal hIcon As Long) As Long

Public Const MAX_PATH = 260
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const SYNCHRONIZE = &H100000
Public Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long

Private x1a0(9) As Long
Private cle(17) As Long
Private x1a2 As Long

Private inter As Long, res As Long, ax As Long, bx As Long
Private cX As Long, dx As Long, si As Long, tmp As Long
Private i As Long, c As Byte

Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_UNLOAD = (11000 + 22)
Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const MAX_IP_STATUS = 11000 + 50
Public Const IP_PENDING = (11000 + 255)
'Public Const PING_TIMEOUT = 200
'Public Const PING_TIMEOUT = 5000
'Public Const PING_TIMEOUT = 3000
Public PING_TIMEOUT As Long
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const INADDR_NONE = &HFFFFFFFF

Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const MAXGETHOSTSTRUCT = 1024

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type


Public Type hostent_async
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
    h_asyncbuffer(MAXGETHOSTSTRUCT) As Byte
End Type

Public hostent_async As hostent_async



Public Const hostent_size = 16



Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Public Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal TimeOut As Long) As Long
    
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
    
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long
    
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, _
    ByVal hpvSource As Long, _
    ByVal cbCopy As Long)


    Public Declare Function htonl Lib "WSOCK32.DLL" (ByVal hostlong As Long) As Long

    Public Declare Function htons Lib "WSOCK32.DLL" (ByVal hostshort As Long) As Integer

    Public Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long

    Public Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long

    Public Declare Function ntohl Lib "WSOCK32.DLL" (ByVal netlong As Long) As Long

    Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Integer

'Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public RETURNCODE As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Public Const VK_SNAPSHOT = &H2C
'Public Const KEYEVENTF_KEYUP = &H2
'Public Const VK_MENU = &H12
Public PointerToPointer, IPLong As Long

Public Const SSM_Desktop = 1
Public Const SSM_ActiveWindow = 0
Const WinHeight = "@Height@"
Const WinWidth = "@Width@"
Const WinTop = "@Top@"
Const WinLeft = "@Left@"

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

Public Const OFS_DEFAULT_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_FILEMUSTEXIST _
             Or OFN_HIDEREADONLY _
             Or OFN_NODEREFERENCELINKS

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
 End Type


Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function MakePath Lib "imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal lpPath As String) As Long
Public Enum SpecialShellFolderIDs
   CSIDL_DESKTOP = &H0
   CSIDL_INTERNET = &H1
   CSIDL_PROGRAMS = &H2
   CSIDL_CONTROLS = &H3
   CSIDL_PRINTERS = &H4
   CSIDL_PERSONAL = &H5
   CSIDL_FAVORITES = &H6
   CSIDL_STARTUP = &H7
   CSIDL_RECENT = &H8
   CSIDL_SENDTO = &H9
   CSIDL_BITBUCKET = &HA
   CSIDL_STARTMENU = &HB
   CSIDL_DESKTOPDIRECTORY = &H10
   CSIDL_DRIVES = &H11
   CSIDL_NETWORK = &H12
   CSIDL_NETHOOD = &H13
   CSIDL_FONTS = &H14
   CSIDL_TEMPLATES = &H15
   CSIDL_COMMON_STARTMENU = &H16
   CSIDL_COMMON_PROGRAMS = &H17
   CSIDL_COMMON_STARTUP = &H18
   CSIDL_COMMON_DESKTOPDIRECTORY = &H19
   CSIDL_APPDATA = &H1A
   CSIDL_PRINTHOOD = &H1B
   CSIDL_ALTSTARTUP = &H1D           ' // DBCS
   CSIDL_COMMON_ALTSTARTUP = &H1E    ' // DBCS
   CSIDL_COMMON_FAVORITES = &H1F
   CSIDL_INTERNET_CACHE = &H20
   CSIDL_COOKIES = &H21
   CSIDL_HISTORY = &H22
End Enum

Public Declare Function SHGetPathFromIDList _
Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
   ByVal pszPath As String) As Long



Public Declare Function SHGetSpecialFolderLocation _
Lib "shell32.dll" _
   (ByVal hwndOwner As Long, _
   ByVal nFolder As SpecialShellFolderIDs, _
   pidl As Long) As Long
   
Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
(ByVal pv As Long)

Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Declare Function GetWindowTextLength Lib "user32" _
Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function GetNextWindow Lib "user32" _
Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) _
As Long

Public Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hwnd As Long, lpdwProcessId As Long) As Long


Public Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" _
  (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function RegisterServiceProcess Lib _
    "kernel32.dll" (ByVal dwProcessId As Long, _
     ByVal dwType As Long) As Long
     
Public Declare Function GetCurrentProcessId Lib _
   "kernel32.dll" () As Long
   
Type ctrObj
    Name As String
    Index As Long
    Parrent As String
    Top As Long
    Left As Long
    Height As Long
    Width As Long
    ScaleHeight As Long
    ScaleWidth As Long
End Type

Private FormRecord() As ctrObj
Private ControlRecord() As ctrObj
Private MaxForm As Long
Private MaxControl As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private pUdtMemStatus As MEMORYSTATUS

Public Declare Sub GlobalMemoryStatus Lib _
"kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

    Public Const GENERIC_WRITE = &H40000000
    Public Const GENERIC_READ = &H80000000
    Public Const FILE_SHARE_READ = &H1
    Public Const FILE_SHARE_WRITE = &H2
    Public Const OPEN_EXISTING = 3
    'Public Const FILE_ATTRIBUTE_NORMAL = &H80
    Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
    Public Const EXCEPTION_CONTINUE_EXECUTION = -1
    Public Const EXCEPTION_CONTINUE_SEARCH = 0
    Public Const EXCEPTION_DEBUG_EVENT = 1
    Public Const EXCEPTION_EXECUTE_HANDLER = 1
    Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15
    
Public Const UA = "A"
Public Const UB = "B"
Public Const UC = "C"
Public Const UD = "D"
Public Const UE = "E"
Public Const UF = "F"
Public Const UG = "G"
Public Const UH = "H"
Public Const UI = "I"
Public Const UJ = "J"
Public Const UK = "K"
Public Const ul = "L"
Public Const UM = "M"
Public Const UN = "N"
Public Const UO = "O"
Public Const UP = "P"
Public Const UQ = "Q"
Public Const UR = "R"
Public Const US = "S"
Public Const ut = "T"
Public Const UU = "U"
Public Const UV = "V"
Public Const UW = "W"
Public Const UX = "X"
Public Const UY = "Y"
Public Const UZ = "Z"
Public Const LA = "a"
Public Const LB = "b"
Public Const LC = "c"
Public Const LD = "d"
Public Const le = "e"
Public Const LF = "f"
Public Const LG = "g"
Public Const LH = "h"
Public Const LI = "i"
Public Const LJ = "j"
Public Const LK = "k"
Public Const LL = "l"
Public Const LM = "m"
Public Const ln = "n"
Public Const LO = "o"
Public Const lp = "p"
Public Const LQ = "q"
Public Const LR = "r"
Public Const ls = "s"
Public Const LT = "t"
Public Const LU = "u"
Public Const LV = "v"
Public Const LW = "w"
Public Const LX = "x"
Public Const LY = "y"
Public Const LZ = "z"

Public Const RSP_SIMPLE_SERVICE = &H1
Public Const RSP_UNREGISTER_SERVICE = &H0
'Option Compare Text
Private Crc32Table(255) As Long

Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetKeyPress Lib "user32" Alias "GetAsyncKeyState" (ByVal Key As Long) As Integer
'I had to alias the ReadProcessMemory API because VB6 thinks "ReadProcessMemory" is ambiguous
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Const MAX_HOSTNAME_LEN = 132
Public Const MAX_DOMAIN_NAME_LEN = 132
Public Const MAX_SCOPE_ID_LEN = 260
Public Const MAX_ADAPTER_NAME_LENGTH = 260
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Public Const ERROR_BUFFER_OVERFLOW = 111
Public Const MIB_IF_TYPE_ETHERNET = 1
Public Const MIB_IF_TYPE_TOKENRING = 2
Public Const MIB_IF_TYPE_FDDI = 3
Public Const MIB_IF_TYPE_PPP = 4
Public Const MIB_IF_TYPE_LOOPBACK = 5
Public Const MIB_IF_TYPE_SLIP = 6

Type IP_ADDR_STRING
            Next As Long
            IpAddress As String * 16
            IpMask As String * 16
            Context As Long
End Type

Type IP_ADAPTER_INFO
            Next As Long
            ComboIndex As Long
            AdapterName As String * MAX_ADAPTER_NAME_LENGTH
            Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
            AddressLength As Long
            Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
            Index As Long
            Type As Long
            DhcpEnabled As Long
            CurrentIpAddress As Long
            IpAddressList As IP_ADDR_STRING
            GatewayList As IP_ADDR_STRING
            DhcpServer As IP_ADDR_STRING
            HaveWins As Boolean
            PrimaryWinsServer As IP_ADDR_STRING
            SecondaryWinsServer As IP_ADDR_STRING
            LeaseObtained As Long
            LeaseExpires As Long
End Type

Type FIXED_INFO
            HostName As String * MAX_HOSTNAME_LEN
            DomainName As String * MAX_DOMAIN_NAME_LEN
            CurrentDnsServer As Long
            DnsServerList As IP_ADDR_STRING
            NodeType As Long
            ScopeId  As String * MAX_SCOPE_ID_LEN
            EnableRouting As Long
            EnableProxy As Long
            EnableDns As Long
End Type

Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long

Public Type tm
    tm_sec As Long ' seconds (0 - 59)
    tm_min As Long ' minutes (0 - 59)
    tm_hour As Long ' hours (0 - 23)
    tm_mday As Long ' day of month (1 - 31)
    tm_mon As Long ' month of year (0 - 11)
    tm_year As Long ' year - 1900
    tm_wday As Long ' day of week (Sunday = 0), Not used
    tm_yday As Long ' day of year (0 - 365), Not used
    tm_isdst As Long ' Daylight Savings Time (0, 1), Not used
End Type
    
 Public Declare Function LoadLibrary Lib "kernel32" Alias _
  "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" _
  (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CreateThread Lib "kernel32" _
  (lpThreadAttributes As Any, ByVal dwStackSize As Long, _
  ByVal lpStartAddress As Long, ByVal lParameter As Long, _
  ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Sub ExitThread Lib "kernel32" _
  (ByVal dwExitCode As Long)
Public Declare Function GetExitCodeThread Lib "kernel32" _
  (ByVal hThread As Long, lpExitCode As Long) As Long


Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const PROCESS_DUP_HANDLE = &H40

Public Enum PP
    NORMAL_PRIORITY = &H20
    IDLE_PRIORITY = &H40
    HIGH_PRIORITY = &H80
    REALTIME_PRIORITY = &H100
End Enum

Public Declare Function SetPriorityClass& Lib "kernel32" (ByVal hProcess As Long, _
    ByVal dwPriorityClass As Long)

Public Type FILEVERSIONINFO       'One invented by me
  Path              As String
  FileName          As String
  Filesize          As String
  OSType            As String
  BinState          As String
  FileCreated       As String
  FileLastWritten   As String
  FileLastRead      As String
  CompanyName       As String
  FileDescription   As String
  FileVersion       As String
  InternalName      As String
  LegalCopyright    As String
  OriginalFileName  As String
  ProductName       As String
  ProductVersion    As String
End Type

  'MICROSOFT STRUCTURES
Private Const OF_READ = &H0
Private Const OF_SHARE_DENY_NONE = &H40

Private Type OFSTRUCTREC
  cBytes As Byte
  fFixedDisk As Byte
  nErrCode As Integer
  Reserved1 As Integer
  Reserved2 As Integer
  szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type FILETIMEREC
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type SYSTEMTIMEREC
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
  dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
  dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
  dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
  dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
  dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
  dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
  dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
  dwFileType As Long             '  e.g. VFT_DRIVER
  dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
  dwFileDateMS As Long           '  e.g. 0
  dwFileDateLS As Long           '  e.g. 0
End Type

'Operating system
Public Const VOS__BASE = &H0&
Public Const VOS_UNKNOWN = &H0&
Public Const VOS__WINDOWS16 = &H1&
Public Const VOS__PM16 = &H2&
Public Const VOS__PM32 = &H3&
Public Const VOS__WINDOWS32 = &H4&
Public Const VOS_DOS = &H10000
Public Const VOS_DOS_WINDOWS16 = &H10001
Public Const VOS_DOS_WINDOWS32 = &H10004
Public Const VOS_OS216 = &H20000
Public Const VOS_OS216_PM16 = &H20002
Public Const VOS_OS232 = &H30000
Public Const VOS_OS232_PM32 = &H30003
Public Const VOS_NT = &H40000
Public Const VOS_NT_WINDOWS32 = &H40004

'FileState
Public Const VS_FF_DEBUG = &H1&
Public Const VS_FF_PRERELEASE = &H2&
Public Const VS_FF_PATCHED = &H4&
Public Const VS_FF_PRIVATEBUILD = &H8&
Public Const VS_FF_INFOINFERRED = &H10&
Public Const VS_FF_SPECIALBUILD = &H20&
Public Const VS_FFI_FILEFLAGSMASK = &H3F&
 
    'KERNEL32.DLL FUNCTIONS


Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
  (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Declare Function FileTimeToSystemTime Lib "kernel32" _
  (lpFileTime As FILETIMEREC, lpSystemTime As SYSTEMTIMEREC) As Long

Public Declare Function GetFileTime Lib "kernel32" _
  (ByVal hFile As Long, lpCreationTime As FILETIMEREC, _
   lpLastAccessTime As FILETIMEREC, lpLastWriteTime As FILETIMEREC) As Long

Public Declare Function OpenFile Lib "kernel32" _
  (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCTREC, ByVal wStyle As Long) As Long

Public Declare Function hread Lib "kernel32" Alias "_hread" _
  (ByVal hFile As Long, lpBuffer As Any, ByVal lBYTES As Long) As Long

Public Declare Function lclose Lib "kernel32" Alias "_lclose" _
  (ByVal hFile As Long) As Long
  
    'VERSION.DLL FUNCTIONS
Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long

Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
  (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Private Enum HuffmanTreeNodeParts
    htnWeight = 1
    htnIsLeaf = 2
    htnAsciiCode = 3
    htnBitCode = 4
    htnLeftSubtree = 5
    htnRightSubtree = 6
End Enum

Public Declare Function GetVolumeInformation Lib _
   "kernel32.dll" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As String, _
   ByVal nVolumeNameSize As Integer, _
   lpVolumeSerialNumber As Long, _
   lpMaximumComponentLength As Long, _
   lpFileSystemFlags As Long, _
   ByVal lpFileSystemNameBuffer As String, _
   ByVal nFileSystemNameSize As Long) As Long
   


Private Const NUM_THREADS = "KERNEL\Threads"
Private Const STAT_DATA = "PerfStats\StatData"
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Dim checkType As Integer
Dim remMsg(2) As String
'Working with wininet.dll declarations and constants
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long 'Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
'Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long 'this function used with IE4
Private Const INTERNET_CONNECTION_MODEM = &H1&
Private Const INTERNET_CONNECTION_LAN = &H2&
Private Const INTERNET_CONNECTION_PROXY = &H4&
Private Const INTERNET_RAS_INSTALLED = &H10&
Private Const INTERNET_CONNECTION_OFFLINE = &H20&
Private Const INTERNET_CONNECTION_CONFIGURED = &H40&


Public xS As Integer
Public nameII(0 To 100) As String

Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
(lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias _
"GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, _
lpTotalNumberOfFreeBytes As Currency) As Long
Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long

Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FileTime) As Long
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long



Public Const REG_BINARY = 3                                        ' Free form binary
Public Const REG_CREATED_NEW_KEY = &H1               ' New Registry Key created
Public Const REG_DWORD_BIG_ENDIAN = 5                   ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4               ' 32-bit number (same as REG_DWORD)
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_LINK = 6                                               ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                                       ' Multiple Unicode strings
Public Const REG_NONE = 0                                             ' No value type
Public Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Public Const REG_NOTIFY_CHANGE_LAST_SET = &H4     ' Time stamp
Public Const REG_NOTIFY_CHANGE_NAME = &H1            ' Create or delete (child)
Public Const REG_OPENED_EXISTING_KEY = &H2            ' Existing Key opened
Public Const REG_NOTIFY_CHANGE_SECURITY = &H8
Public Const REG_OPTION_BACKUP_RESTORE = 4          ' Open for backup or restore
Public Const REG_OPTION_CREATE_LINK = 2                   ' Created key is a symbolic link
Public Const REG_OPTION_RESERVED = 0                       ' Parameter is reserved
Public Const REG_OPTION_VOLATILE = 1                          ' Key is not preserved when system is rebooted
Public Const REG_REFRESH_HIVE = &H2                          ' Unwind changes to last flush
Public Const REG_RESOURCE_LIST = 8                             ' Resource list in the resource map
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Public Const REG_WHOLE_HIVE_VOLATILE = &H1            ' Restore whole hive volatile
Public Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Public Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

Public Type HTMLCryptedText
    Text As String
    Number As Integer
    Password As String
End Type


Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long


Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long


Public Enum enHookTypes
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_DEBUG = 9
    WH_FOREGROUNDIDLE = 11
    WH_GETMESSAGE = 3
    WH_HARDWARE = 8
    WH_JOURNALPLAYBACK = 1
    WH_JOURNALRECORD = 0
    WH_MOUSE = 7
    WH_MSGFILTER = (-1)
    WH_SHELL = 10
    WH_SYSMSGFILTER = 6
    WH_KEYBOARD_LL = 13
    WH_MOUSE_LL = 14
    WH_KEYBOARD = 2
End Enum

Public Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long


Public Declare Function GlobalSize Lib "kernel32" (ByVal hmem As Long) As Long


Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hmem As Long) As Long


Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long


Public Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
    Private mMyData() As Byte
    Private mMyDataSize As Long
    Private mHmem As Long


Public Enum enGlobalmemoryAllocationConstants
    GMEM_FIXED = &H0
    GMEM_DISCARDABLE = &H100
    GMEM_MOVEABLE = &H2
    GMEM_NOCOMPACT = &H10
    GMEM_NODISCARD = &H20
    GMEM_ZEROINIT = &H40
End Enum

Private Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(8) As Byte
End Type

Public Declare Function CoCreateGuid Lib "ole32.dll" ( _
     pguid As GUID) As Long

Public Declare Function StringFromGUID2 Lib "ole32.dll" ( _
     rguid As Any, _
     ByVal lpstrClsId As Long, _
     ByVal cbMax As Long) As Long
   
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function InternetGetConnectedState _
    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long
    'Local system uses a modem to connect to
    '     the Internet.
    ''Public Const INTERNET_CONNECTION_MODEM As Long = &H1
    'Local system uses a LAN to connect to t
    '     he Internet.
    ''Public Const INTERNET_CONNECTION_LAN As Long = &H2
    'Local system uses a proxy server to con
    '     nect to the Internet.
    ''Public Const INTERNET_CONNECTION_PROXY As Long = &H4
    'No longer used.
    Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
    ''Public Const INTERNET_RAS_INSTALLED As Long = &H10
    ''Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
    ''Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
    'InternetGetConnectedState wrapper funct
    '     ions


Public Const lcKEY = "9182736450zaybxcwdveuftgshriqjpkolmnZAYBXCWDVEUFTGSHRIQJPKOLMN .,-!?_=+/*#~'""§$%&()[]{}äöüßÄÜÖ\@€:;^°<>|"
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FileTime, lpLastAccessTime As FileTime, lpLastWriteTime As FileTime) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FileTime) As Long
Public Const OF_WRITE = &H1
Public Const OF_SHARE_DENY_READ = &H30
Public Const OF_SHARE_DENY_WRITE = &H20
Public Const OF_SHARE_EXCLUSIVE = &H10
Public Const OF_READWRITE = &H2
Public Const OF_CREATE = &H1000
Public Const OF_DELETE = &H200
Public Declare Function SetFocusWindow Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWNA = 8
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
'Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
'Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub WinExecError Lib "shell32.dll" Alias "WinExecErrorA" (ByVal hwnd As Long, ByVal error As Long, ByVal lpstrFileName As String, ByVal lpstrTitle As String)
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Type TPROC
    hmem As Long
    vtPtr As Long
End Type
Public aProc() As TPROC



'
'Public Type HOSTENT
'    hName As Long
'    hAliases As Long
'    hAddrType As Integer
'    hLen As Integer
'    hAddrList As Long
'End Type


'Public Type WSADATA
'    wVersion As Integer
'    wHighVersion As Integer
'    szDescription(0 To MAX_WSADescription) As Byte
'    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
'    wMaxSockets As Integer
'    wMaxUDPDG As Integer
'    dwVendorInfo As Long
'End Type



































Public Function isStart(inputText As String, searchFor As String) As Boolean
On Error Resume Next
' Diese Funktion gibt zurück, ob am Anfang von inputText searchFor steht
' ©reated 21/09/00

If Mid(inputText, 1, Len(searchFor)) = searchFor Then
    isStart = True
Else
    isStart = False
End If
End Function

Public Function FileExist(testFile As String) As Boolean
On Error Resume Next
' Diese Funktion überprüft, ob die Datei testFile gefunden werden kann.
' ©reated 21/09/00
' updated 28/09/00

'If fso.FileExists(testFile) Then
'    fileExist = True
'Else
'    fileExist = False
'End If

On Error GoTo gibtsNicht
Dim fn As Long
fn = FreeFile
Open testFile For Input As #fn
Close #fn
FileExist = True
'If Dir(testFile) <> "" Then fileExist = True
Exit Function

gibtsNicht:
FileExist = False

End Function

Public Function cutStart(inputText As String, cutThis As String) As String
On Error Resume Next
' Diese Funktion entfernt aus einem String den Angegebenen Text, sofern
' sich dieser am Anfang des Strings befindet.
' Z.B. Nützlich zum ermitteln von Werten aus Konfiguartionsdateien.
' ©reated 21/09/00

If isStart(inputText, cutThis) = True Then
    cutStart = Mid(inputText, Len(cutThis) + 1)
Else
    cutStart = ""
End If
End Function

Public Function cutPathName(FileName As String) As String
On Error Resume Next
' Diese Funktion entfernt aus einem Dateinamen den Pfad.
' Aus c:\windows\win.ini wird z.B. win.ini
' ©reated 21/09/00

Dim melFileName As String, melFileCuts
melFileName = Replace(FileName, "/", "\")
melFileCuts = Split(melFileName, "\")
cutPathName = melFileCuts(UBound(melFileCuts))
End Function

Public Function getPathName(FileName As String) As String
On Error Resume Next
' Diese Funktion liefert aus einem Dateinamen den Pfad.
' Aus c:\windows\win.ini wird z.B. c:\windows\
' ©reated 02/06/01

Dim melFileName As String, melFileCuts
melFileName = Replace(FileName, "/", "\")
melFileCuts = Split(melFileName, "\")
getPathName = Mid(melFileName, 1, Len(melFileName) - Len(melFileCuts(UBound(melFileCuts))))
End Function


Public Function getName(inputText As String) As String
On Error Resume Next
' Diese Funktion ermittelt aus einem Name=Wert Paar den Namen.
' ©reated 21/09/00

Dim melNameArray
melNameArray = Split(inputText, "=")
getName = melNameArray(0)
End Function

Public Function getValue(inputText As String) As String
On Error Resume Next
' Diese Funktion ermittelt aus einem Name=Wert Paar den Wert.
' ©reated 21/09/00

Dim melPairValue As String, melPairName As String
melPairName = getName(inputText)
If isStart(inputText, melPairName & "=") = True Then
    getValue = cutStart(inputText, melPairName & "=")
End If
End Function

Public Function XORCrypt(inputText As String, cryptPassWord As String)
On Error Resume Next
' Diese Funktion ver-/entschlüsselt einen String mit einem Passwort.
' Bekannt als XOR-Verschlüsselung. 64-Bit.
' ©reated 21/09/00

Dim melPosS As Long, melPosC As Long, melTempString As String
melTempString = Space(Len(inputText))
melPosC = 1
For melPosS = 1 To Len(inputText)
    If melPosC > Len(cryptPassWord) Then melPosC = 1
    Mid(melTempString, melPosS, 1) = Chr(Asc(Mid(inputText, melPosS, 1)) Xor Asc(Mid(cryptPassWord, melPosC, 1)))
    If Asc(Mid(melTempString, melPosS, 1)) = 0 Then Mid(melTempString, melPosS, 1) = Mid(inputText, melPosS, 1)
    melPosC = melPosC + 1
Next
XORCrypt = melTempString
End Function

Public Sub Delay(sekundenWarten As Single)
On Error Resume Next
' Diese Funktion hält die Sub/Function (nicht das ganze Programm) für eine
' bestimmte Anzahl von Sekunden an.
' ©reated UNKNOWN
Dim sglStart As Single
sglStart = Timer
Do While Timer < sglStart + sekundenWarten
    DoEvents
Loop
End Sub


Public Function m175CCode(derText As String) As String
On Error Resume Next
' Diese Funktion kodiert einen String mit der M175C-Kodierung von -=MeL=-.
' Diese Funktion sowie die Verschlüsselung wurden vor einiger Zeit
' von Pablo Hoch aka -=MeL=- erfunden.
' ©reated UNKNOWN
Dim i
Dim k1, k2, t1, t2
Dim Fertig
Fertig = ""
For i = 0 To Len(derText)
    t1 = Mid(derText, i, 1)
    t2 = Asc(t1)
    k1 = t2 + 1
    k2 = Chr(k1)
    Fertig = Fertig & k2
Next
If Chr(Mid(Fertig, 1, 1)) = 10 Then Fertig = Mid(Fertig, 2)
m175CCode = Fertig
End Function

Public Function m175CDecode(derText As String) As String
On Error Resume Next
' Diese Funktion dekodiert einen String mit der M175C-Kodierung von -=MeL=-.
' Diese Funktion sowie die Verschlüsselung wurden vor einiger Zeit
' von Pablo Hoch aka -=MeL=- erfunden.
' ©reated UNKNOWN
Dim i, c
Dim k1, k2, t1, t2
Dim Fertig
Fertig = ""
c = 0
If Mid(derText, 1, 5) = "M175C" Then derText = Mid(derText, 6): c = 1
For i = 0 To Len(derText)
    t1 = Mid(derText, i, 1)
    t2 = Asc(t1)
    k1 = t2 - 1
    k2 = Chr(k1)
    Fertig = Fertig & k2
Next
m175CDecode = Fertig
End Function

Public Function passM175C(inputText As String) As String
On Error Resume Next
' Diese Funktion prüft, ob ein String mit der M175C-Verschlüsselung kodiert ist.
' Falls ja, wird er dekodiert zurückgegeben, wenn er nicht kodiert ist,
' wird er ganz normal zurückgeliefert.
' ©reated 21/09/00

If isStart(inputText, "M175C") = True Then
    passM175C = m175CDecode(inputText)
Else
    passM175C = inputText
End If
End Function

Public Function ReadINI(iniDatei As String, sSection As String, sKeyName As String) As String
On Local Error Resume Next
' Diese Funktion ermittelt den Wert einer Einstellung aus einer INI-Datei.
' WICHTIG: Beim Dateinamen muss der Pfad angegeben werden!
' ©reated UNKNOWN
Dim sRet As String
sRet = String(255, Chr(0))
ReadINI = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), iniDatei))
End Function

Public Function GetINI(iniDatei As String, sSection As String, sKeyName As String) As String
On Error Resume Next
' Diese Funktion ist nur zum Weiterleiten gedacht.
' ©reated UNKNOWN
GetINI = ReadINI(iniDatei, sSection, sKeyName)
End Function

Public Function WriteINI(iniDatei As String, sSection As String, sKeyName As String, sNewString As String) As Boolean
On Local Error Resume Next
' Diese Funktion ändert den Wert einer Einstellung in einer INI-Datei.
' WICHTIG: Beim Dateinamen muss der Pfad angegeben werden!
' ©reated UNKNOWN
Call WritePrivateProfileString(sSection, sKeyName, sNewString, iniDatei)
WriteINI = (Err = 0)
End Function

Public Function SaveINI(iniDatei As String, sSection As String, sKeyName As String, sNewString As String) As Boolean
On Error Resume Next
' Diese Funktion ist nur zum Weiterleiten gedacht.
' ©reated UNKNOWN
SaveINI = WriteINI(iniDatei, sSection, sKeyName, sNewString)
End Function


Public Function updateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
' Diese Funktion basiert auf einem Beispiel von Microsoft.
' ©reated UNKNOWN
    
    Dim rc As Long                                      ' Rückgabe-Code
    Dim hKey As Long                                    ' Zugriffsnummer für Registrierungsschlüssel
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Sicherheitstyp der Registrierung
    
    lpAttr.nLength = 50                                 ' Sicherheitsattribute auf Standardeinstellungen setzen...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Registrierungsschlüssel erstellen/öffnen...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' //KeyRoot//KeyName erstellen/öffnen
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Fehler behandeln...
    
    '------------------------------------------------------------
    '- Schlüsselwert erstellen/bearbeiten...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' Für RegSetValueEx() wird zur korrekten Ausführung ein Leerzeichen benötigt...
    
    ' Schlüsselwert erstellen/bearbeiten
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Fehler behandeln
    '------------------------------------------------------------
    '- Registrierungsschlüssel schließen...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Schlüssel schließen
    
    updateKey = True                                    ' Erfolgreiche Ausführung zurückgeben
    Exit Function                                       ' Beenden
CreateKeyError:
    updateKey = False                                   ' Fehlerrückgabe-Code festlegen
    rc = RegCloseKey(hKey)                              ' Versuchen, den Schlüssel zu schließen
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
' Diese Funktion basiert auf einem Beispiel von Microsoft.
' ©reated UNKNOWN
    Dim i As Long                                           ' Schleifenzähler
    Dim rc As Long                                          ' Rückgabe-Code
    Dim hKey As Long                                        ' Zugriffsnummer für einen offenen Registrierungsschlüssel
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Datentyp eines Registrierungsschlüssels
    Dim tmpVal As String                                    ' Temporärer Speicher eines Registrierungsschlüsselwertes
    Dim KeyValSize As Long                                  ' Größe einer Registrierungsschlüsselvariablen
    
    ' Registrierungsschlüssel unter dem Stamm {HKEY_LOCAL_MACHINE...} öffnen
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Registrierungsschlüssel öffnen
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
    
    tmpVal = String$(1024, 0)                               ' Platz für Variable reservieren
    KeyValSize = 1024                                       ' Größe der Variable markieren
    
    '------------------------------------------------------------
    ' Registrierungsschlüsselwert abrufen...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Schlüsselwert abrufen/erstellen
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError           ' Fehler behandeln
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Schlüsselwerttyp für Konvertierung bestimmen...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Datentypen durchsuchen...
    Case REG_SZ, REG_EXPAND_SZ                               ' Zeichenfolge für Registrierungsschlüsseldatentyp
        sKeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
    Case REG_DWORD                                           ' Registrierungsschlüsseldatentyp DWORD
        For i = Len(tmpVal) To 1 Step -1                     ' Jedes Bit konvertieren
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))  ' Wert Zeichen für Zeichen erstellen
        Next
        sKeyVal = Format$("&h" + sKeyVal)                    ' DWORD in Zeichenfolge konvertieren
    End Select
    
    GetKeyValue = sKeyVal                                   ' Wert zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
    Exit Function                                           ' Beenden
    
GetKeyError:    ' Bereinigen, nachdem ein Fehler aufgetreten ist...
    GetKeyValue = vbNullString                              ' Rückgabewert auf leere Zeichenfolge setzen
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
End Function

Public Function textMulti(derText As String, anzahl As Long) As String
On Error Resume Next
' Diese ältere Funktion kopiert den String so oft wie gewünscht
' und liefert das Ergebnis zurück.
' ©reated UNKNOWN

Dim i As Long
For i = 1 To anzahl
    textMulti = textMulti & derText
Next
End Function

Public Function getAppPath() As String
On Error Resume Next
' Diese Funktion gibt den Programmpfad mit abschliessendem \ zurück.
' ©reated 21/09/00

Dim melTemp1 As String
melTemp1 = App.Path & "\"
melTemp1 = Replace(melTemp1, "/", "\")
melTemp1 = Replace(melTemp1, "\\", "\")
getAppPath = melTemp1
End Function

Public Function getFileName(FileName As String) As String
On Error Resume Next
' Diese Funktion gibt den Dateinamen mit vorangestelltem Pfad zurück.
' ©reated 21/09/00

Dim melTemp1 As String
melTemp1 = getAppPath & FileName
melTemp1 = Replace(melTemp1, "\\", "\")
getFileName = melTemp1
End Function

Public Function isWin98() As Boolean
On Error Resume Next
' Diese Funktion überprüft, ob das Programm unter Windows 98 läuft.
' Nütztlich z.B. für die AnimateWindow-Funktion.
' ©reated 24/09/00

Dim tempResult As String
tempResult = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\", "Version")
If InStr(tempResult, "98") > 0 Then isWin98 = True
tempResult = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\", "ProductName")
If InStr(tempResult, "98") > 0 Then isWin98 = True
End Function

Public Sub OnTop(hwnd As Long)
On Error Resume Next
' Diese Funktion bring das Fenster mit dem Handle hWnd in den Vordergrund.
' hWnd einer Form kann mit Form1.hWnd festgestellt werden.
' ©reated 24/09/00

If SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags) = True Then
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
End If
End Sub

Public Sub AboutMeLProAlt()
On Error Resume Next
' Diese Funktion zeigt ein Info-Dialogfeld über MeLPro an.
' ©reated 24/09/00

MsgBox "Über MeLPro...:" & vbCr & _
"Programmiert von -=MeL=- aka Pablo Hoch" & vbCr & _
"http://www.melaxis.de, mel@melaxis.de" & vbCr & vbCr & _
"!! This is freeware !!", vbInformation + vbOKOnly, _
"Über MeLPro"

End Sub

Public Function ACos(x As Double)
On Error Resume Next
' Eine Rechenfunktion
' benötigt für den Tube-Effekt

    x = x - 1
    If x < 1 And x > -1 Then
        ACos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    Else
        ACos = 0
    End If
End Function

Public Sub Tube(picBox1 As PictureBox, picBox2 As PictureBox)
On Error Resume Next
' Tube-Effekt für Bilder
' Basiert auf Yanidogs Formel
' picBox1 muss das Bild enthalten, picBox2 sollte leer sein und die
' gleiche Größe wie picBox1 haben, dort entsteht das neue Bild.
' (picBox1 kann also unsichtbar sein)

    Const TubeWidth = 50
    Dim XTube As Long, Offset As Long, XPicture As Long, erg As Double
        
    erg = 3.14159265358979 * 2 / (TubeWidth * 2)
    For Offset = 0 To picBox1.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then picBox2.PaintPicture picBox1.Picture, Offset - TubeWidth, 0, 1, picBox1.ScaleHeight, Offset - TubeWidth, 0, 1, picBox1.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / erg
            If Offset + XPicture < picBox1.ScaleWidth Then
                picBox2.PaintPicture picBox1.Picture, Offset + XTube - TubeWidth, 0, 1, picBox1.ScaleHeight, Offset + XPicture, 0, 1, picBox1.ScaleHeight
            Else
                picBox2.PaintPicture picBox1.Picture, Offset + XTube - TubeWidth, 0, 1, picBox1.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, picBox1.ScaleHeight
            End If
        Next XTube
    Next Offset
End Sub

Public Sub typeWriter(typeText As String, lbl As Object, speed As Long, sound As String)
On Error Resume Next
' Diese Funktion simulliert eine Schreibmaschine.
' Ideeal auf schwarzem Hintergrund mit hellgrauer Schrift.
' Dann noch den Rand der Form weg und Maximiert,
' Cursor Weg (ShowCursor 0) und fertig is der Perfekte
' Effekt ;o)
' ©reated 27/09/00

lbl.Caption = ""
lbl.FontName = "Courier New"
If lbl.FontSize < 15 Then lbl.FontSize = 15
lbl.AutoSize = True
Dim i As Long
For i = 1 To Len(typeText)
    lbl.Caption = lbl.Caption & Mid(typeText, i, 1)
    DoEvents
    If FileExist(App.Path & "\" & sound) Then
        sndPlaySound App.Path & "\" & sound, SND_ASYNC
    End If
    Call Sleep(speed + Round(Rnd * 200))
Next
End Sub

Public Sub melRegFileType(dateiEndung As String, programmName As String, ContentType As String, dateiBeschreibung As String, fileIcon As String, fileCommand As String)
On Error Resume Next
' Diese Funktion registriert einen Dateitypen (z.B. .zip)
' für ein Bestimmtes Programm. Erklärung der Parameter:
' dateiEndung = Datei-Endung (Bsp: .zip)
' programmName = Kurzer Name (Bsp: WinZIP)
' contentType = Inhaltsbeschreibung (Bsp: text/html, text/plain, application/x-yourfile)
' dateiBeschreibung = Kurze Erklärung, wird im Explorer angezeigt (Bsp: Zip-Datei)
' fileIcon = Pfad zum Icon (Bsp: c:\programm.exe,0 oder c:\programm.ico)
' fileCommand = Pfad zum Programm (Bsp: c:\programm.exe %1 (%1 = Dateiname))
' ©reated 28/09/00

updateKey HKEY_CLASSES_ROOT, dateiEndung, "Save", GetKeyValue(HKEY_CLASSES_ROOT, dateiEndung, "")
updateKey HKEY_CLASSES_ROOT, dateiEndung, "", programmName
updateKey HKEY_CLASSES_ROOT, dateiEndung, "Content Type", ContentType
updateKey HKEY_CLASSES_ROOT, dateiEndung, "Info", "Dateityp wurde mit Hilfe von MeLPro von Pablo Hoch aka -=MeL=- registriert :o)"
updateKey HKEY_CLASSES_ROOT, programmName, "", dateiBeschreibung
updateKey HKEY_CLASSES_ROOT, programmName & "\DefaultIcon", "", fileIcon
updateKey HKEY_CLASSES_ROOT, programmName & "\shell", "", ""
updateKey HKEY_CLASSES_ROOT, programmName & "\shell\open", "", programmName
updateKey HKEY_CLASSES_ROOT, programmName & "\shell\open\command", "", fileCommand

End Sub

Public Function getFileTypeApp(dateiEndung As String) As melAppInfo
On Error Resume Next
' Diese Funktion liefert Infos über das Programm zurück.
' ©reated 28/09/00

Dim temp1 As String
getFileTypeApp.dateiEndung = dateiEndung
getFileTypeApp.programmName = GetKeyValue(HKEY_CLASSES_ROOT, dateiEndung, "")
getFileTypeApp.prevApp = GetKeyValue(HKEY_CLASSES_ROOT, dateiEndung, "Save")
getFileTypeApp.ContentType = GetKeyValue(HKEY_CLASSES_ROOT, dateiEndung, "Content Type")
temp1 = getFileTypeApp.programmName
getFileTypeApp.dateiBeschreibung = GetKeyValue(HKEY_CLASSES_ROOT, temp1, "")
getFileTypeApp.fileIcon = GetKeyValue(HKEY_CLASSES_ROOT, temp1 & "\DefaultIcon", "")
getFileTypeApp.fileCommand = GetKeyValue(HKEY_CLASSES_ROOT, temp1 & "\shell\open\command", "")
End Function

Public Sub setUninstallInfo(programName As String, uninstallString As String)
On Error Resume Next
' Diese Funktion trägt ein Programm in der Liste der installierten Programme
' (Systemsteuerung -> Software)
' ©reated 29/09/00

Dim uniStr As String
uniStr = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & programName
updateKey HKEY_LOCAL_MACHINE, uniStr, "DisplayName", programName
updateKey HKEY_LOCAL_MACHINE, uniStr, "UninstallString", uninstallString
End Sub

Public Sub setLongUninstallInfo(programName As String, uninstallString As String, programVersion As String, contactInfo As String, programIcon As String, helpPhone As String, helpURL As String, Firma As String, readmeFile As String, Kommentare As String, infoURL As String)
On Error Resume Next
' Diese Funktion trägt ein Programm in der Liste der installierten Programme
' (Systemsteuerung -> Software)
' Dabei werden möglichst viele Parameter verwendet.
' ©reated 29/09/00

Dim uniStr As String
uniStr = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & programName
updateKey HKEY_LOCAL_MACHINE, uniStr, "DisplayName", programName
updateKey HKEY_LOCAL_MACHINE, uniStr, "UninstallString", uninstallString
updateKey HKEY_LOCAL_MACHINE, uniStr, "DisplayVersion", programVersion
updateKey HKEY_LOCAL_MACHINE, uniStr, "Contact", contactInfo
updateKey HKEY_LOCAL_MACHINE, uniStr, "DisplayIcon", programIcon
updateKey HKEY_LOCAL_MACHINE, uniStr, "HelpPhone", helpPhone
updateKey HKEY_LOCAL_MACHINE, uniStr, "HelpLink", helpURL
updateKey HKEY_LOCAL_MACHINE, uniStr, "Publisher", Firma
updateKey HKEY_LOCAL_MACHINE, uniStr, "Readme", readmeFile
updateKey HKEY_LOCAL_MACHINE, uniStr, "Comments", Kommentare
updateKey HKEY_LOCAL_MACHINE, uniStr, "URLInfoAbout", infoURL
End Sub

Public Sub delRegValue(hKey As Long, valueString As String)
On Error Resume Next
' Diese Funktion löscht einen Wert aus der Registry.
' ©reated 29/09/00

RegDeleteValue hKey, valueString
End Sub

Public Sub delRegKey(hKey As Long, keyString As String)
On Error Resume Next
' Diese Funktion löscht einen Schlüssel aus der Registry.
' ©reated 29/09/00

RegDeleteKey hKey, keyString
End Sub

Public Function reverseString(oriString As String) As String
On Error Resume Next
' Diese Funktion dreht einen String um.
' ©reated 29/09/00

Dim i As Long
For i = 0 To Len(oriString)
    reverseString = reverseString + Mid(oriString, Len(oriString) - i, 1)
Next
End Function

Public Function reverseCode(oriString As String) As String
On Error Resume Next
' Diese Funktion dreht einen String um und kodiert ihn dabei mit M175C
' ©reated 29/09/00

Dim i As Long
For i = 0 To Len(oriString)
    reverseCode = reverseCode + Chr(Asc(Mid(oriString, Len(oriString) - i, 1)) + 1)
Next
End Function

Public Function reverseDeCode(oriString As String) As String
On Error Resume Next
' Diese Funktion dreht einen String um und dekodiert ihn dabei mit M175C
' ©reated 29/09/00

Dim i As Long
For i = 0 To Len(oriString)
    reverseDeCode = reverseDeCode + Chr(Asc(Mid(oriString, Len(oriString) - i, 1)) - 1)
Next
End Function

Public Sub DragForm(hwnd As Long)
On Error Resume Next
' Diese Funktion ist zum Verschieben von Forms
' ohne Rand gedacht. Bei MouseMove und Button = 1
' DragDorm(Me.hWnd) ausführen.
' ©reated 29/09/00

Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Sub killWindow(hwnd As Long)
On Error Resume Next
' Diese Funktion schliesst ein Fenster, egal welches Programm,
' durch Senden der Destroy-Nachricht.
' ©reated 30/09/00

SendMessage hwnd, WM_DESTROY, 0, 0
SendMessage hwnd, WM_NCDESTROY, 0, 0
End Sub

Public Function StartMe(Frm As Form, ToOpen As String)
On Error Resume Next
' Diese Funktion startet ein Programm oder ruft eine
' Website auf.
' ©reated 01/10/00

ShellExecute Frm.hwnd, "Open", ToOpen, &O0, &O0, SW_NORMAL
End Function

Sub MakeRoundWindow(theform As Object)
On Error Resume Next
' Diese Funktion erstellt ein rundes Fenster.
' BorderStyle muss 0 sein
' ©reated 01/10/00

    Dim Do_Left
    Dim Do_Top
    Dim Do_X2
    Dim Do_Y2
    Dim Do_hWnd
    Dim Do_Round_RGN&

    Do_Left = 1
    Do_Top = 1
    Do_X2 = theform.Width / Screen.TwipsPerPixelX - 1
    Do_Y2 = theform.Height / Screen.TwipsPerPixelY - 1
    Do_hWnd = theform.hwnd

    Do_Round_RGN = CreateEllipticRgn(Do_Left, Do_Top, Do_X2, Do_Y2)
    SetWindowRgn Do_hWnd, Do_Round_RGN, True
End Sub

Public Function Draw_Gradient(Frm As Form, Color1 As Long, Color2 As Long)
On Error Resume Next
' Diese Funktion erstellt einen Verlauf.
' ©reated 01/10/00

Dim r1, g1, b1, r2, g2, b2, boxStep, posY, i As Integer
Dim redStep, greenStep, BlueStep As Integer
' separate color1 to red,green and blue
r1 = Color1 Mod &H100
g1 = (Color1 \ &H100) Mod &H100
b1 = (Color1 \ &H10000) Mod &H100
' separate color2 to red,green and blue
r2 = Color2 Mod &H100
g2 = (Color2 \ &H100) Mod &H100
b2 = (Color2 \ &H10000) Mod &H100
' calculate box height
boxStep = Frm.ScaleHeight / 63
posY = 0
If g1 > g2 Then
greenStep = 0
ElseIf g2 > g1 Then
greenStep = 1
Else
greenStep = 2
End If
If r1 > r2 Then
redStep = 0
ElseIf r2 > r1 Then
redStep = 1
Else
redStep = 2
End If
If b1 > b2 Then
BlueStep = 0
ElseIf b2 > b1 Then
BlueStep = 1
Else
BlueStep = 2
End If
For i = 1 To 63
Frm.Line (0, posY)-(Frm.ScaleWidth, posY + boxStep), RGB(r1, g1, b1), BF
If redStep = 1 Then
r1 = r1 + 4
If r1 > r2 Then
r1 = r2
End If
ElseIf redStep = 0 Then
r1 = r1 - 4
If r1 < r2 Then
r1 = r2
End If
End If
If greenStep = 1 Then
g1 = g1 + 4
If g1 > g2 Then
g1 = g2
End If
ElseIf greenStep = 0 Then
g1 = g1 - 4
If g1 < g2 Then
g1 = g2
End If
End If
If BlueStep = 1 Then
b1 = b1 + 4
If b1 > b2 Then
b1 = b2
End If
ElseIf BlueStep = 0 Then
b1 = b1 - 4
If b1 < b2 Then
b1 = b2
End If
End If
posY = posY + boxStep
Next i
End Function

Public Function CaptureDesktop(oPictureBox As Object) As Boolean

'This code captures a current image of the desktop and
'saves it to a picture box, image box, or other control that
'supports a picture property

'Pass in the image control as the oPictureBox parameter

'Example CaptureDesktop Picture1

    On Error GoTo Errhandler
    Clipboard.Clear
    keybd_event VK_MENU, 0, 0, 0    ' Plant "Alt" key
    DoEvents


    keybd_event VK_SNAPSHOT, 1, 0, 0
    DoEvents
    ' Release "Alt" key
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
   
    DoEvents
      ' (Image is now in clipboard) put into picture box
    Set oPictureBox.Picture = Clipboard.GetData(0)
    Clipboard.Clear
   
    CaptureDesktop = True
Errhandler:
    End Function

Sub CenterOBJ(Frm As Object)
On Error Resume Next
Dim x%, y%
x = Screen.Width / 2 - Frm.Width / 2
y = Screen.Height / 2 - Frm.Height / 2
Frm.Move x, y
End Sub

Public Function IconFromBinary(FileName As String, myFrm As Form, PictureControl As Object) As Boolean
On Error Resume Next

'*******************************************************
'PURPOSE: Extracts icon from an .exe or .dll file and
'displays it in a PictureBox

'PARAMETERS:

'FileName: Full Path of Binary file
'(.dll or .exe containing the Icon of interest

'ImageControl:  PictureBox to Display Icon In

'RETURNS: True if successful, false otherwise

'EXAMPLE: IconFromBinary ("C:\MyApp.exe", Picture1)

'NOTES:  THIS FUNCTION MUST BE PLACED IN
'A FORM MODULE. IF YOU WANT TO USE IT IN A
'CLASS OR .BAS FILE, YOU MUST ADD A FORM OR
'FORM'S HWND AS A PARAMETER
'*********************************************

On Error GoTo errorhandler:
Dim lret As Long
Dim hIcon As Long
Dim lHdc As Long
Dim sFile As String


'If Dir(FileName) = "" Then Exit Function

lHdc = PictureControl.hDC
If lHdc = 0 Then Exit Function


myFrm.AutoRedraw = True
PictureControl.AutoRedraw = True

sFile = FileName & Chr(0)

hIcon = ExtractIcon(myFrm.hwnd, sFile, 0)

lret = DrawIcon(lHdc, 0, 0, hIcon)
If lret <> 0 Then
    PictureControl.Refresh
    DestroyIcon hIcon
    IconFromBinary = Err.LastDllError = 0
End If

errorhandler:
End Function


Public Function gimmePath(param) As String
On Error Resume Next
Dim mt1, mt2
mt1 = Split(param, "\")
For mt2 = 0 To UBound(mt1) - 1
    gimmePath = gimmePath & mt1(mt2) & "\"
Next
End Function

Public Function gimmeExtension(FileName As String) As String
On Error Resume Next

FileName = cutPathName(FileName)
Dim t1, t2
t1 = Split(FileName, ".")
t2 = t1(UBound(t1))

gimmeExtension = t2

End Function

Public Function gimmeIco(FileName As String, picBox As Object, Frm As Form) As Boolean
On Error GoTo Err
Dim tI1, ti2, tI3
tI1 = gimmeExtension(FileName)
Dim mfat As melAppInfo
mfat = getFileTypeApp("." & tI1)
ti2 = mfat.fileIcon


picBox.Picture = LoadPicture(ti2)
gimmeIco = True
Exit Function

Err:
gimmeIco = IconFromBinary(CStr(ti2), Frm, picBox)

End Function

Public Function fileIsNull(FileName) As Boolean
On Error Resume Next
Dim mFl As Long
mFl = 0
mFl = FileLen(FileName)
If mFl = 0 Then
    fileIsNull = True
Else
    fileIsNull = False
End If
End Function

Public Function FromSz(szStr As String) As String
   If InStr(szStr, vbNullChar) Then
      FromSz = Left(szStr, InStr(szStr, vbNullChar) - 1)
   Else
      FromSz = szStr
   End If
End Function

Public Function ShellGetText(Program As String) As String
       Dim sTempFile As String
       Dim hFile As Long
       Dim pid As Long
       Dim hProcess As Long
       Dim bResult As Boolean

       sTempFile = Space(1024)
       GetTempFileName Environ("TEMP"), "OUT", 0, sTempFile
       sTempFile = FromSz(sTempFile)

       pid = Shell(Environ("COMSPEC") & " /C " & Program & ">" & sTempFile, vbHidden)
       hProcess = OpenProcess(SYNCHRONIZE, True, pid)

       Do Until (hProcess = 0) Or WaitForSingleObject(hProcess, 60000)
          GoTo CloseHandles
       Loop

CloseHandles:
       hFile = FreeFile
       Open sTempFile For Binary As #hFile
       ShellGetText = Input$(LOF(hFile), hFile)
       Close #hFile

       CloseHandle hProcess
       Kill sTempFile
End Function

Public Function getRealNick(ByVal nickName As String) As String
On Error Resume Next
' Diese Funktion sucht aus einem Nick den echten Namen
' ©reated 11/10/00

nickName = Replace(nickName, "-", "")
nickName = Replace(nickName, "~", "")
nickName = Replace(nickName, "=", "")
nickName = Replace(nickName, "<", "")
nickName = Replace(nickName, ">", "")
nickName = Replace(nickName, "|", "")
nickName = Replace(nickName, "[", "")
nickName = Replace(nickName, "]", "")
nickName = Replace(nickName, "_", "")
nickName = Replace(nickName, "^", "")
nickName = Replace(nickName, "°", "")
nickName = Replace(nickName, "*", "")
nickName = Replace(nickName, "+", "")
nickName = Replace(nickName, ".", "")
nickName = Replace(nickName, ",", "")
nickName = Replace(nickName, "&", "")
nickName = Replace(nickName, "@", "")
nickName = Replace(nickName, "#", "")
nickName = Replace(nickName, "'", "")
nickName = Replace(nickName, """", "")
nickName = Replace(nickName, "?", "")
nickName = Replace(nickName, "$", "")
nickName = Replace(nickName, "§", "")
nickName = Replace(nickName, "!", "")
nickName = Replace(nickName, "%", "")
nickName = Replace(nickName, "/", "")
nickName = Replace(nickName, "\", "")
nickName = Replace(nickName, "(", "")
nickName = Replace(nickName, ")", "")
nickName = Replace(nickName, "´", "")
nickName = Replace(nickName, "`", "")
nickName = Replace(nickName, "·", "")
nickName = Replace(nickName, "©", "c")
nickName = Replace(nickName, "®", "r")
nickName = Replace(nickName, "™", "tm")
nickName = Replace(nickName, "¿", "")
nickName = Replace(nickName, ":", "")
nickName = Replace(nickName, ";", "")
Dim Num As Long
For Num = 0 To 9
    nickName = Replace(nickName, Num, "")
Next
nickName = Trim(LCase(nickName))
getRealNick = nickName

End Function

Public Function RexDeCode(rexT) As String
On Error Resume Next
Dim ws1 As String, ws As String, ws0 As Long, b1 As String, ac As Long
ws1 = ""
ws = rexT
For ws0 = 1 To Len(ws)
    b1 = Mid(ws, ws0, 1)
    ac = Asc(b1)
    If ac - 12 < 0 Then ws1 = ws1 & Chr(255 - (12 - ac)) Else ws1 = ws1 & Chr(ac - 12)
Next
RexDeCode = ws1
End Function

Public Function RexCode(rexT) As String
On Error Resume Next
Dim ws1 As String, ws As String, ws0 As Long, b1 As String, ac As Long
ws1 = ""
ws = rexT
For ws0 = 1 To Len(ws)
    b1 = Mid(ws, ws0, 1)
    ac = Asc(b1)
    If ac + 12 < 255 Then ws1 = ws1 & Chr(ac + 12)
Next
RexCode = ws1
End Function

Public Function Bin2Char(Bin As String) As String
On Error Resume Next
    Dim i As Integer
    Dim CharByte As Byte
    CharByte = 0
    For i = 1 To 7
        If Mid(Bin, i, 1) = "1" Then
            CharByte = CharByte + (2 ^ (8 - i))
        End If
    Next
    If Mid(Bin, 8, 1) = "1" Then
        CharByte = CharByte + 1
    End If
    Bin2Char = Chr(CharByte)
End Function

Public Function Char2Bin(Char As String) As String
On Error Resume Next
    Dim CharByte As Byte
    CharByte = Asc(Char)
    Dim BinOut As String
    BinOut = ""
    Dim i As Integer
    For i = 7 To 1 Step -1
        If (CharByte / (2 ^ i)) >= 1 Then
            CharByte = CharByte - (2 ^ i)
            BinOut = BinOut + "1"
        Else
            BinOut = BinOut + "0"
        End If
    Next
    If CharByte = 1 Then
        BinOut = BinOut + "1"
    Else
        BinOut = BinOut + "0"
    End If
    Char2Bin = BinOut
End Function

Public Function UpdateProgress(pb As Control, ByVal Percent)
'Replacement for progress bar..looks nicer also
Dim Num$ 'use percent
If Not pb.AutoRedraw Then 'picture in memory ?
pb.AutoRedraw = -1 'no, make one
End If
pb.Cls 'clear picture in memory
pb.ScaleWidth = 100 'new sclaemodus
pb.DrawMode = 10 'not XOR Pen Modus
Num$ = Format$(Percent, "###") + "%"
pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
pb.Print Num$ 'print percent
pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
pb.Refresh 'show differents
End Function

Public Sub bDown(obj As Object)
On Error Resume Next
obj.Left = obj.Left + 10
obj.Top = obj.Top + 10
End Sub

Public Sub bUp(obj As Object)
On Error Resume Next
obj.Left = obj.Left - 10
obj.Top = obj.Top - 10
End Sub

Private Sub Assemble128()
   x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
   code128
   inter = res
   
   x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
   code128
   inter = inter Xor res
   
   x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
   code128
   inter = inter Xor res
   
   x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
   code128
   inter = inter Xor res
   
   x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
   code128
   inter = inter Xor res
   
   x1a0(5) = x1a0(4) Xor ((cle(11) * 256) + cle(12))
   code128
   inter = inter Xor res
   
   x1a0(6) = x1a0(5) Xor ((cle(13) * 256) + cle(14))
   code128
   inter = inter Xor res
   
   x1a0(7) = x1a0(6) Xor ((cle(15) * 256) + cle(16))
   code128
   inter = inter Xor res
   
   i = 0
End Sub

Private Sub code128()
   dx = (x1a2 + i) Mod 65536
   ax = x1a0(i)
   cX = &H15A
   bx = &H4E35
   
   tmp = ax
   ax = si
   si = tmp
   
   tmp = ax
   ax = dx
   dx = tmp
   
   If (ax <> 0) Then
      ax = (ax * bx) Mod 65536
   End If
   
   tmp = ax
   ax = cX
   cX = tmp
   
   If (ax <> 0) Then
      ax = (ax * si) Mod 65536
      cX = (ax + cX) Mod 65536
   End If
   
   tmp = ax
   ax = si
   si = tmp
   ax = (ax * bx) Mod 65536
   dx = (cX + dx) Mod 65536
   
   ax = ax + 1
   
   x1a2 = dx
   x1a0(i) = ax
   
   res = ax Xor dx
   i = i + 1
End Sub

Private Function Encrypt128(ByVal Plaintext As String, ByVal Key As String) As String
   Dim sData As String
   
   si = 0
   x1a2 = 0
   i = 0
   
   For fois = 1 To 16
      cle(fois) = 0
   Next fois
   
   champ1 = Key
   lngchamp1 = Len(champ1)
   
   For fois = 1 To lngchamp1
      cle(fois) = Asc(Mid(champ1, fois, 1))
   Next fois
   
   champ1 = Plaintext
   lngchamp1 = Len(champ1)
   For fois = 1 To lngchamp1
      c = Asc(Mid(champ1, fois, 1))
      
      Assemble128
      
      cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
      cfd = inter Mod 256
      
      For compte = 1 To 16
         cle(compte) = cle(compte) Xor c
      Next compte
      
      c = c Xor (cfc Xor cfd)
      
      d = (((c / 16) * 16) - (c Mod 16)) / 16
      e = c Mod 16
      
      sData = sData & Chr$(&H61 + d)
      sData = sData & Chr$(&H61 + e)
   Next fois

   Encrypt128 = sData
End Function

Private Function Decrypt128(ByVal Text As String, ByVal Key As String) As String
   Dim sData As String
   si = 0
   x1a2 = 0
   i = 0
   
   For fois = 1 To 16
      cle(fois) = 0
   Next fois
   
   champ1 = Key
   lngchamp1 = Len(champ1)
   
   For fois = 1 To lngchamp1
   cle(fois) = Asc(Mid(champ1, fois, 1))
   Next fois
   
   champ1 = Text
   lngchamp1 = Len(champ1)
   
   For fois = 1 To lngchamp1
      d = Asc(Mid(champ1, fois, 1))
      If (d - &H61) >= 0 Then
          d = d - &H61
          If (d >= 0) And (d <= 15) Then
              d = d * 16
          End If
      End If
      If (fois <> lngchamp1) Then
          fois = fois + 1
      End If
      e = Asc(Mid(champ1, fois, 1))
      If (e - &H61) >= 0 Then
          e = e - &H61
          If (e >= 0) And (e <= 15) Then
              c = d + e
          End If
      End If
      
      Assemble128
      
      cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
      cfd = inter Mod 256
      
      c = c Xor (cfc Xor cfd)
      
      For compte = 1 To 16
          cle(compte) = cle(compte) Xor c
      Next compte
      
      sData = sData & Chr$(c)
   Next fois
   Decrypt128 = sData
End Function

Private Sub Assemble80()
   x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
   code80
   inter = res
   
   x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
   code80
   inter = inter Xor res
       
   x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
   code80
   inter = inter Xor res
   
   x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
   code80
   inter = inter Xor res
   
   x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
   code80
   inter = inter Xor res
   
   i = 0
End Sub

Private Sub code80()
   dx = (x1a2 + i) Mod 65536
   ax = x1a0(i)
   cX = &H15A
   bx = &H4E35
   
   tmp = ax
   ax = si
   si = tmp
   
   tmp = ax
   ax = dx
   dx = tmp
   
   If (ax <> 0) Then
       ax = (ax * bx) Mod 65536
   End If
   
   tmp = ax
   ax = cX
   cX = tmp
   
   If (ax <> 0) Then
       ax = (ax * si) Mod 65536
       cX = (ax + cX) Mod 65536
   End If
   
   tmp = ax
   ax = si
   si = tmp
   ax = (ax * bx) Mod 65536
   dx = (cX + dx) Mod 65536
   
   ax = ax + 1
   
   x1a2 = dx
   x1a0(i) = ax
   
   res = ax Xor dx
   i = i + 1
End Sub

Private Function Encrypt80(ByVal Plaintext, ByRef Key) As String
   Dim Crooked As String
   si = 0
   x1a2 = 0
   i = 0
   
   For fois = 1 To 10
      cle(fois) = 0
   Next fois
   
   champ1 = Key
   lngchamp1 = Len(champ1)
   
   For fois = 1 To lngchamp1
      cle(fois) = Asc(Mid(champ1, fois, 1))
   Next fois
   
   champ1 = Plaintext
   lngchamp1 = Len(champ1)
   For fois = 1 To lngchamp1
      c = Asc(Mid(champ1, fois, 1))
      
      Assemble80
      
      cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
      cfd = inter Mod 256
      
      For compte = 1 To 10
          cle(compte) = cle(compte) Xor c
      Next compte
      
      c = c Xor (cfc Xor cfd)
      
      d = (((c / 16) * 16) - (c Mod 16)) / 16
      e = c Mod 16
      
      Crooked = Crooked + Chr$(&H61 + d)
      Crooked = Crooked + Chr$(&H61 + e)
   Next fois
   Encrypt80 = Crooked
End Function

Private Function Decrypt80(ByVal EncryptedText, ByRef Key) As String
   Dim Plaintext As String
   si = 0
   x1a2 = 0
   i = 0
   
   For fois = 1 To 10
       cle(fois) = 0
   Next fois
   
   champ1 = Key
   lngchamp1 = Len(champ1)
   
   For fois = 1 To lngchamp1
       cle(fois) = Asc(Mid(champ1, fois, 1))
   Next fois
   
   champ1 = EncryptedText
   lngchamp1 = Len(champ1)
   
   For fois = 1 To lngchamp1
      d = Asc(Mid(champ1, fois, 1))
      If (d - &H61) >= 0 Then
         d = d - &H61
         If (d >= 0) And (d <= 15) Then
            d = d * 16
         End If
      End If
      If (fois <> lngchamp1) Then
         fois = fois + 1
      End If
      e = Asc(Mid(champ1, fois, 1))
      If (e - &H61) >= 0 Then
         e = e - &H61
         If (e >= 0) And (e <= 15) Then
            c = d + e
         End If
      End If
      
      Assemble80
      
      cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
      cfd = inter Mod 256
      
      c = c Xor (cfc Xor cfd)
      
      For compte = 1 To 10
          cle(compte) = cle(compte) Xor c
      Next compte
      
      Plaintext = Plaintext + Chr$(c)
   Next fois
   Decrypt80 = Plaintext
End Function

Public Function Encrypt(TextIn As String, Key As String) As String
   Dim sKey As String
   
   On Error GoTo TooLong:
   If Len(Key) < 11 Then
      sKey = Key & Space$(10 - Len(Key))
      Encrypt = Encrypt80(TextIn, sKey)
   Else
      sKey = Key & Space$(16 - Len(Key))
      Encrypt = Encrypt128(TextIn, sKey)
   End If
   Exit Function
TooLong:
   MsgBox "Das Passwort darf maximal nur 16 Zeichen lang sein."
End Function

Public Function Decrypt(TextIn As String, Key As String) As String
   Dim sKey As String
   
   On Error GoTo TooLong:
   If Len(Key) < 11 Then
      sKey = Key & Space$(10 - Len(Key))
      Decrypt = Decrypt80(TextIn, sKey)
   Else
      sKey = Key & Space$(16 - Len(Key))
      Decrypt = Decrypt128(TextIn, sKey)
   End If
   Exit Function
TooLong:
   MsgBox "Das Passwort darf maximal nur 16 Zeichen lang sein."
End Function

Public Function GetHost(ByVal HOST$) As Long
    Dim ListAddress As Long
    Dim ListAddr As Long
    Dim LH&, phe&
    Dim START As Boolean
    Dim heDestHost As HOSTENT
    Dim addrList&, repIP&


  
   START = SocketsInitialize
    If START = False Then GetHost = 0: MsgBox ("Fehler bei der SocketInitialisierung!"): Exit Function

   LH = inet_addr(HOST$)
'   If LH = INADDR_NONE Then
'    repIP = htonl(LH)
'   Else
  repIP = LH
      If LH = INADDR_NONE Then
      phe = gethostbyname(HOST$)

        If phe <> 0 Then

            CopyMemory heDestHost, ByVal phe, hostent_size

            CopyMemory addrList, ByVal heDestHost.hAddrList, 4

            CopyMemory repIP, ByVal addrList, heDestHost.hLen
            
        Else
         'MsgBox ("GetHostByName lieferte ungültiges Ergebnis!")
         GetHost = INADDR_NONE
         Exit Function
        End If
   
'
'
'
'
'     a = GetHostByName(Host)
'     ' Copy Winsock structure to the VisualBasic structure
'
'    CopyMemory hostent_async.h_name, ByVal PointerToPointer, Len(hostent_async)
'
'    ListAddress = hostent_async.h_addr_list        ' Get the ListAddress of the Address List
'
'    CopyMemory ListAddr, ByVal ListAddress, 4      ' Copy Winsock structure to the VisualBasic structure
'    CopyMemory IPLong, ByVal ListAddr, 4           ' Get the first list entry from the Address List
    'CopyMemory GIP, ByVal ListAddr, 4

   ' RecWin.Text = Trim$(CStr(Asc(IPLong.Byte4)) + "." + CStr(Asc(IPLong.Byte3)) _
    '    + "." + CStr(Asc(IPLong.Byte2)) + "." + CStr(Asc(IPLong.Byte1)))
   
   End If
   'Form1.Text4.Text = CStr(repIP)
   GetHost = repIP
End Function


Public Function StatusText(status As Long) As String

   Dim msg As String

   Select Case status
      Case IP_SUCCESS:               msg = "erfolgreich"
      Case IP_BUF_TOO_SMALL:         msg = "buffer zu klein"
      Case IP_DEST_NET_UNREACHABLE:  msg = "netz nicht erreichbar"
      Case IP_DEST_HOST_UNREACHABLE: msg = "host nicht erreichbar"
      Case IP_DEST_PROT_UNREACHABLE: msg = "protokoll nicht erreichbar"
      Case IP_DEST_PORT_UNREACHABLE: msg = "port nicht erreichbar"
      Case IP_NO_RESOURCES:          msg = "keine ressourcen"
      Case IP_BAD_OPTION:            msg = "bad option"
      Case IP_HW_ERROR:              msg = "hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "req timed out"
      Case IP_BAD_REQ:               msg = "bad req"
      Case IP_BAD_ROUTE:             msg = "bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "param_problem"
      Case IP_SOURCE_QUENCH:         msg = "source quench"
      Case IP_OPTION_TOO_BIG:        msg = "option too_big"
      Case IP_BAD_DESTINATION:       msg = "bad destination"
      Case IP_ADDR_DELETED:          msg = "addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "spec mtu change"
      Case IP_MTU_CHANGE:            msg = "mtu_change"
      Case IP_UNLOAD:                msg = "unload"
      Case IP_ADDR_ADDED:            msg = "addr added"
      Case IP_GENERAL_FAILURE:       msg = "general failure"
      Case IP_PENDING:               msg = "pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   StatusText = msg
   
End Function

Public Function GetStatusCode(status As Long) As String

   Dim msg As String

   Select Case status
      Case IP_SUCCESS:               msg = "success"
      Case IP_BUF_TOO_SMALL:         msg = "buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "no resources"
      Case IP_BAD_OPTION:            msg = "bad option"
      Case IP_HW_ERROR:              msg = "hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "req timed out"
      Case IP_BAD_REQ:               msg = "bad req"
      Case IP_BAD_ROUTE:             msg = "bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "param_problem"
      Case IP_SOURCE_QUENCH:         msg = "source quench"
      Case IP_OPTION_TOO_BIG:        msg = "option too_big"
      Case IP_BAD_DESTINATION:       msg = "bad destination"
      Case IP_ADDR_DELETED:          msg = "addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "spec mtu change"
      Case IP_MTU_CHANGE:            msg = "mtu_change"
      Case IP_UNLOAD:                msg = "unload"
      Case IP_ADDR_ADDED:            msg = "addr added"
      Case IP_GENERAL_FAILURE:       msg = "general failure"
      Case IP_PENDING:               msg = "pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   
End Function


'Public Function HiByte(ByVal wParam As Integer)
'
'    HiByte = wParam \ &H100 And &HFF&
'
'End Function


Public Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function


Public Function Ping(szAddress As String, ECHO As ICMP_ECHO_REPLY, Data As String) As Long
On Error Resume Next
If PING_TIMEOUT = 0 Then PING_TIMEOUT = 5000
   Dim hPort As Long
   Dim dwAddress As Long
   Dim sDataToSend As String
   Dim iOpt As Long
   Dim a
   
   sDataToSend = Trim$(Data)
   dwAddress = GetHost(szAddress)
   'dwAddress = AddressStringToLong(szAddress)
   
   hPort = IcmpCreateFile()
   If IcmpSendEcho(hPort, _
                   dwAddress, _
                   sDataToSend, _
                   Len(sDataToSend), _
                   0, _
                   ECHO, _
                   Len(ECHO), _
                   PING_TIMEOUT) Then
   
        '
         Ping = ECHO.RoundTripTime
   Else: Ping = ECHO.status * -1
   End If
   Call IcmpCloseHandle(hPort)
   a = SocketsCleanup
End Function
   

Function AddressStringToLong(ByVal tmp As String) As Long

   Dim i As Integer
   Dim parts(1 To 4) As String
   
   i = 0
   
  '
   While InStr(tmp, ".") > 0
      i = i + 1
      parts(i) = Mid(tmp, 1, InStr(tmp, ".") - 1)
      tmp = Mid(tmp, InStr(tmp, ".") + 1)
   Wend
   
   i = i + 1
   parts(i) = tmp
   
   If i <> 4 Then
      AddressStringToLong = 0
      Exit Function
   End If
   
  '
   AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & _
                         Right("00" & Hex(parts(3)), 2) & _
                         Right("00" & Hex(parts(2)), 2) & _
                         Right("00" & Hex(parts(1)), 2))
   
End Function


Public Function SocketsCleanup() As Boolean

    Dim x As Long
    
    x = WSACleanup()
    
    If x <> 0 Then
        MsgBox "Windows Sockets error " & Trim$(str$(x)) & _
               " occurred in Cleanup.", vbExclamation
        SocketsCleanup = False
    Else
        SocketsCleanup = True
    End If
    
End Function


Public Function SocketsInitialize() As Boolean

    Dim WSAD As WSADATA
    Dim x As Integer
    Dim szLoByte As String, szHiByte As String, szBuf As String
    
    x = WSAStartup(WS_VERSION_REQD, WSAD)
    
    If x <> 0 Then
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
        SocketsInitialize = False
        Exit Function
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
       (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
        HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        
        szHiByte = Trim$(str$(HiByte(WSAD.wVersion)))
        szLoByte = Trim$(str$(LoByte(WSAD.wVersion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        SocketsInitialize = False
        Exit Function
        
    End If
    
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        szBuf = "This application requires a minimum of " & _
                 Trim$(str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox szBuf, vbExclamation
        SocketsInitialize = False
        Exit Function
    End If
    
    SocketsInitialize = True
        
End Function

'-----------------------------------------------------------
' FUNCTION: ResolveResString
' Liest die Ressource und ersetzt vorgegebene
' Makros mit vorgegebenen Werten.
'
' Beispiel: gegeben sei die Ressourcenummer 14:
'    "'|1' auf Laufwerk |2 konnte nicht gelesen werden"
'   Der Aufruf
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   hätte die Rückgabe der folgenden Zeichenfolge zur Folge
'     "'TXTFILE.TXT' auf Laufwerk A: konnte nicht gelesen werden"
'
' Eingabe: [resID] - Ressourcennummer
'     [varReplacements] - Makro/Ersetzungswert-Paar
'-----------------------------------------------------------
'
Public Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    strResString = LoadResString(resID)
    
    ' Für jedes übergebene Makro/Wert-Paar...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Ersetzen aller vorkommenden strMacro durch strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function

Function ConvertString(sInput As String) As String
On Error Resume Next
' Diese Funktion verschlüsselt einen String nach dem URLENCODE-Verfahren
' Benötigt für das HTTP-Protkoll (GET + POST)
    ConvertString = sInput
    Dim cnt As Long
    ChangeString ConvertString, "%", "%" + Hex(Asc("%"))
    ChangeString ConvertString, "+", "%" + Hex(Asc("+"))
    ChangeString ConvertString, " ", "+"
    For cnt = 0 To 255
        If Not (cnt > 64 And cnt < 91) And Not (cnt > 96 And cnt < 123) And Not (cnt > 47 And cnt < 58) And cnt <> Asc("%") And cnt <> 32 And cnt <> 43 Then
            ChangeString ConvertString, Chr$(cnt), "%" + DoubleHex(cnt), vbBinaryCompare
        End If
    Next cnt
End Function

Sub ChangeString(sInput As String, ByVal WhatToReplace As String, ByVal ReplaceWith As String, Optional CM As VbCompareMethod = vbTextCompare)
On Error Resume Next
    Dim Ret As Long
    Ret = -Len(ReplaceWith) + 1
    Do
        Ret = InStr(Ret + Len(ReplaceWith), sInput, WhatToReplace, CM)
        If Ret = 0 Then Exit Do
        sInput = Left$(sInput, Ret - 1) + ReplaceWith + Right$(sInput, Len(sInput) - Ret - Len(WhatToReplace) + 1)
    Loop
End Sub

Function DoubleHex(lNumber As Long) As String
On Error Resume Next
    DoubleHex = Hex$(lNumber)
    If Len(DoubleHex) < 2 Then DoubleHex = "0" + DoubleHex
End Function

Public Function SHR(bincode As String) As String
On Error Resume Next
' Diese Funktion führt Byteverschiebungen nach rechts durch
' ©reated 09/12/2000
Dim r As String
r = "0" & Mid(bincode, 1, Len(bincode) - 1)
SHR = r

End Function

Public Function ActiveConnection() As Boolean
On Error Resume Next
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
ActiveConnection = False
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
RETURNCODE = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, _
phkResult)

If RETURNCODE = ERROR_SUCCESS Then
    hKey = phkResult
    lpValueName = "Remote Connection"
    lpReserved = APINULL
    lpType = APINULL
    lpData = APINULL
    lpcbData = APINULL
    RETURNCODE = RegQueryValueEx(hKey, lpValueName, _
    lpReserved, lpType, ByVal lpData, lpcbData)
    lpcbData = Len(lpData)
    RETURNCODE = RegQueryValueEx(hKey, lpValueName, _
    lpReserved, lpType, lpData, lpcbData)
    
    If RETURNCODE = ERROR_SUCCESS Then
        If lpData = 0 Then
            ActiveConnection = False
        Else
            ActiveConnection = True
        End If
    End If


RegCloseKey (hKey)
End If

End Function

Public Function isTXTfile(FileName) As Boolean
On Error Resume Next
endung = Mid(FileName, Len(FileName) - 3)
If InStr(endung, ".") <= 0 Then endung = Mid(FileName, Len(FileName) - 3)
e = endung
If e = ".txt" Or e = ".bat" Or e = ".ini" Or e = ".inf" Or e = ".htm" Or e = ".html" Or e = ".php" Or e = ".php4" Or e = ".php3" Or e = ".asp" Or e = ".js" Then
    isTXTfile = True
Else
    isTXTfile = False
End If
End Function

Public Function ShortPath(Path As String) As String
On Error Resume Next
Dim kurz As String
kurz = Space$(265)
Call GetShortPathName(Path, kurz, Len(kurz))
ShortPath = kurz
End Function

Public Sub ScreenShot(Optional modus As Long = SSM_Desktop)
On Error Resume Next
If modus <> SSM_ActiveWindow Then
    keybd_event VK_SNAPSHOT, modus, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
Else
  keybd_event VK_MENU, 0, 0, 0
  keybd_event VK_SNAPSHOT, 0, 0, 0
  keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
  keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
End If
End Sub

' Registry functions below are by Kenneth Ives!
Public Function regDelete_Sub_Key(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String, _
                                  ByVal strRegSubKey As String)
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for removing a sub key.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be removed.
'
' Syntax:
'    regDelete_Sub_Key HKEY_CURRENT_USER, _
                  "Software\AAA-Registry Test\Products", "StringTestData"
'
' Removes the sub key "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the sub key.  If it does not exist, then ignore it.
      m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
  
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

Public Function regDoes_Key_Exist(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String) As Boolean
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function to see if a key does exist
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you want to test
'
' Syntax:
'    strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, _
'                       "Software\AAA-Registry Test\Products")
'
' Returns the value of TRUE or FALSE
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long

' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  lngKeyHandle = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regDoes_Key_Exist = False
  Else
      regDoes_Key_Exist = True
  End If
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

Public Function regQuery_A_Key(ByVal lngRootKey As Long, _
                               ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String) As Variant
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for querying a sub key value.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be queryed.
'
' Syntax:
'    strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, _
'                       "Software\AAA-Registry Test\Products", _
                        "StringTestData")
'
' Returns the key value of "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim intPosition As Integer
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngBufferSize As Long
  Dim lngBuffer As Long
  Dim strBuffer As String

' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  lngKeyHandle = 0
  lngBufferSize = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Query the registry and determine the data type.
' --------------------------------------------------------------
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, _
                         lngDataType, ByVal 0&, lngBufferSize)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Make the API call to query the registry based on the type
' of data.
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data (most common)
              ' Preload the receiving buffer area
              strBuffer = Space(lngBufferSize)
      
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, _
                                     ByVal strBuffer, lngBufferSize)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Strip out the string data
                  intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
                  If intPosition > 0 Then
                      ' if we found one, then save everything up to that point
                      regQuery_A_Key = Left(strBuffer, intPosition - 1)
                  Else
                      ' did not find one.  Save everything.
                      regQuery_A_Key = strBuffer
                  End If
              End If
              
         Case REG_DWORD:    ' Numeric data (Integer)
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                     lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Save the captured data
                  regQuery_A_Key = lngBuffer
              End If
         
         Case Else:    ' unknown
              regQuery_A_Key = ""
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function
Public Sub regCreate_Key_Value(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String, varRegData As Variant)
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for saving string data.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be updated.
'       varRegData - Update data.
'
' Syntax:
'    regCreate_Key_Value HKEY_CURRENT_USER, _
'                      "Software\AAA-Registry Test\Products", _
'                      "StringTestData", "22 Jun 1999"
'
' Saves the key value of "22 Jun 1999" to sub key "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngKeyValue As Long
  Dim strKeyValue As String
  
' --------------------------------------------------------------
' Determine the type of data to be updated
' --------------------------------------------------------------
  If IsNumeric(varRegData) Then
      lngDataType = REG_DWORD
  Else
      lngDataType = REG_SZ
  End If
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
    
' --------------------------------------------------------------
' Update the sub key based on the data type
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data
              strKeyValue = Trim(varRegData) & Chr(0)     ' null terminated
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          ByVal strKeyValue, Len(strKeyValue))
                                   
         Case REG_DWORD:    ' numeric data
              lngKeyValue = CLng(varRegData)
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          lngKeyValue, 4&)  ' 4& = 4-byte word (long integer)
                                   
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub
Public Function regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)

' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   This function will create a new key
'
' Parameters:
'          lngRootKey  - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'   strRegKeyPath  - is name of the key you wish to create.
'                  to make sub keys, continue to make this
'                  call with each new level.  MS says you
'                  can do this in one call; however, the
'                  best laid plans of mice and men ...
'
' Syntax:
'   regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test"
'   regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products"
' --------------------------------------------------------------

' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Create the key.  If it already exist, ignore it.
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)

' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function
Public Function regDelete_A_Key(ByVal lngRootKey As Long, _
                                ByVal strRegKeyPath As String, _
                                ByVal strRegKeyName As String) As Boolean
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for removing a complete key.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                        HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'   strRegKeyValue - is the name of the key which will be removed.
'
' Returns a True or False on completion.
'
' Syntax:
'    regDelete_A_Key HKEY_CURRENT_USER, "Software", "AAA-Registry Test"
'
' Removes the key "AAA-Registry Test" and all of its sub keys.
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Preset to a failed delete
' --------------------------------------------------------------
  regDelete_A_Key = False
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the key
      m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
      
      ' If the value returned is equal zero then we have succeeded
      If m_lngRetVal = 0 Then regDelete_A_Key = True
      
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Saves and Loads a form's control values and optionally window size/position.
' ------------------------------------------------------------------------------------
' Written by David Edlen - davidnedlen@cs.com
' Nov-19-2000
'
' A one-step process to save all control values on a form to the Windows registry
' and to retrieve previously saved values and re-apply them to the form.  Will
' automatically save the values for each text box, check box, option button,
' list box, and combo list.  Optionally, the window size and position can also be
' saved and retrieved.  No special coding is required for individual controls.
' No additional coding is required when new controls are added to the form.
' Especially useful for complex "Options" forms.  All settings remain in the Windows
' registry after the application terminates.
' =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Usage:
'
' SaveFormSettings "MyApp", frmOptions, true
'
'   Saves all control values plus the window size and position for form frmOptions.
'
' GetFormSettings "MyApp", frmOptions, true
'
'   Retrieves all previously saved control values for form frmOptions and assigns
'   those values to their respective controls; sets the form size and position to the
'   last saved settings.
'
' ------------------------------------------------------------------------------------
' Public Procedures:
'
' SaveFormSettings Statement
' ==========================
'   Saves or creates a series of entries in the application's entry in the
'   Windows registry.  The entries will consist of the values for each of the form's
'   text boxes, check boxes, option buttons, list boxes, and combo lists.  Optionally,
'   the form's window size and position are included in the entries.
'
' Syntax:
'   SaveFormSettings appname, form, [winsettings], [errnumber]
'   ----------------------------------------------------------
'   appname:        String expression containing the name of the application
'   form:           A form object to be saved
'   winsettings:    Boolean expression; window size and position settings will be
'                   saved if true.
'   errnumber:      Returned Long value: the code number of any procedure error.
'                   0 = no error.
' ------------------------------------------------------------------------------------
'
' GetFormSettings Statement
' =========================
' Retrieves all settings for a form saved by the SaveFormSettings statement and
' assigns them to their respective controls.  Optionally, the window size and
' position is also retrieved and applied.
'
' Syntax:
'   GetFormSettings appname, form, [winsettings], [errnumber]
'   ---------------------------------------------------------
'   appname:        String expression containing the name of the application
'   form:           A form object
'   winsettings:    Boolean expression; The last saved window size and position
'                   will be applied to the form if true.
'   errnumber:      Returned Long value: the code number of any procedure error.
'                   0 = no error.
' ------------------------------------------------------------------------------------
' Notes:
'   Uses the VB SaveSettings and GetSettings statements.
'   Registry entries are stored in:
'       HKEY_CURRENT_USER\Software\VB and VBA Program Settings
'
'   The registry entries are named as follows:
'   ------------------------------------------
'   Application name = appname parameter value
'   Section name = name of the form object
'   Key names = names of each of the form's controls
'   Settings = the text, value, or list items of each control.  List items are
'       saved as a single string with each item delimited by a chr(11).
' ------------------------------------------------------------------------------------

Public Sub SaveFormSettings(ByVal pAppName As String, pForm As Form, Optional pFormPosition As Boolean, Optional pError As Long)

    Dim ix As Long
    Dim vName As String
    Dim vControl As Control
    Dim vError As Long
    
    On Error GoTo errSaveFormSettings
    
    ' Windows Settings
    
    If pFormPosition = True Then
        SaveSetting pAppName, pForm.Name, WinHeight, pForm.Height
        SaveSetting pAppName, pForm.Name, WinWidth, pForm.Width
        SaveSetting pAppName, pForm.Name, WinTop, pForm.Top
        SaveSetting pAppName, pForm.Name, WinLeft, pForm.Left
    End If
    
    ' Loop through the form's control collection.
    ' Save the value parameter for each control.
    
    For Each vControl In pForm.Controls
        
        On Error Resume Next
        With vControl
            ix = .Index
            If Err.Number = 343 Then
                vName = .Name
                Err.Clear
            Else
                vName = .Name & ":" & Trim(CStr(ix))
            End If
        End With
        
        On Error GoTo errSaveFormSettings
        If TypeOf vControl Is TextBox Then
            SaveSetting pAppName, pForm.Name, vName, vControl.Text
        ElseIf TypeOf vControl Is CheckBox _
        Or TypeOf vControl Is OptionButton Then
            SaveSetting pAppName, pForm.Name, vName, vControl.Value
        ElseIf TypeOf vControl Is ListBox _
        Or TypeOf vControl Is ComboBox Then
            SaveSetting pAppName, pForm.Name, vName, GetListString(vControl, vError)
            If vError <> 0 Then Err.Raise vError
        End If
    
    Next vControl

    Set vControl = Nothing

errSaveFormSettings:
    
    pError = Err.Number

End Sub

Public Sub GetFormSettings(ByVal pAppName As String, pForm As Form, Optional pFormPosition As Boolean, Optional pError As Long)

    Dim ix As Long
    Dim vName As String
    Dim vControl As Control
    Dim vError As Long
    
    On Error GoTo errGetFormSettings
    
    ' Windows Settings
    
    If pFormPosition = True Then
        pForm.Height = GetSetting(pAppName, pForm.Name, WinHeight, pForm.Height)
        pForm.Width = GetSetting(pAppName, pForm.Name, WinWidth, pForm.Width)
        pForm.Top = GetSetting(pAppName, pForm.Name, WinTop, pForm.Top)
        pForm.Left = GetSetting(pAppName, pForm.Name, WinLeft, pForm.Left)
    End If
    
    ' Loop through the form's control collection.
    ' Retrieve the value parameter for each control.
    
    For Each vControl In pForm.Controls
        
        On Error Resume Next
        With vControl
            ix = .Index
            If Err.Number = 343 Then
                vName = .Name
                Err.Clear
            Else
                vName = .Name & ":" & Trim(CStr(ix))
            End If
        End With

        On Error GoTo errGetFormSettings
        If TypeOf vControl Is TextBox Then
            vControl.Text = GetSetting(pAppName, pForm.Name, vName, vControl.Text)
        ElseIf TypeOf vControl Is CheckBox _
        Or TypeOf vControl Is OptionButton Then
            vControl.Value = GetSetting(pAppName, pForm.Name, vName, vControl.Value)
        ElseIf TypeOf vControl Is ListBox _
        Or TypeOf vControl Is ComboBox Then
            PopulateList vControl, GetSetting(pAppName, pForm.Name, vName, ""), vError
            If vError <> 0 Then Err.Raise vError
        End If
    
    Next vControl

    Set vControl = Nothing

errGetFormSettings:
    
    pError = Err.Number

End Sub

Private Function GetListString(pControl As Control, pError As Long) As String

' Convert the contents of the specified list control to a string expression.
' The string will consist of all list items delimeted by a chr(11) (vbVerticalTab).

    Dim strList As Variant
    Dim ix As Long
    
    On Error GoTo errGetListString
    strList = ""
    
    If TypeOf pControl Is ListBox _
    Or TypeOf pControl Is ComboBox Then
        With pControl
            For ix = 0 To .ListCount - 1
                If strList <> "" Then
                    strList = strList & vbVerticalTab
                End If
                strList = strList & .List(ix)
            Next ix
        End With
    End If

    GetListString = strList
    
errGetListString:

    pError = Err.Number

End Function

Private Sub PopulateList(pControl As Control, pListString As String, pError As Long)

' Convert a list string to list items and populate the specified list control.

    Dim arList As Variant
    Dim ix As Integer
    
    On Error GoTo errPopulateList

    If TypeOf pControl Is ListBox _
    Or TypeOf pControl Is ComboBox Then
    
        pControl.Clear
        arList = Split(pListString, vbVerticalTab)
        
        If IsArray(arList) Then
            For ix = LBound(arList) To UBound(arList)
                pControl.AddItem arList(ix)
            Next ix
        Else
            pControl.AddItem arList
        End If
        
    End If
    
errPopulateList:

    pError = Err.Number

End Sub

Public Function NullTrim(NullStr As String) As String
    NullTrim = Left(NullStr, InStr(NullStr, Chr(0)) - 1)
End Function


' Create a complete path
'  IN: Path to be created, including drive and ending backslash, e.g.:
'      C:\Windows\Desktop\Hello\Test\
' OUT: FALSE on error, else TRUE
' Exists since Black-Box-Version: 1.0
Function DirMake(ByVal Fi As String) As Boolean
  On Error Resume Next
  If Right$(Fi, 1) <> "\" Then Fi = Fi + "\"
  Fi = Fi + "dummy" + vbNullChar$      ' !!! tricky!
  DirMake = (MakePath(Fi) <> 0)
  On Error GoTo 0
End Function

Public Function AddFavorite(SiteName As String, _
 URL As String) As Boolean
 
'PURPOSE:  ADDS a Favorite to IE 4 or 5 List of Favorites
'INPUT: SiteName = Name of Web Site
        'URL = URL FOR THE WEB SITE
'RETURNS: TRUE IF SUCCESSFUL, FALSE OTHERWISE

'EXAMPLE AddFavorite "FreeVBCode", "http://www.freevbcode.com"


Dim pidl As Long
Dim psFullPath As String
Dim iFile As Integer

On Error GoTo errorhandler
iFile = FreeFile
psFullPath = Space(255)

If SHGetSpecialFolderLocation(0, CSIDL_FAVORITES, pidl) _
  = 0 Then
    
   If pidl Then
      
      If SHGetPathFromIDList(pidl, psFullPath) Then
         
        psFullPath = TrimWithoutPrejudice(psFullPath)
        If Right(psFullPath, 1) <> "\" Then psFullPath = psFullPath & "\"
        psFullPath = psFullPath & SiteName & ".URL"
        Open psFullPath For Output As #iFile
        Print #iFile, "[InternetShortcut]"
        Print #iFile, "URL=" & URL
        Close #iFile
      
      End If
    
     CoTaskMemFree pidl
     AddFavorite = True
     
   End If

End If

errorhandler:
End Function

Public Function TrimWithoutPrejudice(ByVal InputString As String) As String

' Trim all non-printing characters from a string
' Snippet taken from http://www.freevbcode.com

Dim sAns            As String
Dim lLen            As Long
Dim lPtr            As Long

sAns = InputString
lLen = Len(InputString)

If lLen > 0 Then
    'LTrim
    For lPtr = 1 To lLen
        If Asc(Mid$(sAns, lPtr, 1)) > 32 Then Exit For
    Next

    sAns = Mid$(sAns, lPtr)
    lLen = Len(sAns)

    'RTtrim
    If lLen > 0 Then
        For lPtr = lLen To 1 Step -1
            If Asc(Mid$(sAns, lPtr, 1)) > 32 Then Exit For
        Next
    End If
    
    sAns = Left$(sAns, lPtr)

End If

TrimWithoutPrejudice = sAns

End Function

Public Function EndAllInstances(ByVal WindowCaption As String, Frm As Form) _
  As Boolean
'*********************************************
'PURPOSE: ENDS ALL RUNNING INSTANCES OF A PROCESS
'THAT CONTAINS ANY PART OF THE WINDOW CAPTION

'INPUT: ANY PART OF THE WINDOW CAPTION

'RETURNS: TRUE IF SUCCESSFUL (AT LEASE ONE PROCESS WAS STOPPED,
'FALSE OTHERWISE)

'EXAMPLE EndProcess "Notepad"

'NOTES:
'1. THIS IS DESIGNED TO TERMINATE THE PROCESS IMMEDIATELY,
'   THE APP WILL NOT RUN THROUGH IT'S NORMAL SHUTDOWN PROCEDURES
'   E.G., THERE WILL BE NO DIALOG BOXES LIKE "ARE YOU SURE
'   YOU WANT TO QUIT"

'2. BE CAREFUL WHEN USING:
'   E.G., IF YOU CALL ENDPROCESS("A"), ANY PROCESS WITH A
'   WINDOW THAT HAS THE LETTER "A" IN ITS CAPTION WILL BE
'   TERMINATED

'3. AS WRITTEN, ALL THIS CODE MUST BE PLACED WITHIN
'   A FORM MODULE

'***********************************************
Dim hwnd As Long
Dim hInst As Long
Dim hProcess As Long
Dim lProcessID
Dim bAns As Boolean
Dim lExitCode As Long
Dim lret As Long

On Error GoTo errorhandler

If Trim(WindowCaption) = "" Then Exit Function
Do
hwnd = FindWin(WindowCaption, Frm)
If hwnd = 0 Then Exit Do
hInst = GetWindowThreadProcessId(hwnd, lProcessID)
'Get handle to process
hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lProcessID)
If hProcess <> 0 Then
    'get exit code
    GetExitCodeProcess hProcess, lExitCode
        If lExitCode <> 0 Then
                'bye-bye
            lret = TerminateProcess(hProcess, lExitCode)
            If bAns = False Then bAns = lret > 0
        End If
End If
Loop

EndAllInstances = bAns
errorhandler:

End Function
Public Function FindWin(WinTitle As String, Frm As Form) As Long

Dim lhWnd As Long, sAns As String
Dim sTitle As String

lhWnd = Frm.hwnd
sTitle = LCase(WinTitle)

Do

   DoEvents
      If lhWnd = 0 Then Exit Do
        sAns = LCase$(GetCaption(lhWnd))
             

       If InStr(sAns, sTitle) Then

          FindWin = lhWnd
          Exit Do
       Else
         FindWin = 0
       End If

       lhWnd = GetNextWindow(lhWnd, 2)

Loop

End Function

Private Function GetCaption(lhWnd As Long) As String

Dim sAns As String, lLen As Long

   lLen& = GetWindowTextLength(lhWnd)
    sAns = String(lLen, 0)
    Call GetWindowText(lhWnd, sAns, lLen + 1)
   GetCaption = sAns

End Function

'CALL TO HIDE THE CURRENT APP
Public Sub HideMeFromTaskList()
    RegisterServiceProcess GetCurrentProcessId, 1
End Sub

'CALL TO DISPLAY THE CURRENT APP
Public Sub ShowMeInTaskList()
    RegisterServiceProcess GetCurrentProcessId, 0
End Sub

Public Sub CenterForm(Frm As Form)
'CenterForm me
    Dim x%, y%
    x = Screen.Width / 2 - Frm.Width / 2
    y = Screen.Height / 2 - Frm.Height / 2
    Frm.Move x, y
End Sub

Public Sub CenterFormInMDI(Frm As Form, mdi As MDIForm)
'CenterForm me
    Dim x%, y%
    x = mdi.Width / 2 - Frm.Width / 2
    y = mdi.Height / 2 - Frm.Height / 2
    Frm.Move x, y
End Sub

Public Sub CenterObject(obj As Object, Frm As Form)
On Error Resume Next
Dim x, y
x = Frm.Width / 2 - obj.Width / 2
y = Frm.Height / 2 - obj.Height / 2
obj.Left = x
obj.Top = y
End Sub

Public Sub CenterObjectOnScreen(obj As Object)
On Error Resume Next
Dim x, y
x = Screen.Width / 2 - obj.Width / 2
y = Screen.Height / 2 - obj.Height / 2
obj.Left = x
obj.Top = y
End Sub

Private Function ActualPos(plLeft As Long) As Long
    If plLeft < 0 Then
        ActualPos = plLeft + 75000
    Else
        ActualPos = plLeft
    End If
End Function
Private Function FindForm(pfrmIn As Form) As Long
Dim i As Long
    FindForm = -1
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                FindForm = i
                Exit Function
            End If
        Next i
    End If
End Function
Private Function AddForm(pfrmIn As Form) As Long
Dim FormControl As Control
Dim i As Long
    ReDim Preserve FormRecord(MaxForm + 1)
    FormRecord(MaxForm).Name = pfrmIn.Name
    FormRecord(MaxForm).Top = pfrmIn.Top
    FormRecord(MaxForm).Left = pfrmIn.Left
    FormRecord(MaxForm).Height = pfrmIn.Height
    FormRecord(MaxForm).Width = pfrmIn.Width
    FormRecord(MaxForm).ScaleHeight = pfrmIn.ScaleHeight
    FormRecord(MaxForm).ScaleWidth = pfrmIn.ScaleWidth
    AddForm = MaxForm
    MaxForm = MaxForm + 1
    For Each FormControl In pfrmIn
        i = FindControl(FormControl, pfrmIn.Name)
        If i < 0 Then
            i = AddControl(FormControl, pfrmIn.Name)
        End If
    Next FormControl
End Function
Private Function FindControl(inControl As Control, inName As String) As Long
Dim i As Long
    FindControl = -1
    For i = 0 To (MaxControl - 1)
        If ControlRecord(i).Parrent = inName Then
            If ControlRecord(i).Name = inControl.Name Then
                On Error Resume Next
                If ControlRecord(i).Index = inControl.Index Then
                    FindControl = i
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Function
Private Function AddControl(inControl As Control, inName As String) As Long
    ReDim Preserve ControlRecord(MaxControl + 1)
    On Error Resume Next
    ControlRecord(MaxControl).Name = inControl.Name
    ControlRecord(MaxControl).Index = inControl.Index
    ControlRecord(MaxControl).Parrent = inName
    If TypeOf inControl Is Line Then
        ControlRecord(MaxControl).Top = inControl.Y1
        ControlRecord(MaxControl).Left = ActualPos(inControl.X1)
        ControlRecord(MaxControl).Height = inControl.Y2
        ControlRecord(MaxControl).Width = ActualPos(inControl.x2)
    Else
        ControlRecord(MaxControl).Top = inControl.Top
        ControlRecord(MaxControl).Left = ActualPos(inControl.Left)
        ControlRecord(MaxControl).Height = inControl.Height
        ControlRecord(MaxControl).Width = inControl.Width
    End If
    On Error GoTo 0
    AddControl = MaxControl
    MaxControl = MaxControl + 1
End Function
Private Function PerWidth(pfrmIn As Form) As Long
Dim i As Long
    i = FindForm(pfrmIn)
    If i < 0 Then
        i = AddForm(pfrmIn)
    End If
    PerWidth = (pfrmIn.ScaleWidth * 100) \ FormRecord(i).ScaleWidth
End Function
Private Function PerHeight(pfrmIn As Form) As Single
Dim i As Long
    i = FindForm(pfrmIn)
    If i < 0 Then
        i = AddForm(pfrmIn)
    End If
    PerHeight = (pfrmIn.ScaleHeight * 100) \ FormRecord(i).ScaleHeight
End Function
Private Sub ResizeControl(inControl As Control, pfrmIn As Form)
Dim i As Long
Dim yRatio, xRatio, lTop, lLeft, lWidth, lHeight As Long
    yRatio = PerHeight(pfrmIn)
    xRatio = PerWidth(pfrmIn)
    i = FindControl(inControl, pfrmIn.Name)
    On Error GoTo Moveit
    If inControl.Left < 0 Then
        lLeft = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
    Else
        lLeft = CLng((ControlRecord(i).Left * xRatio) \ 100)
    End If
    lTop = CLng((ControlRecord(i).Top * yRatio) \ 100)
    lWidth = CLng((ControlRecord(i).Width * xRatio) \ 100)
    lHeight = CLng((ControlRecord(i).Height * yRatio) \ 100)
    GoTo Moveit
Moveit:
    On Error GoTo MoveError1
    If TypeOf inControl Is Line Then
        If inControl.X1 < 0 Then
            inControl.X1 = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
        Else
            inControl.X1 = CLng((ControlRecord(i).Left * xRatio) \ 100)
        End If
        inControl.Y1 = CLng((ControlRecord(i).Top * yRatio) \ 100)
        If inControl.x2 < 0 Then
            inControl.x2 = CLng(((ControlRecord(i).Width * xRatio) \ 100) - 75000)
        Else
            inControl.x2 = CLng((ControlRecord(i).Width * xRatio) \ 100)
        End If
        inControl.Y2 = CLng((ControlRecord(i).Height * yRatio) \ 100)
    Else
        If TypeOf inControl Is Timer Then
            GoTo subExit
        End If
        inControl.Move lLeft, lTop, lWidth, lHeight
    End If
    GoTo subExit
MoveError1:
    On Error GoTo MoveError2
    inControl.Move lLeft, lTop, lWidth
    GoTo subExit
MoveError2:
    On Error GoTo subExit
    inControl.Move lLeft, lTop
subExit:
    On Error GoTo 0
End Sub
Public Sub ResizeForm(pfrmIn As Form)
Dim FormControl As Control
Dim isVisible As Boolean
If pfrmIn.Top < 30000 Then
    isVisible = pfrmIn.Visible
    pfrmIn.Visible = False
    For Each FormControl In pfrmIn
        ResizeControl FormControl, pfrmIn
    Next FormControl
    pfrmIn.Visible = isVisible
End If
End Sub
Public Sub SaveFormPosition(pfrmIn As Form)
Dim i As Long
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                FormRecord(i).Top = pfrmIn.Top
                FormRecord(i).Left = pfrmIn.Left
                FormRecord(i).Height = pfrmIn.Height
                FormRecord(i).Width = pfrmIn.Width
                Exit Sub
            End If
        Next i
        AddForm (pfrmIn)
    End If
End Sub
Public Sub RestoreFormPosition(pfrmIn As Form)
Dim i As Long
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                If FormRecord(i).Top < 0 Then
                    pfrmIn.WindowState = 2
                ElseIf FormRecord(i).Top < 30000 Then
                    pfrmIn.WindowState = 0
                    pfrmIn.Move FormRecord(i).Left, FormRecord(i).Top, FormRecord(i).Width, FormRecord(i).Height
                Else
                    pfrmIn.WindowState = 1
                End If
                Exit Sub
            End If
        Next i
    End If
End Sub

Public Sub AboutMeLPro(Frm As Form, icohwnd As Long)
On Error Resume Next
ShellAbout Frm.hwnd, "This application uses MeLPro", "Avalible at http://www.melaxis.com", icohwnd
End Sub

Public Function Hex0(HexNum As String)
On Error Resume Next
' by mel ;-)
Hex0 = String(8 - Len(HexNum), "0") & HexNum
End Function

' This function brings you the missing PAUSE (or DELAY)-command in VB
' and allows MultiTasking (same as PauseMT, but more comfortable).
' Exists since Black-Box-Version: 2.1
Sub Wait(hours As Integer, minutes As Integer, seconds As Integer)
  Dim ti, ti2
  Dim ho As Integer, mi As Integer, se As Integer
  
  ho = Val(Mid$(Format$(Now, "hh:mm:ss"), 1, 2))
  mi = Val(Mid$(Format$(Now, "hh:mm:ss"), 4, 2))
  se = Val(Mid$(Format$(Now, "hh:mm:ss"), 7, 2))
  ti = TimeSerial(ho, mi, se)
  ti2 = TimeSerial(ho + hours, mi + minutes, se + seconds)
  Do
    DoEvents
    ho = Val(Mid$(Format$(Now, "hh:mm:ss"), 1, 2))
    mi = Val(Mid$(Format$(Now, "hh:mm:ss"), 4, 2))
    se = Val(Mid$(Format$(Now, "hh:mm:ss"), 7, 2))
    ti = TimeSerial(ho, mi, se)
    'Debug.Print ti, ti2
    If ti >= ti2 Then Exit Do
  Loop
End Sub

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

Public Function AvailablePhysicalMemory() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    AvailablePhysicalMemory = BytesToMegabytes(dblAns)
    
End Function

Public Function TotalPhysicalMemory() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPhys
    TotalPhysicalMemory = BytesToMegabytes(dblAns)
End Function

Public Function PercentMemoryFree() As Double

   PercentMemoryFree = Format(AvailableMemory / TotalMemory * _
   100, "0#")
End Function

Public Function AvailablePageFile() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPageFile
    AvailablePageFile = BytesToMegabytes(dblAns)
End Function

Public Function PageFileSize() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPageFile
    PageFileSize = BytesToMegabytes(dblAns)

End Function

Public Function AvailableMemory() As Double
'Return Value in Megabytes
     AvailableMemory = AvailablePhysicalMemory + AvailablePageFile
End Function

Public Function TotalMemory() As Double
'Return Value in Megabytes
    TotalMemory = PageFileSize + TotalPhysicalMemory
End Function

Private Function BytesToMegabytes(Bytes As Double) As Double
 
  Dim dblAns As Double
  dblAns = (Bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
  
End Function

'**************************************
' Name: HackerScan Routine
' Description:This code will scan for po
'     pular hacking tools: FileMon, RegMon and
'     SoftICE (both Win 9x and NT versions). T
'     his code was inspired by the SoftICE det
'     ection routine by Joox (http://www.plane
'     tsourcecode.com/vb/scripts/ShowCode.asp?
'     lngWId=1&txtCodeId=7600). If any of thes
'     e programs are in memory an access viola
'     tion is generated. You should call this
'     routine before you read or write any sen
'     sitive information (ie license files) to
'     files or the regsitry.
'     i 'm certain that there are workarounds For this code, but its intent is To make things harder for the hacker.
'     I would love To see other methods added to this to detect other debuggers, tools, etc. so please leave whatever comments come to mind. Go ahead and vote too!
'      Enjoy!
' By: Kevin Lingofelter
'
' Assumes:Simply call this routine befor
'     e doing any sensitive reading or writing
'     to files or the registry...ie license in
'     formation.
'
' Side Effects:Acces violations, but it
'     is by design. See the comment in the cod
'     e for details.
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.10000/lngWId.1/qx/vb/scripts/ShowCode
'     .htm'for details.'**************************************



Public Sub HackerScan()
    Dim hFile As Long, retVal As Long
    Dim sRegMonClass As String, sFileMonClass As String
    Check4Manipulation      ' diese zeile wurde von pablo hoch eingefügt *fg*
    '\\We break up the class names to avoid
    '     detection in a hex editor
    sRegMonClass = "R" & "e" & "g" & "m" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    'sRegMonClass = UR & LE & LG & LM & LO & LN & UC & LL & LA & LS & LS & LS
    sFileMonClass = "F" & "i" & "l" & "e" & "M" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    'sFileMonClass = UF & LI & LL & LE & UM & LP & LN & UC & LL & LA & LS & LS
    '\\See if RegMon or FileMon are running


    Select Case True
        Case FindWindow(sRegMonClass, vbNullString) <> 0
        'Regmon is running...throw an access vio
        '     lation
        RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
        Case FindWindow(sFileMonClass, vbNullString) <> 0
        'FileMon is running...throw an access vi
        '     olation
        RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
    End Select
'\\So far so good...check for SoftICE in
'     memory
hFile = CreateFile("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)


If hFile <> -1 Then
    ' SoftICE is detected.
    retVal = CloseHandle(hFile) ' Close the file handle
    RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
Else
    ' SoftICE is not found for windows 9x, c
    '     heck for NT.
    hFile = CreateFile("\\.\NTICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)


    If hFile <> -1 Then
        ' SoftICE is detected.
        retVal = CloseHandle(hFile) ' Close the file handle
        RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
    End If
End If
End Sub

Public Function DebuggerRunning() As Boolean
    Dim hFile As Long, retVal As Long
    Dim sRegMonClass As String, sFileMonClass As String
    If WurdeManipuliert Then DebuggerRunning = True: Exit Function      ' diese zeile wurde von pablo hoch eingefügt *fg*
    '\\We break up the class names to avoid
    '     detection in a hex editor
    sRegMonClass = "R" & "e" & "g" & "m" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    'sRegMonClass = UR & LE & LG & LM & LO & LN & UC & LL & LA & LS & LS & LS
    sFileMonClass = "F" & "i" & "l" & "e" & "M" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    'sFileMonClass = UF & LI & LL & LE & UM & LP & LN & UC & LL & LA & LS & LS
    '\\See if RegMon or FileMon are running


    Select Case True
        Case FindWindow(sRegMonClass, vbNullString) <> 0
        'Regmon is running...throw an access vio
        '     lation
        DebuggerRunning = True: Exit Function
        Case FindWindow(sFileMonClass, vbNullString) <> 0
        'FileMon is running...throw an access vi
        '     olation
        DebuggerRunning = True: Exit Function
    End Select
'\\So far so good...check for SoftICE in
'     memory
hFile = CreateFile("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)


If hFile <> -1 Then
    ' SoftICE is detected.
    retVal = CloseHandle(hFile) ' Close the file handle
    DebuggerRunning = True: Exit Function
Else
    ' SoftICE is not found for windows 9x, c
    '     heck for NT.
    hFile = CreateFile("\\.\NTICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)


    If hFile <> -1 Then
        ' SoftICE is detected.
        retVal = CloseHandle(hFile) ' Close the file handle
        DebuggerRunning = True: Exit Function
    End If
End If
DebuggerRunning = False
End Function

Public Sub Check4Manipulation()
On Error Resume Next
' von pablo hoch ;-)
' überprüft, ob die buchstaben per hexeditor manipuliert wurden

If UB <> "B" Then GoTo manipuliert
If UC <> "C" Then GoTo manipuliert
If UD <> "D" Then GoTo manipuliert
If UE <> "E" Then GoTo manipuliert
If UF <> "F" Then GoTo manipuliert
If UG <> "G" Then GoTo manipuliert
If UH <> "H" Then GoTo manipuliert
If UI <> "I" Then GoTo manipuliert
If UJ <> "J" Then GoTo manipuliert
If UK <> "K" Then GoTo manipuliert
If ul <> "L" Then GoTo manipuliert
If UM <> "M" Then GoTo manipuliert
If UN <> "N" Then GoTo manipuliert
If UO <> "O" Then GoTo manipuliert
If UP <> "P" Then GoTo manipuliert
If UQ <> "Q" Then GoTo manipuliert
If UR <> "R" Then GoTo manipuliert
If US <> "S" Then GoTo manipuliert
If ut <> "T" Then GoTo manipuliert
If UU <> "U" Then GoTo manipuliert
If UV <> "V" Then GoTo manipuliert
If UW <> "W" Then GoTo manipuliert
If UX <> "X" Then GoTo manipuliert
If UY <> "Y" Then GoTo manipuliert
If UZ <> "Z" Then GoTo manipuliert
If LA <> "a" Then GoTo manipuliert
If LB <> "b" Then GoTo manipuliert
If LC <> "c" Then GoTo manipuliert
If LD <> "d" Then GoTo manipuliert
If le <> "e" Then GoTo manipuliert
If LF <> "f" Then GoTo manipuliert
If LG <> "g" Then GoTo manipuliert
If LH <> "h" Then GoTo manipuliert
If LI <> "i" Then GoTo manipuliert
If LJ <> "j" Then GoTo manipuliert
If LK <> "k" Then GoTo manipuliert
If LL <> "l" Then GoTo manipuliert
If LM <> "m" Then GoTo manipuliert
If ln <> "n" Then GoTo manipuliert
If LO <> "o" Then GoTo manipuliert
If lp <> "p" Then GoTo manipuliert
If LQ <> "q" Then GoTo manipuliert
If LR <> "r" Then GoTo manipuliert
If ls <> "s" Then GoTo manipuliert
If LT <> "t" Then GoTo manipuliert
If LU <> "u" Then GoTo manipuliert
If LV <> "v" Then GoTo manipuliert
If LW <> "w" Then GoTo manipuliert
If LX <> "x" Then GoTo manipuliert
If LY <> "y" Then GoTo manipuliert
If LZ <> "z" Then GoTo manipuliert
' alles ok ;o)

GoTo ende
manipuliert:
' datei wurde manipuliert
RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
GoTo ende

ende:
End Sub

Public Function WurdeManipuliert() As Boolean
On Error Resume Next
' von pablo hoch ;-)
' überprüft, ob die buchstaben per hexeditor manipuliert wurden

If UB <> "B" Then GoTo manipuliert
If UC <> "C" Then GoTo manipuliert
If UD <> "D" Then GoTo manipuliert
If UE <> "E" Then GoTo manipuliert
If UF <> "F" Then GoTo manipuliert
If UG <> "G" Then GoTo manipuliert
If UH <> "H" Then GoTo manipuliert
If UI <> "I" Then GoTo manipuliert
If UJ <> "J" Then GoTo manipuliert
If UK <> "K" Then GoTo manipuliert
If ul <> "L" Then GoTo manipuliert
If UM <> "M" Then GoTo manipuliert
If UN <> "N" Then GoTo manipuliert
If UO <> "O" Then GoTo manipuliert
If UP <> "P" Then GoTo manipuliert
If UQ <> "Q" Then GoTo manipuliert
If UR <> "R" Then GoTo manipuliert
If US <> "S" Then GoTo manipuliert
If ut <> "T" Then GoTo manipuliert
If UU <> "U" Then GoTo manipuliert
If UV <> "V" Then GoTo manipuliert
If UW <> "W" Then GoTo manipuliert
If UX <> "X" Then GoTo manipuliert
If UY <> "Y" Then GoTo manipuliert
If UZ <> "Z" Then GoTo manipuliert
If LA <> "a" Then GoTo manipuliert
If LB <> "b" Then GoTo manipuliert
If LC <> "c" Then GoTo manipuliert
If LD <> "d" Then GoTo manipuliert
If le <> "e" Then GoTo manipuliert
If LF <> "f" Then GoTo manipuliert
If LG <> "g" Then GoTo manipuliert
If LH <> "h" Then GoTo manipuliert
If LI <> "i" Then GoTo manipuliert
If LJ <> "j" Then GoTo manipuliert
If LK <> "k" Then GoTo manipuliert
If LL <> "l" Then GoTo manipuliert
If LM <> "m" Then GoTo manipuliert
If ln <> "n" Then GoTo manipuliert
If LO <> "o" Then GoTo manipuliert
If lp <> "p" Then GoTo manipuliert
If LQ <> "q" Then GoTo manipuliert
If LR <> "r" Then GoTo manipuliert
If ls <> "s" Then GoTo manipuliert
If LT <> "t" Then GoTo manipuliert
If LU <> "u" Then GoTo manipuliert
If LV <> "v" Then GoTo manipuliert
If LW <> "w" Then GoTo manipuliert
If LX <> "x" Then GoTo manipuliert
If LY <> "y" Then GoTo manipuliert
If LZ <> "z" Then GoTo manipuliert
' alles ok ;o)

WurdeManipuliert = False
GoTo ende
manipuliert:
' datei wurde manipuliert
WurdeManipuliert = True: Exit Function
GoTo ende

ende:
End Function

Public Sub ShowInTaskManager()
On Error Resume Next
RegisterServiceProcess GetCurrentProcessId(), RSP_UNREGISTER_SERVICE
End Sub

Public Sub HideInTaskManager()
On Error Resume Next
RegisterServiceProcess GetCurrentProcessId(), RSP_SIMPLE_SERVICE
End Sub

Public Function GetCRCforFile(TheFile As String) As String
Dim filenum As Long
On Error GoTo Err:
'you must speify the complete file and directory the file is in.
    
    Dim lCrc32Value As Long
    Dim CRCStr As String * 8
    Dim FL As Long  'file length
    On Error Resume Next
    Dim FileStr$
    FL = FileLen(TheFile)
    FileStr$ = String(FL, 0)
    filenum = FreeFile
    Open TheFile For Binary As #filenum
     Get #filenum, 1, FileStr$
    Close #filenum
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(FileStr$, lCrc32Value)
    Dim RealCRC As String * 8
    RealCRC = CStr(Hex$(GetCrc32(lCrc32Value)))
    
    'This is to just infom you that your crc has been generated. you can remove this msgbox
    'MsgBox "This is the CRC32 that was generated for the file you askd for: " & RealCRC, vbInformation, "CRC Completed"
    'end of msgbox
    
    GetCRCforFile = RealCRC
    Exit Function
Err:
    'MsgBox "An error has been Reported by GOP CRC Wizard v1.00" & vbCrLf & vbCrLf & "Message Generated By : Function GetCRCforFile()", vbCritical, "GOP CRC Wizard v1.00"
End Function



Public Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    '// Declare counter variable iBytes, counter variable iBits, value variables lCrc32 and lTempCrc32
    Dim ibytes As Integer, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate 256 times

    For ibytes = 0 To 255
        '// Initiate lCrc32 to counter variable
        lCrc32 = ibytes
        '// Now iterate through each bit in counter byte


        For iBits = 0 To 7
            '// Right shift unsigned long 1 bit
            lTempCrc32 = lCrc32 And &HFFFFFFFE
            lTempCrc32 = lTempCrc32 \ &H2
            lTempCrc32 = lTempCrc32 And &H7FFFFFFF
            '// Now check if temporary is less than zero and then mix Crc32 checksum with Seed value


            If (lCrc32 And &H1) <> 0 Then
                lCrc32 = lTempCrc32 Xor Seed
            Else
                lCrc32 = lTempCrc32
            End If
        Next
        '// Put Crc32 checksum value in the holding array
        Crc32Table(ibytes) = lCrc32
    Next
    '// After this is done, set function value to the precondition value
    InitCrc32 = Precondition
End Function
'// The function above is the initializing function, now we have to write the computation function


Public Function AddCrc32(ByVal Item As String, ByVal CRC32 As Long) As Long
    '// Declare following variables
    Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
    Dim lAccValue As Long, lTableValue As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate through the string that is to be checksum-computed


    For iCounter = 1 To Len(Item)
        '// Get ASCII value for the current character
        bCharValue = Asc(Mid$(Item, iCounter, 1))
        '// Right shift an Unsigned Long 8 bits
        lAccValue = CRC32 And &HFFFFFF00
        lAccValue = lAccValue \ &H100
        lAccValue = lAccValue And &HFFFFFF
        '// Now select the right adding value from the holding table
        lIndex = CRC32 And &HFF
        lIndex = lIndex Xor bCharValue
        lTableValue = Crc32Table(lIndex)
        '// Then mix new Crc32 value with previous accumulated Crc32 value
        CRC32 = lAccValue Xor lTableValue
    Next
    '// Set function value the the new Crc32 checksum
    AddCrc32 = CRC32
End Function
'// At last, we have to write a function so that we can get the Crc32 checksum value at any time


Public Function GetCrc32(ByVal CRC32 As Long) As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Set function to the current Crc32 value
    GetCrc32 = CRC32 Xor &HFFFFFFFF
End Function
'// To Test the Routines Above...


Public Function Compute(ToGet As String) As String
    Dim lCrc32Value As Long
    On Error Resume Next
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(ToGet, lCrc32Value)
    Compute = Hex$(GetCrc32(lCrc32Value))
End Function

Public Function WriteAByte(gamewindowtext As String, Address As Long, Value As Byte)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
Exit Function
End If
WriteProcessMemory phandle, Address, Value, 1, 0&
CloseHandle hProcess
End Function

Public Function WriteAnInt(gamewindowtext As String, Address As Long, Value As Integer)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
Exit Function
End If
WriteProcessMemory phandle, Address, Value, 2, 0&
CloseHandle hProcess
End Function

Public Function WriteALong(gamewindowtext As String, Address As Long, Value As Long)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
Exit Function
End If
WriteProcessMemory phandle, Address, Value, 4, 0&
CloseHandle hProcess
End Function

Public Function ReadAByte(gamewindowtext As String, Address As Long, valbuffer As Byte)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
Exit Function
End If
ReadProcessMem phandle, Address, valbuffer, 1, 0&
CloseHandle hProcess
End Function
Public Function ReadAnInt(gamewindowtext As String, Address As Long, valbuffer As Integer)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
Exit Function
End If
ReadProcessMem phandle, Address, valbuffer, 2, 0&
CloseHandle hProcess
End Function

Public Function ReadALong(gamewindowtext As String, Address As Long, valbuffer As Long)
Dim hwnd As Long
Dim pid As Long
Dim phandle As Long
hwnd = FindWindow(vbNullString, gamewindowtext)
If (hwnd = 0) Then
Exit Function
End If
GetWindowThreadProcessId hwnd, pid
phandle = OpenProcess(PROCESS_ALL_ACCESS, False, pid)
If (phandle = 0) Then
Exit Function
End If
ReadProcessMem phandle, Address, valbuffer, 4, 0&
CloseHandle hProcess
End Function

Public Function Text2Bin(EIN As String) As String
On Error Resume Next
Dim r As String
For i = 1 To Len(EIN)
    r = r & Char2Bin(Mid(EIN, i, 1))
    DoEvents
Next
Text2Bin = r
End Function

Public Function Bin2Text(EIN As String) As String
On Error Resume Next
Dim r As String
For i = 1 To Len(EIN) Step 8
    r = r & Bin2Char(Mid(EIN, i, 8))
    DoEvents
Next
Bin2Text = r
End Function

Public Function LocalTime(ByVal lValue As Long) As String
On Error Resume Next
' manipuliert von mel, wegen deutschem format ;-)
    ' Now for the LocalTime function. Take t
    '     he long value representing the number
    ' of seconds since January 1, 1970 and c
    '     reate a useable time structure from it.
    ' Return a formatted string YYYY/MM/DD H
    '     H:MM:SS
    Dim lSecPerYear
    Dim Year As Long
    Dim Month As Long
    Dim Day As Long
    Dim Hour As Long
    Dim Minute As Long
    Dim Second As Long
    Dim Temp As Long
    ' [0] = normal year, [1] = leap year
    lSecPerYear = Array(31536000, 31622400)
    lSecPerDay = 86400 ' 60*60*24
    lSecPerHour = 3600 ' 60 * 60
    Year = 70 ' starting point
    ' Calculate the year


    Do While (lValue > 0)
        Temp = isLeapYear(Year)


        If (lValue - lSecPerYear(Temp)) > 0 Then
            lValue = lValue - lSecPerYear(Temp)
            Year = Year + 1
        Else
            Exit Do
        End If
    Loop
    If Year < 100 Then Year = Year + 1900
    'Debug.Print "Year = " & Year
    ' Calculate the month
    Month = 1


    Do While (lValue > 0)
        Temp = secsInMonth(Year, Month)


        If (lValue - Temp) > 0 Then
            lValue = lValue - Temp
            Month = Month + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Month = " & Month
    ' Now calculate day
    Day = 1


    Do While (lValue > 0)


        If (lValue - lSecPerDay) > 0 Then
            lValue = lValue - lSecPerDay
            Day = Day + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Day = " & Day
    ' Now calculate Hour
    Hour = 0


    Do While (lValue > 0)


        If (lValue - lSecPerHour) > 0 Then
            lValue = lValue - lSecPerHour
            Hour = Hour + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Hour = " & Hour
    Minute = Int(lValue / 60)
    'Debug.Print "Minute = " & Minute
    Second = lValue Mod 60
    'Debug.Print "Second = " & Second
    ' Standard date format is YYYY/MM/DD HH:
    '     MM:SS
    If Len(CStr(Year)) = 3 And Year >= 100 Then Year = Year + 2000 - 100
    'LocalTime = Year & "/" & Month & "/" & Day & ", " & _
    Hour & ":" & Minute & ":" & Second
    Day = DateAdd0(Day)
    Month = DateAdd0(Month)
    Year = DateAdd0(Year)
    Hour = DateAdd0(Hour)
    Minute = DateAdd0(Minute)
    Second = DateAdd0(Second)
    LocalTime = IIf(Day < 10, "0" & Day, Day) & "." & IIf(Month < 10, "0" & Month, Month) & "." & Year & " " & IIf(Hour < 10, "0" & Hour, Hour) & ":" & IIf(Minute < 10, "0" & Minute, Minute) & ":" & IIf(Second < 10, "0" & Second, Second)
    'LocalTime = Day & "." & Month & "." & Year & " " & Hour & ":" & Minute & ":" & Second
End Function


Public Function mktm(pdatetime As tm) As Long
On Error Resume Next
    Dim lResult As Long
    Dim nMonth As Long
    ' Time validation
    If (pdatetime.tm_sec < 0 Or pdatetime.tm_sec > 59 Or _
    pdatetime.tm_min < 0 Or pdatetime.tm_min > 59 Or _
    pdatetime.tm_hour < 0 Or pdatetime.tm_hour > 23) Then
    kcl_mktm = -1
    Exit Function
End If
' Date validation. This routine bites it
'     in 2038
If (pdatetime.tm_year < 70 Or pdatetime.tm_year > 138 Or _
pdatetime.tm_mon < 1 Or pdatetime.tm_mon > 12 Or _
pdatetime.tm_mday < 1 Or pdatetime.tm_mday > 31) Then
kcl_mktm = -1
Exit Function
End If
' Sum seconds in previous whole years
lResult = 0 ' Initialize
lResult = lResult + secsInYears(pdatetime.tm_year)
' Sum seconds in previous whole months f
'     or the current year


For nMonth = 1 To (pdatetime.tm_mon - nMonth) - 1
lResult = lResult + secsInMonth(pdatetime.tm_year, nMonth)
Next nMonth
' Sum seconds in whole days for the curr
'     ent month
lResult = lResult + ((pdatetime.tm_mday - 1) * 86400)
' Sum seconds in whole hours for the cur
'     rent day
lResult = lResult + CDec((pdatetime.tm_hour * 3600))
' Sum seconds in whole minutes for the c
'     urrent hour
lResult = lResult + (pdatetime.tm_min * 60)
' Sum remaining seconds for the current
'     minute
lResult = lResult + (pdatetime.tm_sec)
mktm = lResult
End Function
'*******************P R I V A T E*******
'     *******************


Private Function isLeapYear(Year As Long) As Integer
On Error Resume Next
    ' Determine if given ANSI datetime struc
    '     t represents a leap year
    ' Private function: assumes valid parame
    '     ters
    Dim nYear As Integer
    Dim nIsLeap As Integer
    nYear = Year + 1900


    If ((nYear Mod 4 = 0 And Not (nYear Mod 100)) Or nYear Mod 400 = 0) Then
        nIsLeap = 1 ' its a leap year
    Else
        nIsLeap = 0 ' Not a leap year
    End If
    isLeapYear = nIsLeap
End Function


Private Function secsInMonth(Year As Long, Month As Long) As Long
On Error Resume Next
    ' Return total number of seconds in the
    '     given month
    ' Private function: assumes valid parame
    '     ters
    Dim lResult As Long
    Dim lSecPerMonth
    lSecPerMonth = Array(2678400, 2419200, 2678400, 2592000, _
    2678400, 2592000, 2678400, 2678400, _
    2592000, 2678400, 2592000, 2678400)
    ' Compute result
    lResult = lSecPerMonth(Month - 1)


    If (isLeapYear(Year) And Month = 2) Then
        lResult = lResult + 86400 ' its February In a leap year
    End If
    secsInMonth = lResult
End Function


Private Function secsInYears(Year As Long) As Double
On Error Resume Next
    ' Return sum of seconds for years since
    '     Jan 1, 1970 00:00
    ' up to but excluding the given year.
    ' Private function: assumes valid parame
    '     ters
    Dim lResult As Long
    Dim thisYear As Long
    Dim Temp As Long
    lResult = 0
    ' 0 = normal year, 1 = leap year
    Dim lSecPerYear
    lSecPerYear = Array(31536000, 31622400)


    If (Year > 97) Then
        ' shorten summation iterations for typic
        '     al cases
        lResult = 883612800 ' seconds To Jan 1,1998 00:00:00
        thisYear = 98
    Else
        ' sum all years since 1970
        thisYear = 70
    End If
    ' Sum total seconds since Jan 1, 1970 00
    '     :00


    While (thisYear < Year)
        'for ( ; thisYear < year; thisYear++
        '     )
        Temp = isLeapYear(thisYear)
        lResult = lResult + lSecPerYear(Temp)
        thisYear = thisYear + 1
    Wend
    secsInYears = lResult
End Function

Public Function DateAdd0(Eingabe As Long) As String
On Error Resume Next
Dim dat As String
dat = CStr(Eingabe)
If Len(dat) < 2 Then dat = "0" & dat
DateAdd0 = dat
End Function

'Die nachfolgende Funktion führt die Registrierung durch
Public Function RegisterFile(ByVal sFile As String, _
  Register As Boolean) As Boolean

  'Der Parameter sFile enthält die zu
  'registrierende Datei (inkl. Pfad)
  'Register: True  -> Datei soll registriert werden
  '          False -> Datei soll deregistriert werden

  Dim Result As Boolean
  Dim Lib As Long
  Dim sProc As String
  Dim r1 As Long
  Dim r2 As Long
  Dim Thread As Long

  On Local Error GoTo RegError

  Result = False
  Lib = LoadLibrary(sFile)
  If Lib Then
    sProc = IIf(Register, "DllRegisterServer", _
      "DllUnregisterServer")
    r1 = GetProcAddress(Lib, sProc)
    If r1 Then
      Thread = CreateThread(ByVal 0, 0, ByVal r1, _
                         ByVal 0, 0, r2)
      If Thread Then
        r2 = WaitForSingleObject(Thread, 10000)
        If r2 Then
          'Fehler aufgetreten
          FreeLibrary Lib
          r2 = GetExitCodeThread(Thread, r2)
          ExitThread r2
          Exit Function
        End If
        CloseHandle Thread
        'OK
        Result = True
      End If
    End If
    FreeLibrary Lib
  End If
  
RegError:
  RegisterFile = Result
  Exit Function

End Function

Public Function urlencode(str As String) As String
On Error Resume Next
urlencode = ConvertString(str)
End Function

Public Sub ChangePriority(dwPriorityClass As Long)
On Error Resume Next
    Dim hProcess&
    Dim Ret&, pid&
    pid = GetCurrentProcessId() ' get my proccess id
    ' get a handle to the process
    hProcess = OpenProcess(PROCESS_DUP_HANDLE, True, pid)


    If hProcess = 0 Then
        Err.Raise 2, "ChangePriority", "Unable To open the source process"
        Exit Sub
    End If
    ' change the priority
    Ret = SetPriorityClass(hProcess, dwPriorityClass)
    ' Close the source process handle
    Call CloseHandle(hProcess)


    If Ret = 0 Then
        Err.Raise 4, "ChangePriority", "Unable To close source handle"
        Exit Sub
    End If
End Sub

Public Sub SetPP(p As PP)
On Error Resume Next
ChangePriority p
End Sub

Private Function hex2(v As Byte) As String
  hex2 = Right$("00" & Hex$(v), 2)
End Function

Private Function FmtFileTime(ft As FILETIMEREC) As String
  Dim st As SYSTEMTIMEREC
    
  If FileTimeToSystemTime(ft, st) <> 0 Then
    FmtFileTime = Format$(st.wYear, "0000") _
                & "-" _
                & Format$(st.wMonth, "00") _
                & "-" _
                & Format$(st.wDay, "00") _
                & " " _
                & Format$(st.wHour, "00") _
                & ":" _
                & Format$(st.wMinute, "00") _
                & ":" _
                & Format$(st.wSecond, "00") _
                & "." _
                & Format$(st.wMilliseconds, "000")
  Else
   FmtFileTime = "?"
  End If
End Function

Public Function GetFileTimeStamps(FVI As FILEVERSIONINFO) As Integer
  Dim hFile As Integer
  Dim FileStruct As OFSTRUCTREC
  Dim CreationTime As FILETIMEREC
  Dim LastAccessTime As FILETIMEREC
  Dim LastWriteTime As FILETIMEREC
  Dim rc As Integer
  
  ' Open it to get a stream handle
  hFile = OpenFile(FVI.Path, FileStruct, OF_READ Or OF_SHARE_DENY_NONE)
  If hFile <> 0 Then
    If GetFileTime(hFile, CreationTime, LastAccessTime, LastWriteTime) Then
      FVI.FileCreated = FmtFileTime(CreationTime)
      FVI.FileLastRead = FmtFileTime(LastAccessTime)
      FVI.FileLastWritten = FmtFileTime(LastWriteTime)
    End If
    rc = lclose(hFile)
  Else
    rc = 0
  End If
  GetFileTimeStamps = rc
End Function

Public Function SetFileTimeStamps(FVI As FILEVERSIONINFO, CREATED As SYSTEMTIME, LASTACCESS As SYSTEMTIME, LASTWRITE As SYSTEMTIME) As Integer
 ' by mel@melaxis.de - vorlage war GetFileTimeStamps
  
  Dim hFile As Integer
  Dim FileStruct As OFSTRUCTREC
  Dim CreationTime As FileTime
  Dim LastAccessTime As FileTime
  Dim LastWriteTime As FileTime
  Dim ft As FILETIMEREC
  Dim st As SYSTEMTIMEREC
  Dim rc As Integer
  
  ' Open it to get a stream handle
  hFile = OpenFile(FVI.Path, FileStruct, OF_READWRITE Or OF_SHARE_DENY_NONE)
  If hFile <> 0 Then
    Call SystemTimeToFileTime(CREATED, CreationTime)
    Call SystemTimeToFileTime(LASTACCESS, LastAccessTime)
    Call SystemTimeToFileTime(LASTWRITE, LastWriteTime)
    Call SetFileTime(hFile, CreationTime, LastAccessTime, LastWriteTime)
    rc = lclose(hFile)
  Else
    rc = 0
  End If
  SetFileTimeStamps = rc
  
End Function

Private Function FmtOSFlags(OSFlags As Long) As String
  If OSFlags = 0 Then
    FmtOSFlags = "?"
  Else
    Select Case OSFlags
      Case VOS__WINDOWS16:    FmtOSFlags = "WIN16"
      Case VOS__PM16:         FmtOSFlags = "PM16"
      Case VOS__PM32:         FmtOSFlags = "PM32"
      Case VOS__WINDOWS32:    FmtOSFlags = "WIN32"
      Case VOS_DOS:           FmtOSFlags = "DOS"
      Case VOS_DOS_WINDOWS16: FmtOSFlags = "DOS16"
      Case VOS_DOS_WINDOWS32: FmtOSFlags = "DOS32"
      Case VOS_OS216:         FmtOSFlags = "OS216"
      Case VOS_OS216_PM16:    FmtOSFlags = "OS2PM16"
      Case VOS_OS232:         FmtOSFlags = "OS232"
      Case VOS_OS232_PM32:    FmtOSFlags = "OS2PM32"
      Case VOS_NT:            FmtOSFlags = "NT"
      Case VOS_NT_WINDOWS32:  FmtOSFlags = "NTW32"
      Case Else:              FmtOSFlags = "OTHER"
    End Select
  End If
End Function

Private Function FmtBinFlags(BinFlags As Long) As String
  BinFlags = BinFlags And VS_FFI_FILEFLAGSMASK
  FmtBinFlags = ""
  If BinFlags And VS_FF_DEBUG Then
    FmtBinFlags = FmtBinFlags & "D"
  ElseIf BinFlags And VS_FF_PRERELEASE = VS_FF_PRERELEASE Then
    FmtBinFlags = FmtBinFlags & "b"
  ElseIf BinFlags And VS_FF_PATCHED = VS_FF_PATCHED Then
    FmtBinFlags = FmtBinFlags & "p"
  ElseIf BinFlags And VS_FF_PRIVATEBUILD = VS_FF_PRIVATEBUILD Then
    FmtBinFlags = FmtBinFlags & "P"
  ElseIf BinFlags And VS_FF_INFOINFERRED = VS_FF_INFOINFERRED Then
    FmtBinFlags = FmtBinFlags & "I"
  ElseIf BinFlags And VS_FF_SPECIALBUILD = VS_FF_SPECIALBUILD Then
    FmtBinFlags = FmtBinFlags & "S"
  End If
End Function


Private Function FmtVersion(MSLong As Long, LSLong As Long) As String
  FmtVersion = Format$((MSLong And &HFFFF0000) \ 65536, "####0.") _
             & Format$((MSLong And &HFFFF&), "###00.") _
             & Format$((LSLong And &HFFFF0000) \ 65536, "###00.") _
             & Format$((LSLong And &HFFFF&), "#0000")
End Function

'We will get the complete version info into FILEVERSIONINFO
'a return code of <=0 implies we had a failure, >0 all OK

Public Function GetFVInfo(ByRef FVI As FILEVERSIONINFO) As Long
  
  Dim q As Long, i As Long, vptr As Long, vlen As Long, vsffi As VS_FIXEDFILEINFO
  Dim InfoSize As Long, Info() As Byte, wsp(0 To 15) As Byte, buf As String
  Dim SubBlock As String, Lang_Charset As String, VersionInfo(0 To 7) As String
  
  FVI.Filesize = "?"
  FVI.BinState = "?"
  FVI.OSType = "?"
  FVI.CompanyName = ""
  FVI.FileDescription = ""
  FVI.FileVersion = "?.?.?.?"
  FVI.FileCreated = "?"
  FVI.InternalName = ""
  FVI.LegalCopyright = ""
  FVI.OriginalFileName = ""
  FVI.ProductName = ""
  FVI.ProductVersion = "?.?.?.?"
  
  On Error Resume Next                                   'get the filesize
  FVI.Filesize = Format$(FileLen(FVI.Path), "###,###,###")
  On Error GoTo 0
  
  Call GetFileTimeStamps(FVI)                            'get file timestamps :-)
  
  'now get the other version information
  
  InfoSize = GetFileVersionInfoSize(FVI.Path, q)
  If InfoSize <= 0 Then
    GetFVInfo = -1      'version not available
    Exit Function
  End If
  
  ReDim Info(0 To InfoSize) As Byte
  If GetFileVersionInfo(FVI.Path, q, InfoSize, Info(0)) <= 0 Then
    GetFVInfo = -2      'read versioninfo failed
    Exit Function
  End If
  
  SubBlock = "\"    'Root for FixedFileInfo
  If VerQueryValue(Info(0), SubBlock, vptr, vlen) > 0 Then
    If vlen > 0 Then
      Call CopyMemory(vsffi, vptr, vlen)
      FVI.FileVersion = FmtVersion(vsffi.dwFileVersionMS, vsffi.dwFileVersionLS)
      FVI.ProductVersion = FmtVersion(vsffi.dwProductVersionMS, vsffi.dwProductVersionLS)
      FVI.OSType = FmtOSFlags(vsffi.dwFileOS)
      FVI.BinState = FmtBinFlags(vsffi.dwFileFlags)
    End If
  End If
  
  SubBlock = "\VarFileInfo\Translation"
  If VerQueryValue(Info(0), SubBlock, vptr, vlen) <= 0 Then
    Lang_Charset = "040904E4"       'read translation key failed so assume MS default
  Else
    'vptr is a pointer to four 4 bytes of Hex number,
    'first two bytes are language id, and last two bytes are code page.
    'However, VerQueryValue needs a  string of 4 hex digits,
    ' the first two characters correspond to the language id and
    ' the last two characters correspond to the code page id.
  
    'now we change the order of the language id and code page
    'and convert it into a string representation.
    'For example, it may look like 040904E4
    'Or to pull it all apart:
    '04------        = SUBLANG_ENGLISH_USA
    '--09----        = LANG_ENGLISH
    ' ----04E4 = 1252 = Codepage for Windows:Multilingual
  
    Call CopyMemory(wsp(0), vptr, vlen)
    Lang_Charset = hex2(wsp(1)) & hex2(wsp(0)) & hex2(wsp(3)) & hex2(wsp(2))
  End If
  
  VersionInfo(0) = "CompanyName"
  VersionInfo(1) = "FileDescription"
  VersionInfo(2) = "FileVersion"
  VersionInfo(3) = "InternalName"
  VersionInfo(4) = "LegalCopyright"
  VersionInfo(5) = "OriginalFileName"
  VersionInfo(6) = "ProductName"
  VersionInfo(7) = "ProductVersion"

  For i = 0 To 7
    buf = String(255, 0)
    SubBlock = "\StringFileInfo\" & Lang_Charset & "\" & VersionInfo(i)
    
    If VerQueryValue(Info(0), SubBlock, vptr, vlen) = 0 Then
      GetFVInfo = -3      'read subblock key failed
      Exit Function
    End If
    If vlen > 0 Then
      Call lstrcpy(buf, vptr)
      VersionInfo(i) = Mid$(buf, 1, InStr(buf, Chr$(0)) - 1)
    Else
      VersionInfo(i) = ""
    End If
  Next i
  
  'fill the outbound array  - we dont actually use all of these in the main program
  FVI.CompanyName = VersionInfo(0)
  FVI.FileDescription = VersionInfo(1)
  If FVI.FileVersion < VersionInfo(2) Then FVI.FileVersion = VersionInfo(2)
  FVI.InternalName = VersionInfo(3)
  FVI.LegalCopyright = VersionInfo(4)
  FVI.OriginalFileName = VersionInfo(5)
  FVI.ProductName = VersionInfo(6)
  If FVI.ProductVersion < VersionInfo(7) Then FVI.ProductVersion = VersionInfo(7)
  GetFVInfo = 1

End Function

Public Function HuffmanEncode(Text As String, Optional Force As Boolean) As String
On Error Resume Next
    Dim TextLen As Long, Char As Byte, i As Long, j As Long
    Dim CodeCounts(255) As Long, BitStrings(255), BitString
    Dim HuffmanTrees As Collection
    Dim HTRootNode As Collection, HTNode As Collection
    Dim NextByte As Byte, BitPos As Integer, Temp As String
    
    'Initialize for processing.
    TextLen = Len(Text)
    Set HuffmanTrees = New Collection
    
    'Is there anything to encode?
    If TextLen = 0 Then
        HuffmanEncode = "HE0" & vbCr  'Version 0 = Plain text
        Exit Function  'No point in continuing
    End If
    
    HuffmanEncode = "HE2" & vbCr  'Version 1
    
    'Count how many times each ASCII code is encountered in text.
    For i = 1 To TextLen
        Char = Asc(Mid(Text, i, 1))
        CodeCounts(Char) = CodeCounts(Char) + 1
    Next
    
    'Initialize the forest of Huffman trees; one for each ASCII
    'character used.
    For i = 0 To UBound(CodeCounts)
        If CodeCounts(i) > 0 Then
            Set HTNode = NewNode
            s HTNode, htnAsciiCode, Chr(i)
            s HTNode, htnWeight, CDbl(CodeCounts(i) / TextLen)
            s HTNode, htnIsLeaf, True
            
            'Now place it in its reverse-ordered position.
            For j = 1 To HuffmanTrees.Count + 1
                If j > HuffmanTrees.Count Then
                    HuffmanTrees.Add HTNode
                    Exit For
                End If
                If HTNode(htnWeight) >= HuffmanTrees(j)(htnWeight) Then
                    HuffmanTrees.Add HTNode, , j
                    Exit For
                End If
            Next
        End If
    Next
    
    'Now assemble all these single-level Huffman trees into
    'one single tree, where all the leaves have the ASCII codes
    'associated with them.
    If HuffmanTrees.Count = 1 Then
        Set HTNode = NewNode
        s HTNode, htnLeftSubtree, HuffmanTrees(1)
        s HTNode, htnWeight, 1
        HuffmanTrees.Remove (1)
        HuffmanTrees.Add HTNode
    End If
    While HuffmanTrees.Count > 1
        Set HTNode = NewNode
        s HTNode, htnRightSubtree, HuffmanTrees(HuffmanTrees.Count)
        HuffmanTrees.Remove HuffmanTrees.Count
        s HTNode, htnLeftSubtree, HuffmanTrees(HuffmanTrees.Count)
        HuffmanTrees.Remove HuffmanTrees.Count
        s HTNode, htnWeight, HTNode(htnLeftSubtree)(htnWeight) + HTNode(htnRightSubtree)(htnWeight)
        
        'Place this new tree it in its reverse-ordered position.
        For j = 1 To HuffmanTrees.Count + 1
            If j > HuffmanTrees.Count Then
                HuffmanTrees.Add HTNode
                Exit For
            End If
            If HTNode(htnWeight) >= HuffmanTrees(j)(htnWeight) Then
                HuffmanTrees.Add HTNode, , j
                Exit For
            End If
        Next
    Wend
    Set HTRootNode = HuffmanTrees(1)
    AttachBitCodes BitStrings, HTRootNode, Array()
    For i = 0 To UBound(BitStrings)
        If Not IsEmpty(BitStrings(i)) Then
            Set HTNode = BitStrings(i)
            Temp = Temp & HTNode(htnAsciiCode) _
              & BitsToString(HTNode(htnBitCode))
        End If
    Next
    HuffmanEncode = HuffmanEncode & Len(Temp) & vbCr & Temp
    
    'The next part of the header is a checksum value, which
    'we'll use later to verify our decompression.
    Char = 0
    For i = 1 To TextLen
        Char = Char Xor Asc(Mid(Text, i, 1))
    Next
    HuffmanEncode = HuffmanEncode & Chr(Char)
    
    'The final part of the header identifies how many bytes
    'the original text strings contains.  We will probably
    'have a few unused bits in the last byte that we need to
    'account for.  Additionally, this serves as a final check
    'for corruption.
    HuffmanEncode = HuffmanEncode & TextLen & vbCr
    
    'Now we can encode the data by exchanging each ASCII byte for
    'its appropriate bit string.
    BitPos = -1
    Char = 0
    Temp = ""
    For i = 1 To TextLen
        BitString = BitStrings(Asc(Mid(Text, i, 1)))(htnBitCode)
        'Add each bit to the end of the output stream's 1-byte buffer.
        For j = 0 To UBound(BitString)
            BitPos = BitPos + 1
            If BitString(j) = 1 Then
                Char = Char + 2 ^ BitPos
            End If
            'If the bit buffer is full, dump it to the output stream.
            If BitPos >= 7 Then
                Temp = Temp & Chr(Char)
                'If the temporary output buffer is full, dump it
                'to the final output stream.
                If Len(Temp) > 1024 Then
                    HuffmanEncode = HuffmanEncode & Temp
                    Temp = ""
                End If
                BitPos = -1
                Char = 0
            End If
        Next
    Next
    If BitPos > -1 Then
        Temp = Temp & Chr(Char)
    End If
    If Len(Temp) > 0 Then
        HuffmanEncode = HuffmanEncode & Temp
    End If
    
    'If it takes up more space compressed because the source is
    'small and the header is big, we'll leave it uncompressed
    'and prepend it with a 4 byte header.
    If Len(HuffmanEncode) > TextLen And Not Force Then
        HuffmanEncode = "HE0" & vbCr & Text
    End If
End Function


'Decompress the string back into its original text.
Public Function HuffmanDecode(ByVal Text As String) As String
On Error Resume Next
    Dim pos As Long, Temp As String, Char As Byte, Bits
    Dim i As Long, j As Long, CharsFound As Long, BitPos As Integer
    Dim CheckSum As Byte, SourceLen As Long, TextLen As Long
    Dim HTRootNode As Collection, HTNode As Collection
    
    'If this was left uncompressed, this will be easy.
    If Left(Text, 4) = "HE0" & vbCr Then
        HuffmanDecode = Mid(Text, 5)
        Exit Function
    End If
    
    'If this is any version other than 2, we'll bow out.
    If Left(Text, 4) <> "HE2" & vbCr Then
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "The data either was not compressed with HE2 or is corrupt"
    End If
    Text = Mid(Text, 5)
    
    'Extract the ASCII character bit-code table's byte length.
    pos = InStr(1, Text, vbCr)
    If pos = 0 Then
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "The data either was not compressed with HE2 or is corrupt"
    End If
    On Error Resume Next
    TextLen = Left(Text, pos - 1)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "The header is corrupt"
    End If
    On Error GoTo 0
    Text = Mid(Text, pos + 1)
    Temp = Left(Text, TextLen)
    Text = Mid(Text, TextLen + 1)
    'Now extract the ASCII character bit-code table.
    Set HTRootNode = NewNode
    pos = 1
    While pos <= Len(Temp)
        Char = Asc(Mid(Temp, pos, 1))
        pos = pos + 1
        Bits = StringToBits(pos, Temp)
        Set HTNode = HTRootNode
        For j = 0 To UBound(Bits)
            If Bits(j) = 1 Then
                If HTNode(htnLeftSubtree) Is Nothing Then
                    s HTNode, htnLeftSubtree, NewNode
                End If
                Set HTNode = HTNode(htnLeftSubtree)
            Else
                If HTNode(htnRightSubtree) Is Nothing Then
                    s HTNode, htnRightSubtree, NewNode
                End If
                Set HTNode = HTNode(htnRightSubtree)
            End If
        Next
        s HTNode, htnIsLeaf, True
        s HTNode, htnAsciiCode, Chr(Char)
        s HTNode, htnBitCode, Bits
    Wend
    
    'Extract the checksum.
    CheckSum = Asc(Left(Text, 1))
    Text = Mid(Text, 2)
    
    'Extract the length of the original string.
    pos = InStr(1, Text, vbCr)
    If pos = 0 Then
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "The header is corrupt"
    End If
    On Error Resume Next
    SourceLen = Left(Text, pos - 1)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "The header is corrupt"
    End If
    On Error GoTo 0
    Text = Mid(Text, pos + 1)
    TextLen = Len(Text)
    
    'Now that we've processed the header, let's decode the actual data.
    i = 1
    BitPos = -1
    Set HTNode = HTRootNode
    Temp = ""
    While CharsFound < SourceLen
        If BitPos = -1 Then
            If i > TextLen Then
                Err.Raise vbObjectError, "HuffmanDecode()", _
                  "Expecting more bytes in data stream"
            End If
            Char = Asc(Mid(Text, i, 1))
            i = i + 1
        End If
        BitPos = BitPos + 1
        
        If (Char And 2 ^ BitPos) > 0 Then
            Set HTNode = HTNode(htnLeftSubtree)
        Else
            Set HTNode = HTNode(htnRightSubtree)
        End If
        If HTNode Is Nothing Then
            'Uh oh.  We've followed the tree to a Huffman tree to a dead
            'end, which won't happen unless the data is corrupt.
            Err.Raise vbObjectError, "HuffmanDecode()", _
              "The header (lookup table) is corrupt"
        End If
        
        If HTNode(htnIsLeaf) Then
            Temp = Temp & HTNode(htnAsciiCode)
            If Len(Temp) > 1024 Then
                HuffmanDecode = HuffmanDecode & Temp
                Temp = ""
            End If
            CharsFound = CharsFound + 1
            Set HTNode = HTRootNode
        End If
        
        If BitPos >= 7 Then BitPos = -1
    Wend
    If Len(Temp) > 0 Then
        HuffmanDecode = HuffmanDecode & Temp
    End If
    If i <= TextLen Then
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "Found extra bytes at end of data stream"
    End If
    
    'Verify data to check for corruption.
    If Len(HuffmanDecode) <> SourceLen Then
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "Data corrupt because check sums do not match"
    End If
    Char = 0
    For i = 1 To SourceLen
        Char = Char Xor Asc(Mid(HuffmanDecode, i, 1))
    Next
    If Char <> CheckSum Then
        Err.Raise vbObjectError, "HuffmanDecode()", _
          "Data corrupt because check sums do not match"
    End If
End Function



'----------------------------------------------------------------
' Everything below here is only for supporting the two main
' routines above.
'----------------------------------------------------------------


'Follows the tree, now built, to its end leaf nodes, where the
'character codes are, in order to tell those character codes
'what their bit string representations are.
Private Sub AttachBitCodes(BitStrings, HTNode As Collection, ByVal Bits)
On Error Resume Next
    If HTNode Is Nothing Then Exit Sub
    If HTNode(htnIsLeaf) Then
        s HTNode, htnBitCode, Bits
        Set BitStrings(Asc(HTNode(htnAsciiCode))) = HTNode
    Else
        ReDim Preserve Bits(UBound(Bits) + 1)
        Bits(UBound(Bits)) = 1
        AttachBitCodes BitStrings, HTNode(htnLeftSubtree), Bits
        Bits(UBound(Bits)) = 0
        AttachBitCodes BitStrings, HTNode(htnRightSubtree), Bits
    End If
End Sub

'Turns a string of '0' and '1' characters into a string of bytes
'containing the bits, preceeded by 1 byte indicating the
'number of bits represented.
Private Function BitsToString(Bits) As String
On Error Resume Next
    Dim Char As Byte, i As Long
    BitsToString = Chr(UBound(Bits) + 1)  'Number of bits
    For i = 0 To UBound(Bits)
        If i Mod 8 = 0 Then
            If i > 0 Then BitsToString = BitsToString & Chr(Char)
            Char = 0
        End If
        If Bits(i) = 1 Then  'Bit value = 1
            'Mask the bit into its proper position in the byte
            Char = Char + 2 ^ (i Mod 8)
        End If
    Next
    BitsToString = BitsToString & Chr(Char)
End Function

'The opposite of BitsToString() function.
Private Function StringToBits(StartPos As Long, Bytes As String)
On Error Resume Next
    Dim Char As Byte, i As Long, BitCount As Long, Bits
    Bits = Array()
    BitCount = Asc(Mid(Bytes, StartPos, 1))
    StartPos = StartPos + 1
    For i = 0 To BitCount - 1
        If i Mod 8 = 0 Then
            Char = Asc(Mid(Bytes, StartPos, 1))
            StartPos = StartPos + 1
        End If
        ReDim Preserve Bits(UBound(Bits) + 1)
        If (Char And 2 ^ (i Mod 8)) > 0 Then   'Bit value = 1
            Bits(UBound(Bits)) = 1
        Else  'Bit value = 0
            Bits(UBound(Bits)) = 0
        End If
    Next
    StringToBits = Bits
End Function

'Remove the specified item and put the specified value in its place.
Private Sub s(COL As Collection, Index As HuffmanTreeNodeParts, Value)
On Error Resume Next
    COL.Remove Index
    If Index > COL.Count Then
        COL.Add Value
    Else
        COL.Add Value, , Index
    End If
End Sub

'Creates a new Huffman tree node with the default values set.
Private Function NewNode() As Collection
On Error Resume Next
    Dim Node As New Collection
    Node.Add 0  'htnWeight
    Node.Add False  'htnIsLeaf
    Node.Add Chr(0)  'htnAsciiCode
    Node.Add ""  'htnBitCode
    Node.Add Nothing  'htnLeftSubtree
    Node.Add Nothing  'htnRightSubtree
    Set NewNode = Node
End Function

Public Function DriveSerialNumber(ByVal Drive As String) As Long
On Error Resume Next
    'usage: SN = DriveSerialNumber("C:\")
 
    Dim lAns As Long
    Dim lret As Long
    Dim sVolumeName As String, sDriveType As String
    Dim sDrive As String

    'Deal with one and two character input values
    sDrive = Drive
    If Len(sDrive) = 1 Then
        sDrive = sDrive & ":\"
    ElseIf Len(sDrive) = 2 And Right(sDrive, 1) = ":" Then
        sDrive = sDrive & "\"
    End If
  
    sVolumeName = String$(255, Chr$(0))
    sDriveType = String$(255, Chr$(0))
    
    lret = GetVolumeInformation(sDrive, sVolumeName, _
    255, lAns, 0, 0, sDriveType, 255)

DriveSerialNumber = lAns
End Function

Public Function Threads() As Long
    On Error GoTo errorhandler:
    Dim lResult As Long
    Dim lData As Long
    Dim lType As Long
    Dim hKey As Long
   
    
    lResult = RegOpenKeyEx(HKEY_DYN_DATA, STAT_DATA, _
       0, KEY_QUERY_VALUE, hKey)
    
    If lResult = 0 Then
        lResult = RegQueryValueEx(hKey, NUM_THREADS, 0, _
           lType, lData, 4)
        If lResult = 0 Then
            Threads = lData
            lResult = RegCloseKey(hKey)
        End If
    End If
    Exit Function

errorhandler:
    On Error Resume Next
    RegCloseKey hKey
    Exit Function
End Function

Public Function CheckConnection1() As Boolean
'On Error Resume Next
Dim RETURNCODE As Long
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess" & Chr$(0)
RETURNCODE = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)
If RETURNCODE = ERROR_SUCCESS Then
   hKey = phkResult
   lpValueName = "Remote Connection"
   lpReserved = APINULL
   lpType = APINULL
   lpData = APINULL
   lpcbData = APINULL
   RETURNCODE = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
   lpcbData = Len(lpData)
   RETURNCODE = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData)
   If RETURNCODE = ERROR_SUCCESS Then
       If lpData = 0 Then
          CheckConnection1 = False
       Else
          CheckConnection1 = True
       End If
   Else
       CheckConnection1 = True ' lan
   End If
End If
RegCloseKey (hKey)
End Function

Public Function CheckConnection2(Optional ByRef ConnectionInfo As Long, Optional ByRef sConnectionName As String) As Boolean
Dim dwFlags As Long
Dim sNameBuf As String, msg As String
Dim lPos As Long
sNameBuf = String$(513, 0)
If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&) Then
   lPos = InStr(sNameBuf, vbNullChar)
   If lPos > 0 Then
     sConnectionName = Left$(sNameBuf, lPos - 1)
   Else
     sConnectionName = ""
   End If
   msg = "Your computer is connected to Internet" & vbCrLf & "Connection Name: " & sConnectionName
   If (dwFlags And INTERNET_CONNECTION_LAN) Then
       msg = msg & vbCrLf & "Connection use LAN"
   ElseIf lFlags And INTERNET_CONNECTION_MODEM Then
       msg = msg & vbCrLf & "Connection use modem"
   End If
   If lFlags And INTERNET_CONNECTION_PROXY Then msg = msg & vbCrLf & "Connection use Proxy"
   If lFlags And INTERNET_RAS_INSTALLED Then
      msg = msg & vbCrLf & "RAS INSTALLED"
   Else
      msg = msg & vbCrLf & "RAS NOT INSTALLED"
   End If
   If lFlags And INTERNET_CONNECTION_OFFLINE Then
      CheckConnection2 = False
      msg = msg & vbCrLf & "You are OFFLINE"
   Else
      CheckConnection2 = True
      msg = msg & vbCrLf & "You are ONLINE"
   End If
   If lFlags And INTERNET_CONNECTION_CONFIGURED Then
      msg = msg & vbCrLf & "Your connection is Configured"
   Else
      msg = msg & vbCrLf & "Your connection is not Configured"
   End If
Else
   CheckConnection2 = False
   msg = "Your computer is NOT connected to Internet"
End If
   'MsgBox msg, vbInformation, "Checking connection"
End Function

Public Function Utf2Asc(DC1) As String
' Diese Funktion stammt von makasy.de
dx1 = 1
1 dcx = InStr(dx1, DC1, Chr(195))
If dcx <> 0 Then
    dx1 = dx1 + 1
    dcb1 = Left(DC1, dcx - 1)
    dcb2 = Chr(Asc(Mid(DC1, dcx + 1, 1)) + 64)
    dcb3 = Right(DC1, Len(DC1) - Len(dcb1) - 2)
    DC1 = dcb1 & dcb2 & dcb3
    GoTo 1
End If
Utf2Asc = DC1
End Function

Public Function Asc2Utf(DC1) As String
' Diese Funktion stammt von makasy.de
On Error GoTo 1
Do Until dx1 >= Len(DC1)
    dx1 = dx1 + 1
    dd1 = Mid(DC1, dx1, 1)
    If Asc(dd1) >= 128 Then
        dcb1 = Left(DC1, dx1 - 1)
        dcb2 = Chr(195) & Chr(Asc(dd1) - 64)
        dcb3 = Right(DC1, Len(DC1) - Len(dcb1) - 1)
        DC1 = dcb1 & dcb2 & dcb3
        dx1 = dx1 + 1
    End If
Loop
1 Asc2Utf = DC1
End Function

Public Function Hex2Dec(HexWert As String) As Long
On Error Resume Next
Dim erg As String
erg = "&H" & HexWert
Hex2Dec = Val(erg)
End Function

Public Function LoescheAlteDaten(Verzeichnis As String, Tage As Integer)
On Error Resume Next
Dim TagX As Date
Dim fs, f, s, ordner, Datei, erstellungsdatum

TagX = Format(Now() - Tage, "dd.mm.yyyy hh:nn")

Set fs = CreateObject("Scripting.FileSystemObject")
Set ordner = fs.getfolder(Verzeichnis)

For Each Datei In ordner.Files
 erstellungsdatum = Datei.DateCreated
 If erstellungsdatum < TagX Then Kill Datei
Next

End Function

Public Sub InPapierkorb(Datei As String)
On Error Resume Next
Dim SHFileOp As SHFILEOPSTRUCT
Dim l As Long

Datei = Datei & Chr$(0)

With SHFileOp
.wFunc = FO_DELETE
.pFrom = Datei
.fFlags = FOF_ALLOWUNDO
End With

l = SHFileOperation(SHFileOp)
End Sub

Public Sub DLFile(Datei As String)
On Error Resume Next
Dim URL As String
Dim l As Long
URL = StrConv(Datei, vbUnicode)
l = DoFileDownload(URL)
End Sub

Public Function stringVonBis(str As String, Von As String, Bis As String) As String
On Error Resume Next
' by mel@melaxis.de
Dim s As String, pos As String, t As String, Ret As String
Ret = ""
For i = 1 To Len(str)
    s = Mid(str, i)
    Debug.Print "s = " & s
    If isStart(s, Von) Then
        t = cutStart(s, Von)
        For c = 1 To Len(t)
            Debug.Print "c = " & c
            If Mid(t, c, Len(Bis)) <> Bis Then
                Ret = Ret & Mid(t, c, 1)
            Else
                Exit For
            End If
        Next
    End If
Next
stringVonBis = Ret
End Function

Public Function IP2LongIP(IP As String) As String
On Error Resume Next
' by mel@melaxis.de
Dim n As Long, fip As String
n = 256
sIP = Split(IP, ".")
fip = (sIP(0) * (n * n * n)) + (sIP(1) * (n * n)) + (sIP(2) * n) + sIP(3)
IP2LongIP = fip
End Function

Function GetLocalTZ(Optional ByRef strTZName As String) As Long
    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    lngResult = GetTimeZoneInformation&(objTimeZone)


    Select Case lngResult
        Case 0&, 1& 'use standard time
        GetLocalTZ = -(objTimeZone.Bias + objTimeZone.StandardBias) * 60 'into minutes


        For i = 0 To 31
            If objTimeZone.StandardName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.StandardName(i))
        Next
        Case 2& 'use daylight savings time
        GetLocalTZ = -(objTimeZone.Bias + objTimeZone.DaylightBias) * 60 'into minutes


        For i = 0 To 31
            If objTimeZone.DaylightName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
        Next
    End Select
End Function

Function GetTime() As Double
Dim TheDate As Date 'target date
Dim iResult As Long
    'set target date.
    TheDate = "01/01/1970"
    'compute # of seconds left to target date
    SecondsToTarget = DateDiff("s", Now, TheDate)
    iResult = 36 * Mid$((GetLocalTZ + 23200), 2)
    
    ' 3600
    GetTime = Mid(SecondsToTarget, 2) + iResult
End Function

Public Function sUnixDate(ByVal lValue As Long) As String
    ' Now for the LocalTime function. Take
    '     the long value representing the number
    ' of seconds since January 1, 1970 and c
    '     reate a useable time structure from it.
    ' Return a formatted string YYYY/MM/DD H
    '     H:MM:SS
    Dim lSecPerYear
    Dim Year As Long
    Dim Month As Long
    Dim Day As Long
    Dim Hour As Long
    Dim Minute As Long
    Dim Second As Long
    Dim Temp As Long
    ' [0] = normal year, [1] = leap year
    lSecPerYear = Array(31536000, 31622400)
    lSecPerDay = 86400 ' 60*60*24
    lSecPerHour = 3600 ' 60 * 60
    Year = 70 ' starting point
    ' Calculate the year


    Do While (lValue > 0)
        Temp = isLeapYear(Year)


        If (lValue - lSecPerYear(Temp)) > 0 Then
            lValue = lValue - lSecPerYear(Temp)
            Year = Year + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Year = " & Year
    ' Calculate the month
    Month = 1


    Do While (lValue > 0)
        Temp = secsInMonth(Year, Month)


        If (lValue - Temp) > 0 Then
            lValue = lValue - Temp
            Month = Month + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Month = " & Month
    ' Now calculate day
    Day = 1


    Do While (lValue > 0)


        If (lValue - lSecPerDay) > 0 Then
            lValue = lValue - lSecPerDay
            Day = Day + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Day = " & Day
    ' Now calculate Hour
    Hour = 0


    Do While (lValue > 0)


        If (lValue - lSecPerHour) > 0 Then
            lValue = lValue - lSecPerHour
            Hour = Hour + 1
        Else
            Exit Do
        End If
    Loop
    'Debug.Print "Hour = " & Hour
    Minute = Int(lValue / 60)
    'Debug.Print "Minute = " & Minute
    Second = lValue Mod 60
    'Debug.Print "Second = " & Second
    ' Standard date format is YYYY/MM/DD HH:
    '     MM:SS
    'If Year < 100 Then
    Year = Year + 1900
    'sUnixDate = Month & "/" & Day & "/" & Year & ", " & Hour & ":" & Minute & ":" & Second
    sUnixDate = Day & "." & Month & "." & Year & " " & Hour & ":" & Minute & ":" & Second
End Function

Public Sub CreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant)

    Dim hHnd As Long
    
    If Not IsMissing(SubKey) Then
        Temp = RegCreateKey(hKey, Key & "\" & SubKey, hHnd)
        Temp = RegCloseKey(hHnd)
    Else
        Temp = RegCreateKey(hKey, Key, hHnd)
        Temp = RegCloseKey(hHnd)
    End If

End Sub

Public Function GetString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As String

    Dim hHnd As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lValueType As Long
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufferSize, " ")
        lResult = RegQueryValueEx(hHnd, ValueName, 0&, 0&, ByVal strBuf, lDataBufferSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        
        End If
    End If
End Function

Public Sub SaveString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As String)

    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    Temp = RegSetValueEx(hHnd, ValueTitle, 0, REG_SZ, ByVal ValueData, Len(ValueData))
    Temp = RegCloseKey(hHnd)

End Sub



Public Function GetDWord(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As Long

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufferSize As Long
    Dim Temp As Long
    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lDataBufferSize = 4
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, lBuf, lDataBufferSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If

    Temp = RegCloseKey(hHnd)

End Function

Public Function GetBinary(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As Long

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufferSize As Long
    Dim Temp As Long
    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    lDataBufferSize = 4
    lResult = RegQueryValueEx(hHnd, ValueName, 0&, lValueType, lBuf, lDataBufferSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            GetBinary = lBuf
        End If
    End If

    Temp = RegCloseKey(hHnd)

End Function


Public Sub SaveDWord(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As Variant, Optional ByVal DataLength As Long = 4)

    Dim lResult As Long
    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    lResult = RegSetValueEx(hHnd, ValueTitle, 0&, REG_DWORD, ValueData, DataLength) ' 4
    Temp = RegCloseKey(hHnd)

End Sub
Public Sub SaveBinary(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueTitle As String, ByVal ValueData As Variant)

    Dim lResult As Long
    Dim hHnd As Long
    Dim Temp As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegCreateKey(hKey, KeyPath, hHnd)
    lResult = RegSetValueEx(hHnd, ValueTitle, 0&, REG_BINARY, ValueData, 4)
    Temp = RegCloseKey(hHnd)

End Sub




Public Sub DeleteKey(ByVal hKey As Long, ByVal Key As String)

    Dim Temp As Long
    
    Temp = RegDeleteKey(hKey, Key)

End Sub

Public Sub DeleteValue(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal Value As String)

    Dim hHnd As Long
    
    KeyPath = Key + "\" + SubKey
    Temp = RegOpenKey(hKey, KeyPath, hHnd)
    Temp = RegDeleteValue(hHnd, Value)
    Temp = RegCloseKey(hHnd)

End Sub

Public Function ZahlZwischen(ZahlMin As Long, ZahlMax As Long)
On Error Resume Next
' ©reated by mel @ 09/04/01
Randomize
ZahlZwischen = ((Timer + (Rnd * 100)) Mod (ZahlMax - ZahlMin + 1)) + ZahlMin
End Function

Public Function melWurzel(Zahl As Long) As Long
On Error Resume Next
' ©reated by mel @ unknown :-)
Dim x As Long, a As Long, y As Long, i As Integer, w As Boolean
a = Zahl
x = 1
w = True
Do While w = True
    y = x
    x = (y + a / y) * 0.5
    If x = y Then w = False
Loop
melWurzel = x
End Function

Public Function HTMLCrypt(ByVal Eingabe As String) As HTMLCryptedText
On Error Resume Next
' ©reated by mel @ 12/04/01
Dim mys As HTMLCryptedText
mys.Password = ";Qr8(9O=xSu>/AMRZPJw60t:kGiT" & Chr(34) & "bUnaW)c31XNI4mfhHC+-YFq2&yg!LlsKeB7d D.p5Vv_oEjz<"
mys.Number = 78
mys.Text = "noch nicht fertig"
End Function

Public Function HTMLDECrypt(CryptedText As HTMLCryptedText) As String
On Error Resume Next
' ©reated by mel @ 12/04/01
Dim l As String, p As Integer, i As String, a As String, k As String, m As Integer, v As String, w As Integer, q As Integer
' Variablen initialisieren
l = CryptedText.Text
i = CryptedText.Password
p = CryptedText.Number
' Schleife starten, um jedes Zeichen durchzugehen
For m = 1 To Len(l)
    v = Mid(l, m, 1)
    w = InStr(i, v)
    If w > 0 Then
        q = ((w Mod p - 1))
        If q <= 0 Then
            q = q + p
        End If
        k = k & Mid(i, q, 1)
    Else
        k = k & v
    End If
Next
a = a & k
HTMLDECrypt = a
End Function

'**************************************
' Name: System wide keyboard event
' Description:This windows hook is calle
'     d whenever a key is pressed which allows
'     you to install a system wide keyboard tr
'     ap - to simulate a hotkey, for example.
' By: Duncan Jones
'
' Side Effects:This hook only works on N
'     T4 and above.
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.22300/lngWId.1/qx/vb/scripts/ShowCode
'     .htm'for details.'**************************************

'\\ [VB_HOOKLOWLEVELKEYBOARDPROC]-------
'     ----------------------------------------
'     ----------------------------------------
'     -------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)
'     (int code, WPARAM wParam, LPARAM lParam)
'     ;
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the
'     calling code
'\\ ------------------------------------
'     ----------------------------------------
'     ----------------------------------------
'     --------------
'\\ You have a royalty free right to use
'     , reproduce, modify, publish and mess wi
'     th this code
'\\ I'd like you to visit http://www.mer
'     rioncomputing.com for updates, but won't


'     force you
    '\\ ------------------------------------
    '     ----------------------------------------
    '     ----------------------------------------
    '     --------------


Public Function VB_HOOKLOWLEVELKEYBOARDPROC(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Local Error Resume Next
    Dim Params() As Variant
    Dim lret As Long
    Dim lMsgRet As Long
    '\\ Note: If the code passed in is less
    '     than zero, it must be passed direct to t
    '     he next hook proc


    If code < 0 Then
        VB_HOOKLOWLEVELKEYBOARDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD_LL), code, wParam, lParam)
    End If
    '\\ 2 - Call the event firer....i.e. tes
    '     t for your key combo here
    '\\ 3 - Pass this message on to the next
    '     hook proc in the chain (if any)
    lret = CallNextHookEx(lHookId, code, wParam, lParam)
    '\\ If the message isn't cancelled, retu
    '     rn the next hook's message...


    If Not (lMsgRet) Then
        '\\ Return value to calling code....
        VB_HOOKLOWLEVELKEYBOARDPROC = lret
    End If
End Function
'\\ To start the hook:
'\\ If a hook of this type is already se
'     t, unhook this first
'
'
'If lHookId > 0 Then
'    Call UnhookWindowsHookEx(lHookId)
'End If
'lHookId = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf VB_HOOKKEYBOARDPROC, GetModuleHandle(App.EXEName), 0)
''\\ To end the hook:
'
'
'If lHookId > 0 Then
'    Call UnhookWindowsHookEx(lHookId)
'End If

'\\ --[CopyFromHandle]------------------
'     ---------
'\\ Copies the data from a global memory
'     handle
'\\ to a private byte array copy
'\\ ------------------------------------
'     ---------


Public Sub CopyFromHandle(ByVal hMemHandle As Long)
    Dim lret As Long
    Dim lPtr As Long
    lret = GlobalSize(hMemHandle)


    If lret > 0 Then
        mMyDataSize = lret
        lPtr = GlobalLock(hMemHandle)


        If lPtr > 0 Then
            ReDim mMyData(0 To mMyDataSize - 1) As Byte
            CopyMemory mMyData(0), ByVal lPtr, mMyDataSize
            Call GlobalUnlock(hMemHandle)
        End If
    End If
End Sub
'\\ --[CopyToHandle]--------------------
'     ---------
'\\ Copies the private data to a memory
'     handle
'\\ passed in
'\\ ------------------------------------
'     ---------


Public Sub CopyToHandle(ByVal hMemHandle As Long)
    Dim lSIZE As Long
    Dim lPtr As Long
    '\\ Don't copy if its empty


    If Not (mMyDataSize = 0) Then
        lSIZE = GlobalSize(hMemHandle)
        '\\ Don't attempt to copy if zero size..
        '     .


        If lSIZE > 0 Then


            If lPtr > 0 Then
                CopyMemory ByVal lPtr, mMyData(0), lSIZE
                Call GlobalUnlock(hMemHandle)
            End If
        End If
    End If
End Sub

'Author:    Dion Wiggins
'Purpose:   Creates a GUID
'Notes:
'Inputs:
'   - strRemoveChars    The characters to remove from the GUID (usually the {}- characters)
'History
'Date           Author          Description
'1 Jun 1999     Dion Wiggins    Created
Public Function CreateGUID( _
    Optional strRemoveChars As String = "{}-") As String
Dim udtGUID As GUID
Dim strGUID As String
Dim bytGUID() As Byte
Dim lngLen As Long
Dim lngRetVal As Long
Dim lngPos As Long

'Initialize
lngLen = 40
bytGUID = String(lngLen, 0)

'Create the GUID
CoCreateGuid udtGUID

'Convert the structure into a displayable string
lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
strGUID = bytGUID
If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
    lngRetVal = lngRetVal - 1
End If

'Trim the trailing characters
strGUID = Left$(strGUID, lngRetVal)

'Remove the unwanted characters
For lngPos = 1 To Len(strRemoveChars)
    strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), "")
Next

CreateGUID = strGUID
End Function

Function ConvertTime(TheTime As Single)
    Dim NewTime As String
    Dim Sec As Single
    Dim Min As Single
    Dim H As Single

    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If


    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function

Public Function Hex2HTMLColor(HexColor As String) As String
On Error Resume Next
Hex2HTMLColor = Mid(HexColor, 5, 2) & Mid(HexColor, 3, 2) & Mid(HexColor, 1, 2)
On Error GoTo 0
End Function

Public Function HTMLColor2Hex(HTMLColor As String) As String
On Error Resume Next
HTMLColor2Hex = Mid(HTMLColor, 5, 2) & Mid(HTMLColor, 3, 2) & Mid(HTMLColor, 1, 2)
On Error GoTo 0
End Function

Public Function UnHex(HexZahl As String) As Long
On Error Resume Next
UnHex = Val("&H" & HexZahl)
On Error GoTo 0
End Function

'<-- MEL'S HEX-TOOLS START -->
Public Function DecStelle2Hex(rein) As String
Select Case rein
Case 10:    raus = "A"
Case 11:    raus = "B"
Case 12:    raus = "C"
Case 13:    raus = "D"
Case 14:    raus = "E"
Case 15:    raus = "F"
Case Else:  raus = rein
End Select
DecStelle2Hex = raus
End Function

Public Function HexStelle2Dec(HexZahl As String) As Long
Dim raus As Long
Select Case HexZahl
Case "A":   raus = 10
Case "B":   raus = 11
Case "C":   raus = 12
Case "D":   raus = 13
Case "E":   raus = 14
Case "F":   raus = 15
Case Else:  raus = HexZahl
End Select
HexStelle2Dec = raus
End Function

Public Function melHex2Dec(HexZahl As String) As Long
Dim z As String, i As Integer, buf As Long, n As Integer, hoch As Integer, erg As Long
erg = 0
For i = 1 To Len(HexZahl)
    n = Len(HexZahl) - i + 1
    z = Mid(HexZahl, n, 1)
    hoch = i - 1
    buf = HexStelle2Dec(z) * (16 ^ hoch)
    erg = erg + buf
Next
melHex2Dec = erg
End Function

Public Function melDec2Hex(DecZahl As Long) As String
Dim was As Long, ganz As Long, rest As Long, erg As String
was = DecZahl
Do While was <> 0
    ganz = Int(was / 16)
    rest = was Mod 16
    erg = DecStelle2Hex(rest) & erg
    was = ganz
Loop
melDec2Hex = erg
End Function
'<-- MEL'S HEX-TOOLS ENDE -->

'Function melDec2HexSimpel(DecZahl As Long) As String
'' nur bis 255!
'Dim links As String, rechts As String
'links = Int(Text1 / 16)
'rechts = Text1 Mod 16
'links = DecStelle2Hex(links)
'rechts = DecStelle2Hex(rechts)
'melDec2HexSimpel = links & rechts
'End Function

Public Function IsNetConnectViaLAN() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the flags indicate a LAN
    '     connection
    IsNetConnectViaLAN = dwFlags And INTERNET_CONNECTION_LAN
End Function


Public Function IsNetConnectViaModem() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the flags indicate a mod
    '     em connection
    IsNetConnectViaModem = dwFlags And INTERNET_CONNECTION_MODEM
End Function


Public Function IsNetConnectViaProxy() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the flags indicate a pro
    '     xy connection
    IsNetConnectViaProxy = dwFlags And INTERNET_CONNECTION_PROXY
End Function


Public Function IsNetConnectOnline() As Boolean
    'no flags needed here - the API returns
    '     True
    'if there is a connection of any type
    IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function


Public Function IsNetRASInstalled() As Boolean
    Dim dwFlags As Long
    'pass an empty varialbe into which the A
    '     PI will
    'return the flags associated with the co
    '     nnection
    Call InternetGetConnectedState(dwFlags, 0&)
    'return True if the falgs include RAS in
    '     stalled
    IsNetRASInstalled = dwFlags And INTERNET_RAS_INSTALLED
End Function


Public Function GetNetConnectString() As String
    Dim dwFlags As Long
    Dim msg As String
    'build a string for display


    If InternetGetConnectedState(dwFlags, 0&) Then


        If dwFlags And INTERNET_CONNECTION_CONFIGURED Then
            msg = msg & "You have a network connection configured." & vbCrLf
        End If


        If dwFlags And INTERNET_CONNECTION_LAN Then
            msg = msg & "The local system connects To the Internet via a LAN"
        End If


        If dwFlags And INTERNET_CONNECTION_PROXY Then
            msg = msg & ", and uses a proxy server. "
        Else: msg = msg & "."
        End If


        If dwFlags And INTERNET_CONNECTION_MODEM Then
            msg = msg & "The local system uses a modem To connect to the Internet. "
        End If


        If dwFlags And INTERNET_CONNECTION_OFFLINE Then
            msg = msg & "The connection is currently offline. "
        End If


        If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then
            msg = msg & "The local system's modem is busy With a non-Internet connection. "
        End If


        If dwFlags And INTERNET_RAS_INSTALLED Then
            msg = msg & "Remote Access Services are installed On this system."
        End If
    Else
        msg = "Not connected To the internet now."
    End If
    GetNetConnectString = msg
End Function

' ********************************************************************************
' * lazyCrypt - (C) Copyright 2001 by Pablo Hoch - mel@melaxis.de - www.melaxis.de
' *
' * ERKLÄRUNG DES LAZYCRYPT-VERFAHRENS:
' *  lazyCrypt ist eine relativ einfache Verschlüsselung, die aber viele mögliche
' *  Kombinationen anbietet, die von einem Key und einer Basis abhängen. Nur wenn
' *  der zum Entschlüsseln verwendete Key und die Basis indentisch mit den Werten
' *  sind, die zum Verschlüsseln gewählt worden sind, wird der Text korrekt ent-
' *  schlüsslet. Je länger der Key ist, d.h. je mehr verschiedene Zeichen er ent-
' *  hält, desto sicherer ist die Verschlüsselung. Im Key nicht enthaltene Zeichen
' *  werden nicht verschlüsselt. Wählt man die Basis zu hoch, wird fast nichts
' *  verschlüsselt. Ein lustiges Beispiel: ";)" mit dem Standard-Key und Basis
' *  100 verschlüsslet und wieder mit Basis 101 entschlüsselt ergibt ":(" *gg*
' *  Warum? ( und ), sowie ; und : liegen im Key jeweils nebeneinander. Durch die
' *  Differenz 1 der Basis (101 - 100) wird der nächste Buchstabe gewählt.
' *  In diesem Fall entsteht dadurch ein anderer Smiley :) Die Decrypt-Funktion
' *  arbeitet nur mit korrekt verschlüsselten Strings!!
' *
' * ERKLÄRUNG DER PARAMETER FÜR lazyEncrypt:
' *  Data:  Zu verschlüsselnder Text
' *  Key:   Schlüssel, Beispiel: "9182736450zaybxcwdveuftgshriqjpkolmnZAYBXCWDVEUFTGSHRIQJPKOLMN .,-!?_=+/*#~'""§$%&()[]{}äöüßÄÜÖ\@€:;^°<>|"
' *  Base:  Basis für Ascii-Werte, je kleiner, desto besser verschlüsselt.
' * RÜCKGABEWERT VON lazyEncrypt:
' *  Verschlüsselter Text
' *
' * ERKLÄRUNG DER PARAMETER FÜR lazyDecrypt:
' *  Data:  Verschlüsselter Text
' *  Key:   Schlüssel, muss wie bei Encrypt sein!
' *  Base:  Basis für Ascii-Werte, muss wie bei Encrypt sein!
' * RÜCKGABEWERT VON lazyDecrypt:
' *  Entschlüsselter Text
' *

Public Function lazyEncrypt(ByVal Data As String, ByVal Key As String, Optional Base As Integer = 100) As String
Dim Ret As String, i As Double, ac As Integer, cc As String, kp As Integer, kc As String, kac As Integer
' lazyCrypt  (C) Copyright 2001 by Pablo Hoch - mel@melaxis.de - www.melaxis.com
Ret = ""
If Base > 255 Then Base = 255
For i = 1 To Len(Data)
    cc = Mid(Data, i, 1)
    ac = Asc(cc)
    kp = InStr(Key, cc)
    If kp <> 0 Then
        kc = Mid(Key, kp, 1)
        kac = Asc(kc)
        If Base + kp > 255 Then
            cc = Chr(1) & cc
            GoTo weiter
        End If
        cc = Chr(Base + kp)
    Else
        cc = Chr(1) & cc
    End If
weiter:
    Ret = Ret & cc
Next
lazyEncrypt = Ret
End Function

Public Function lazyDecrypt(ByVal Data As String, ByVal Key As String, Optional Base As Integer = 100) As String
Dim Ret As String, i As Double, ac As Integer, cc As String, kp As Integer, kc As String, kac As Integer
' lazyCrypt  (C) Copyright 2001 by Pablo Hoch - mel@melaxis.de - www.melaxis.com
Ret = ""
If Base > 255 Then Base = 255
For i = 1 To Len(Data)
    cc = Mid(Data, i, 1)
    If Asc(cc) = 1 Then
        cc = Mid(Data, i + 1, 1)
        i = i + 1
        GoTo weiter
    End If
    'kc = Chr(Asc(cc) - Base)
    'kp = InStr(Key, kc)
    'If kp <> 0 Then
        'cc = Mid(Key, kp, 1)
        cc = Mid(Key, Asc(cc) - Base, 1)
    'Else
        'cc = kc
    'End If
weiter:
    Ret = Ret & cc
Next
lazyDecrypt = Ret
End Function

' * lazyCrypt - Ende
' ********************************************************************************

Public Function Text2JavaScriptHex(ByVal Text As String) As String
On Error Resume Next
' ©reated 01/05/01 by mel@melaxis.de
' to be inserted as in the following sample:
'<Script Language='Javascript'>
'<!--
'eval(unescape('INSERT THE RETURN VALUE HERE'));
'//-->
'</Script>
' this executes the code you've entered.
' remember to remove all script and comment tags in text!
Dim Ret As String
For i = 1 To Len(Text)
    Ret = Ret & "%" & Hex(Asc(Mid(Text, i, 1)))
Next
Text2JavaScriptHex = Ret
On Error GoTo 0
End Function

'***********************************************************************************
'** Base-Funktionen für die Konvertierung von Dec->Base, Base-Dec und Base->Base.
'** (C) Copyright 2001 by Pablo Hoch - mel@melaxis.de - www.melaxis.com

Public Function Dec2Base(DecZahl As Double, Base As Integer, Optional Chars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ") As String
Dim was As Double, ganz As Double, rest As Double, erg As String
was = DecZahl
Do While was <> 0
    ganz = Int(was / Base)
    rest = was Mod Base
    erg = Mid(Chars, rest + 1, 1) & erg
    was = ganz
Loop
Dec2Base = erg
End Function

Public Function Base2Dec(BaseZahl As String, Base As Integer, Optional Chars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ") As Double
Dim z As String, i As Integer, buf As Long, n As Integer, hoch As Integer, erg As Long
erg = 0
For i = 1 To Len(BaseZahl)
    n = Len(BaseZahl) - i + 1
    z = Mid(BaseZahl, n, 1)
    hoch = i - 1
    buf = (InStr(Chars, z) - 1) * (Base ^ hoch)
    erg = erg + buf
Next
Base2Dec = erg
End Function

Public Function Base2Base(BaseZahl1 As String, Base1 As Integer, Base2 As Integer, Optional Chars1 As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", Optional Chars2 As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ") As String
Dim v1 As Double, v2 As String
v1 = Base2Dec(BaseZahl1, Base1, Chars1)
v2 = Dec2Base(v1, Base2, Chars2)
Base2Base = v2
End Function

'***********************************************************************************

Public Function MostOftenChar(ByVal Text As String, IgnoreSpaces As Boolean, IgnoreCase As Boolean) As String
Dim z As String, i As Double, anz(0 To 255) As Double, hChar As Byte, hCount As Double, canz As Double
If IgnoreCase Then Text = LCase(Text)
For i = 1 To Len(Text)
    z = Mid(Text, i, 1)
    If IgnoreSpaces Then
        If z <> " " And z <> vbCr And z <> vbLf And z <> vbCrLf And z <> vbTab Then
            anz(Asc(z)) = anz(Asc(z)) + 1
        End If
    Else
        anz(Asc(z)) = anz(Asc(z)) + 1
    End If
Next
For i = 0 To 255
    canz = anz(i)
    If hCount < canz Then
        hCount = canz
        hChar = CByte(i)
    End If
Next
MostOftenChar = Chr(hChar)
End Function

Public Function LeastOftenChar(ByVal Text As String, IgnoreSpaces As Boolean, IgnoreCase As Boolean) As String
Dim z As String, i As Double, anz(0 To 255) As Double, hChar As Byte, hCount As Double, canz As Double
If IgnoreCase Then Text = LCase(Text)
For i = 1 To Len(Text)
    z = Mid(Text, i, 1)
    If IgnoreSpaces Then
        If z <> " " And z <> vbCr And z <> vbLf And z <> vbCrLf And z <> vbTab Then
            anz(Asc(z)) = anz(Asc(z)) + 1
        End If
    Else
        anz(Asc(z)) = anz(Asc(z)) + 1
    End If
Next
hCount = Len(Text)
For i = 0 To 255
    canz = anz(i)
    If canz = 0 Then GoTo weiter
    If canz < hCount Then
        hCount = canz
        hChar = CByte(i)
    End If
weiter:
Next
LeastOftenChar = Chr(hChar)
End Function

'****************************************************************************
'** Package of useful assembler byte functions
'** (C) Copyright 2001 by Pablo Hoch - mel@melaxis.de - www.melaxis.com
'**
'** This functions are very useful byte-operations which are not included in
'** vb 6. see the functions for further details. This code may be used
'** in your own projects for free. If you have any suggestions please contact
'** me at mel@melaxis.de. Thanks and have a lot of fun ;-)

Public Function mySHR(Val1 As Double, Val2 As Double) As Double
'** SHR - Shift Right
'** Performs a shift-right operation.
'** Val1 is the number to shift.
'** Val2 is the ammount
'** Example:
'**  Val1 = 149
'**  Val2 = 1
'**   Operation: 10010101 -> 01001010
mySHR = Val1 / (2 ^ Val2)
End Function

Public Function mySHL(Val1 As Double, Val2 As Double) As Double
'** SHL - Shift Left
'** Performs a shift-left operation.
'** Val1 is the number to shift.
'** Val2 is the ammount
'** Example:
'**  Val1 = 149
'**  Val2 = 1
'**   Operation: 10010101 -> 00101010
mySHL = Val1 * (2 ^ Val2)
End Function

Public Function myROR(Val1 As Double, Val2 As Double) As Double
Dim Ret As Double, buf As String, b1 As String, b2 As String
'** ROR - Rotate Right
'** Performs a rotate-right operation.
'** Val1 is the number to rotate.
'** Val2 is the ammount
'** Example:
'**  Val1 = 149
'**  Val2 = 1
'**   Operation: 10010101 -> 11001010
Ret = Val1
buf = Dec2Base(Val1, 2)
For i = 1 To Val2
    b1 = Right(buf, 1)
    buf = b1 & Mid(buf, 1, 7)
    Ret = Base2Dec(buf, 2)
Next i
myROR = Ret
End Function

Public Function myROL(Val1 As Double, Val2 As Double) As Double
Dim Ret As Double, buf As String, b1 As String, b2 As String
'** ROL - Rotate Left
'** Performs a rotate-left operation.
'** Val1 is the number to rotate.
'** Val2 is the ammount
'** Example:
'**  Val1 = 149
'**  Val2 = 1
'**   Operation: 10010101 -> 00101011
Ret = Val1
buf = Dec2Base(Val1, 2)
For i = 1 To Val2
    b1 = Left(buf, 1)
    buf = Mid(buf, 2) & b1
    Ret = Base2Dec(buf, 2)
Next i
myROL = Ret
End Function

'****************************************************************************

Public Function AlphaChar(Char As String) As Boolean
Dim cuca As Byte
'    48-57   0-9
'    65-90   A-Z
'    97-122  a-z
    cuca = Asc(Char)
    If (cuca >= 48 And cuca <= 57) Or (cuca >= 65 And cuca <= 90) Or (cuca >= 97 And cuca <= 122) Then
        AlphaChar = True
    Else
        AlphaChar = False
    End If
End Function

Public Function MakeFlatButton(btn As CommandButton)
On Error Resume Next
Dim style As Long
 
style = GetWindowLong(btn.hwnd, GWL_STYLE)
style = style Or BS_FLAT
 
SetWindowLong btn.hwnd, GWL_STYLE, style
btn.Refresh
 
End Function

Public Function MakeFlatObject(obj As Object)
On Error Resume Next
Dim style As Long
 
style = GetWindowLong(obj.hwnd, GWL_STYLE)
style = style Or BS_FLAT
 
SetWindowLong obj.hwnd, GWL_STYLE, style
obj.Refresh

' auch frames
 
End Function

Public Function MakeFlathWnd(hwnd As Long)
On Error Resume Next
Dim style As Long
 
style = GetWindowLong(hwnd, GWL_STYLE)
style = style Or BS_FLAT
 
SetWindowLong hwnd, GWL_STYLE, style
 
End Function

Public Sub MakeAllFlat(Frm As Form)
On Error Resume Next
Dim Item As Object
For Each Item In Frm.Controls
    MakeFlatObject Item
Next
End Sub


Public Function IsFormLoaded(Form As Form) As Boolean
  Dim nForm As Form
  
  For Each nForm In Forms
    If nForm Is Form Then
      IsFormLoaded = True
      Exit For
    End If
  Next
End Function


Public Sub SetBitB(Value As Byte, ByVal Position As Byte)
  Select Case Position
    Case 0 To 7
      Value = Value Or 2 ^ Position
    Case Else
      Err.Raise 6
  End Select
End Sub

Public Sub ClearBitB(Value As Byte, ByVal Position As Byte)
  Select Case Position
    Case 0 To 7
      Value = Value And Not 2 ^ Position
    Case Else
      Err.Raise 6
  End Select
End Sub

Public Function BitB(ByVal Value As Byte, ByVal Position As Byte) _
 As Boolean

  Select Case Position
    Case 0 To 7
      BitB = CBool(Value And 2 ^ Position)
    Case Else
      Err.Raise 6
  End Select
End Function

Public Sub SetBitI(Value As Integer, ByVal Position As Byte)
  Select Case Position
    Case 0 To 14
      Value = Value Or 2 ^ Position
    Case 15
      Value = Value Or &H8000
    Case Else
      Err.Raise 6
  End Select
End Sub

Public Sub ClearBitI(Value As Integer, ByVal Position As Byte)
  Select Case Position
    Case 0 To 14
      Value = Value And Not 2 ^ Position
    Case 15
      Value = Value And Not &H8000
    Case Else
      Err.Raise 6
  End Select
End Sub

Public Function BitI(ByVal Value As Integer, ByVal Position As Byte) _
 As Boolean

  Select Case Position
    Case 0 To 14
      BitI = CBool(Value And 2 ^ Position)
    Case 15
      BitI = CBool(Value < 0)
    Case Else
      Err.Raise 6
  End Select
End Function

Public Sub SetBitL(Value As Long, ByVal Position As Byte)
  Dim nVal As Variant
  
  Select Case Position
    Case 0 To 30
      Value = Value Or 2 ^ Position
    Case 31
      Value = Value Or &H80000000
    Case Else
      Err.Raise 6
  End Select
End Sub

Public Sub ClearBitL(Value As Long, ByVal Position As Byte)
  Select Case Position
    Case 0 To 30
      Value = Value And Not 2 ^ Position
    Case 31
      Value = Value And Not &H80000000
    Case Else
      Err.Raise 6
  End Select
End Sub

Public Function BitL(ByVal Value As Long, ByVal Position As Byte) _
 As Boolean

  Select Case Position
    Case 0 To 30
      BitL = CBool(Value And 2 ^ Position)
    Case 31
      BitL = CBool(Value < 0)
    Case Else
      Err.Raise 6
  End Select
End Function

Public Function DecBToBin(ByVal Value As Byte, _
 Optional ByVal DoFormat As Boolean) As String

  Dim i As Integer
  Dim nBin As String
  
  nBin = String$(8, "0")
  For i = 0 To 7
    If BitL(Value, i) Then
      Mid$(nBin, 8 - i, 1) = "1"
    End If
  Next
  If DoFormat Then
    DecBToBin = nBin
  Else
    i = InStr(nBin, "1")
    If i Then
      DecBToBin = Mid$(nBin, i)
    Else
      DecBToBin = "0"
    End If
  End If
End Function

Public Function BinToDecB(Bin As String) As Byte
  Dim i As Integer
  Dim nDec As Byte
  Dim nPos As Integer
  
  If Len(Bin) > 7 Then
    Err.Raise 6
  Else
    For i = Len(Bin) To 1 Step -1
      If Mid$(Bin, i, 1) = "1" Then
        SetBitB nDec, nPos
      End If
      nPos = nPos + 1
    Next 'i
  End If
  BinToDecB = nDec
End Function

Public Function EasyCRC32(str As String) As Long
  Dim i As Long
  Dim j As Long
  Dim nPowers(0 To 7) As Integer
  Dim nCRC As Long
  Dim nByte As Integer
  Dim nBit As Boolean
  
  For i = 0 To 7
     nPowers(i) = 2 ^ i
  Next 'i
  For i = 1 To Len(str)
    nByte = Asc(Mid$(str, i, 1))
    For j = 7 To 0 Step -1
      nBit = CBool((nCRC And 32768) = 32768) Xor _
       ((nByte And nPowers(j)) = nPowers(j))
      nCRC = (nCRC And 32767&) * 2&
      If nBit Then
        nCRC = nCRC Xor &H8005&
      End If
    Next 'j
  Next 'i
  EasyCRC32 = nCRC
End Function

Public Function PrimesToArray(ByVal Number As Long) _
 As Variant

  Dim nArray() As Long
  Dim na As Long
  Dim nb As Long
  Dim nW As Long
  Dim i As Integer
  Dim nIndex As Long
  
  If Number = 0 Then
    Exit Function
  End If
  ReDim nArray(1 To 10)
  na = Number
  nb = 2
  nW = na \ 2
  For i = 1 To 10
    nW = (nW + (na \ nW)) \ 2
  Next 'i
  Do While na <> 1
    Do While na Mod nb = 0
      na = na \ nb
      nIndex = nIndex + 1
      If UBound(nArray) < nIndex Then
        ReDim Preserve nArray(1 To nIndex + 10)
      End If
      nArray(nIndex) = nb
    Loop
    If nb > nW Then
      nb = na - 2
    End If
    If nb = 2 Then
      nb = 1
    End If
    nb = nb + 2
  Loop
  ReDim Preserve nArray(1 To nIndex)
  PrimesToArray = nArray
End Function

Public Function PrimesToCollection(ByVal Number As Long) _
 As Collection

  Dim nColl As Collection
  Dim na As Long
  Dim nb As Long
  Dim nW As Long
  Dim i As Integer
  
  If Number = 0 Then
    Exit Function
  End If
  Set nColl = New Collection
  na = Number
  nb = 2
  nW = na \ 2
  For i = 1 To 10
    nW = (nW + (na \ nW)) \ 2
  Next 'i
  Do While na <> 1
    Do While na Mod nb = 0
      na = na \ nb
      nColl.Add nb
    Loop
    If nb > nW Then
      nb = na - 2
    End If
    If nb = 2 Then
      nb = 1
    End If
    nb = nb + 2
  Loop
  Set PrimesToCollection = nColl
End Function



Public Function melGetStringCount(ByVal myStr As String, ByVal searchFor As String) As Double
On Error Resume Next
Dim Ret As Double
Ret = 0
For i = 1 To Len(myStr)
    If Mid(myStr, i, Len(searchFor)) = searchFor Then
        Ret = Ret + 1
    End If
Next
melGetStringCount = Ret
End Function

Public Function KillProcessWithAllThreads(ByVal CaptionPart As String, ByRef KILLED As Collection, Frm As Form) As Long
' Copyright 2001 by Pablo Hoch aka mel
' www.melaxis.com - mel@melaxis.de

Dim WND As Long, str As String, le As Long, pid As Long
Dim PROCS As New Collection, buf As String, anz As Long
Dim hPro As Long, ExCo As Long, KiCo As Long
WND = Frm.hwnd
WND = GetNextWindow(WND, GW_HWNDFIRST)
Do While WND <> 0
    WND = GetNextWindow(WND, GW_HWNDNEXT)
    If WND = 0 Then Exit Do
    le = GetWindowTextLength(WND)
    str = String(le, 0)
    Call GetWindowText(WND, str, le + 1)
    GetWindowThreadProcessId WND, pid
    If Trim(str) <> "" Then
        buf = WND & vbTab & pid & vbTab & vbTab & str
    Else
        buf = WND & vbTab & pid & vbTab & vbTab & "???"
    End If
    PROCS.Add buf
Loop
DoEvents
For i = 1 To PROCS.Count
    d = Split(PROCS(i), vbTab)
    ' 0     1     2 3
    ' hWnd  pid     Caption
    '     ^    ^  ^
    If InStr(d(3), CaptionPart) Then
        ' Gefunden!!
        WND = d(0)
        pid = d(1)
        str = d(3)
        Exit For
    End If
Next
If WND = 0 Then
    ' nichts gefunden ;-(
    SetLastError 1
    KillProcessWithAllThreads = 0
    Exit Function
End If
anz = 0
For i = 1 To PROCS.Count
    d = Split(PROCS(i), vbTab)
    ' 0     1     2 3
    ' hWnd  pid     Caption
    '     ^    ^  ^
    If d(1) = pid Then
        anz = anz + 1
        KILLED.Add d(0) & vbTab & vbTab & d(3)
    End If
Next
' und jetzt abschiessen...
hPro = OpenProcess(PROCESS_ALL_ACCESS, 0&, pid)
If hPro <> 0 Then
    ' exit code holen
    Call GetExitCodeProcess(hPro, ExCo)
    If ExCo <> 0 Then
        KiCo = TerminateProcess(hPro, ExCo)
        If KiCo = 0 Then
            SetLastError 2
            KillProcessWithAllThreads = 0
            Exit Function
        End If
    End If
Else
    ' kann process nicht öffnen
    SetLastError 3
    KillProcessWithAllThreads = 0
    Exit Function
End If
SetLastError 0
KillProcessWithAllThreads = anz
End Function

Public Sub KillAllThreads(Frm As Form, KillSelf As Boolean)
' Copyright 2001 by Pablo Hoch aka mel
' www.melaxis.com - mel@melaxis.de

Dim WND As Long, str As String, le As Long, pid As Long
Dim PROCS As New Collection, buf As String, anz As Long
Dim hPro As Long, ExCo As Long, KiCo As Long, myPid As Long
WND = Frm.hwnd
Call GetWindowThreadProcessId(WND, myPid)
WND = GetNextWindow(WND, GW_HWNDFIRST)
Do While WND <> 0
    WND = GetNextWindow(WND, GW_HWNDNEXT)
    If WND = 0 Then Exit Do
    le = GetWindowTextLength(WND)
    str = String(le, 0)
    Call GetWindowText(WND, str, le + 1)
    Call GetWindowThreadProcessId(WND, pid)
    If Trim(str) <> "" Then
        buf = WND & vbTab & pid & vbTab & vbTab & str
    Else
        buf = WND & vbTab & pid & vbTab & vbTab & "???"
    End If
    PROCS.Add buf
Loop
DoEvents
anz = 0
For i = 1 To PROCS.Count
    d = Split(PROCS(i), vbTab)
    ' 0     1     2 3
    ' hWnd  pid     Caption
    '     ^    ^  ^
    If d(1) <> myPid Then
        WND = d(0)
        pid = d(1)
        str = d(3)
        hPro = OpenProcess(PROCESS_ALL_ACCESS, 0&, pid)
        If hPro <> 0 Then
            ' exit code holen
            Call GetExitCodeProcess(hPro, ExCo)
            If ExCo <> 0 Then
                KiCo = TerminateProcess(hPro, ExCo)
                If KiCo = 0 Then
                    anz = anz + 1
                End If
            End If
        Else
            ' kann process nicht öffnen
        End If
    End If
Next
DoEvents
If KillSelf Then
    WND = Frm.hwnd
    pid = myPid
    str = Frm.Caption
    hPro = OpenProcess(PROCESS_ALL_ACCESS, 0&, pid)
    If hPro <> 0 Then
        ' exit code holen
        Call GetExitCodeProcess(hPro, ExCo)
        If ExCo <> 0 Then
            KiCo = TerminateProcess(hPro, ExCo)
            If KiCo = 0 Then
                anz = anz + 1
            End If
        End If
    Else
        ' kann process nicht öffnen
    End If
End If
End Sub

Public Sub ExecuteAsmHex(s As String)
On Error Resume Next
' ripped from a Damian-source.
' example parameter: "578B7C240C33C00FA28AC366AB8AC766ABC1EB108AC366AB8AC766AB8BDA8AC366AB8AC766ABC1EB108AC366AB8AC766AB8BD98AC366AB8AC766ABC1EB108AC366AB8AC766AB5F33C0C20800"
' donno what it does :p
    s = Replace$(s, " ", "")
    Dim i As Long, aSize As Long, aB() As Byte
    aSize = Len(s) \ 2

    ReDim Preserve aB(1 To aSize)
    For i = 1 To aSize
        aB(i) = Val("&H" & Mid$(s, i * 2 - 1, 2))
    Next
    
    Static cp As Long
    ReDim Preserve aProc(cp)
    Dim hmem As Long, lPtr As Long
    hmem = GlobalAlloc(0, aSize)
    lPtr = GlobalLock(hmem)
    CopyMemory ByVal lPtr, aB(1), aSize
    GlobalUnlock hmem
    
'    aProc(cp).hMem = hMem
'    aProc(cp).vtPtr = VTable(cp)
'    VTable(cp) = lPtr
'    cp = cp + 1
End Sub

' CHECKSUMXOR,CHECKSUMOR,CHECKSUMAND,CHECKSUMXOA (C) Copyright 2001 Pablo Hoch aka mel - mel@melaxis.de - www.melaxis.com
Public Function CheckSumXOR(Data As String) As Byte
Dim buf As String
If Len(Data) < 3 Then
    CheckSumXOR = Data
    Exit Function
End If
buf = Chr(Asc(Mid(Data, 1, 1)) Xor Asc(Mid(Data, 2, 1)))
For i = 3 To Len(Data)
    buf = Chr(Asc(buf) Xor Asc(Mid(Data, i, 1)))
Next
CheckSumXOR = Asc(buf)
End Function

Public Function CheckSumOR(Data As String) As Byte
Dim buf As String
If Len(Data) < 3 Then
    CheckSumOR = Data
    Exit Function
End If
buf = Chr(Asc(Mid(Data, 1, 1)) Or Asc(Mid(Data, 2, 1)))
For i = 3 To Len(Data)
    buf = Chr(Asc(buf) Or Asc(Mid(Data, i, 1)))
Next
CheckSumOR = Asc(buf)
End Function

Public Function CheckSumAND(Data As String) As Byte
Dim buf As String
If Len(Data) < 3 Then
    CheckSumAND = Data
    Exit Function
End If
buf = Chr(Asc(Mid(Data, 1, 1)) And Asc(Mid(Data, 2, 1)))
For i = 3 To Len(Data)
    buf = Chr(Asc(buf) And Asc(Mid(Data, i, 1)))
Next
CheckSumAND = Asc(buf)
End Function

Public Function CheckSumXOA(Data As String) As String
Dim buf As String
buf = Hex(CheckSumXOR(Data)) & Hex(CheckSumOR(Data)) & Hex(CheckSumAND(Data))
CheckSumXOA = String(6 - Len(buf), "0") & buf
End Function

Public Function GetCharCount(Txt As String, Char As String) As Long
Dim cnt As Long
cnt = 0
For i = 1 To Len(Char)
    If Mid(Txt, i, 1) = Char Then
        cnt = cnt + 1
    End If
Next
GetCharCount = cnt + 1
End Function

Public Function Deg2Rad(Grad As Single) As Single
Dim pi As Double
pi = Atn(1) * 4
Deg2Rad = Grad * (pi / 180#)
End Function

Public Function Rad2Deg(Rad As Single) As Single
Dim pi As Double
pi = Atn(1) * 4
Rad2Deg = Rad * (180# / pi)
End Function

Public Function IPToLong(sIP As String) As Long
    Dim ssIP() As String, s As String, i As Integer
    ssIP = Split(sIP, ".")
    s = "&H"


    For i = 0 To 3
        s = s & Format(Hex(ssIP(i)), "00")
    Next
    IPToLong = Val(s)
End Function


Public Function LongToIP(IP As Long) As String
    Dim s As String, iIP(3) As Integer, i As Integer
    s = Format(Hex(IP), "00000000")


    For i = 0 To 3
        iIP(i) = Val("&H" & Mid(s, 1 + i * 2, 2))
    Next
    LongToIP = CStr(iIP(0)) & "." & CStr(iIP(1)) & "." & CStr(iIP(2)) & "." & CStr(iIP(3))
End Function



Public Function GetIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim HOST As HOSTENT
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim i As Integer
    Dim sIPAddr As String


    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If


    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPAddress = ""
        MsgBox "Windows Sockets Error " & str$(WSAGetLastError()) & _
        " has occurred. Unable To successfully Get Host Name."
        SocketsCleanup
        Exit Function
    End If
    sHostName = Trim$(sHostName)
    lpHost = gethostbyname(sHostName)


    If lpHost = 0 Then
        GetIPAddress = ""
        MsgBox "Windows Sockets are Not responding. " & _
        "Unable To successfully Get Host Name."
        SocketsCleanup
        Exit Function
    End If
    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4
    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen


    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next
    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function


Public Function GetIPHostName() As String
    Dim sHostName As String * 256


    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If


    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets Error " & str$(WSAGetLastError()) & _
        " has occurred. Unable To successfully Get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function


Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
    
End Function


'Public Function LoByte(ByVal wParam As Integer)
'    LoByte = wParam And &HFF&
'End Function


'Public Sub SocketsCleanup()
'
'
'    If WSACleanup() <> ERROR_SUCCESS Then
'        MsgBox "Socket Error occurred In Cleanup."
'    End If
'End Sub


'Public Function SocketsInitialize() As Boolean
'    Dim WSAD As WSADATA
'    Dim sLoByte As String
'    Dim sHiByte As String
'
'
'    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
'        MsgBox "The 32-bit Windows Socket is Not responding."
'        SocketsInitialize = False
'        Exit Function
'    End If
'
'
'    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
'        MsgBox "This application requires a minimum of " & _
'        CStr(MIN_SOCKETS_REQD) & " supported sockets."
'        SocketsInitialize = False
'        Exit Function
'    End If
'    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
'    (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
'    HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
'
'    sHiByte = CStr(HiByte(WSAD.wVersion))
'    sLoByte = CStr(LoByte(WSAD.wVersion))
'
'    MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
'    " is Not supported by 32-bit Windows Sockets."
'
'    SocketsInitialize = False
'    Exit Function
'
'End If
'SocketsInitialize = True
'End Function


' EOF
