Attribute VB_Name = "mdlAPI"
Option Explicit

Public Const MAX_PATH As Integer = 260

' Processes dan Module --------------------------
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type

' Set the priority class
' for the specified process -----------------------------
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long

Public Const THREAD_SUSPEND_RESUME As Long = &H2
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const PROCESS_SET_INFORMATION As Long = (&H200)

' Sets a file’s attributes -------------------------
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As FILE_ATTRIBUTE) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Enum FILE_ATTRIBUTE
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_TEMPORARY = &H100
    FILE_ATTRIBUTE_COMPRESSED = &H800
End Enum

' Opens the access token associated with a thread ------------------
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapShot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean

Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

' Adjusts privileges in the specified access token -------------------
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges32 Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES32, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES32, ReturnLength As Long) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2
Public Const ANYSIZE_ARRAY = 1

Public Const VER_PLATFORM_WIN32_NT = 2
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4

Public Type LUID
    lowpart As Long
    highpart As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Public Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Public Type TOKEN_PRIVILEGES32
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

' Retrieves the handle of the menu
' assigned to the given window ----------------------
Public Declare Function GetMenu& Lib "user32" (ByVal hWnd&)
Public Declare Function GetSubMenu& Lib "user32" (ByVal hMenu&, ByVal nPos&)
Public Declare Function GetMenuItemID& Lib "user32" (ByVal hMenu&, ByVal nPos&)
Public Declare Function ModifyMenu& Lib "user32" Alias "ModifyMenuA" (ByVal hMenu&, ByVal nPosition&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpString$)
Public Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)

' Creates an Open common dialog box that lets the user specify
' the drive, directory, and the name of a file or set of files to open -----------
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As SAVEFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As SAVEFILENAME) As Long
Public Declare Function GetOpenFileNameEx Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Const BIF_NEWDIALOGSTYLE As Long = &H40
Public Const BIF_EDITBOX As Long = &H10
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_SETSELECTION = (WM_USER + 102)
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_STATUSTEXT = &H4&
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_EXPLORER = &H80000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_HIDEREADONLY = &H4

Public Type SAVEFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type BrowseInfo
    lnghWnd As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

' -------------------------------
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' Retrieves the location of a special folder
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long

Public Enum sFolder
    CSIDL_DESKTOP = &H0
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
    CSIDL_PROFILER = &H28
    CSIDL_HISTORY = &H22
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_APPDATA = &H1A
End Enum

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

' -------------------------------
Public Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal sID As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Integer) As Long
Public Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Public Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const WTS_CURRENT_SERVER_HANDLE = 0&

Public Type WTS_PROCESS_INFO
    SessionID As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
End Type

' -------------------------------
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Public blnClick As Boolean
Public vbTray As NOTIFYICONDATA

Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const FLAGS As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONCLK As Long = &H204
Public Const WM_LBUTTONCLK As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_MOUSEMOVE As Long = &H200
Public Const NIM_ADD As Long = &H0
Public Const NIM_DELETE As Long = &H2
Public Const NIF_ICON As Long = &H2
Public Const NIF_MESSAGE As Long = &H1
Public Const NIM_MODIFY As Long = &H1
Public Const NIF_TIP As Long = &H4
Public Const NIF_INFO As Long = &H10
Public Const NIS_HIDDEN As Long = &H1
Public Const NIS_SHAREDICON As Long = &H2

Public Enum TypeBallon
    NIIF_NONE = &H0
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_INFO = &H1
    NIIF_GUID = &H4
End Enum

Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

' -----------------------------------
Public Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal Color As Long, ByVal X As Byte, ByVal ALPHA As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Const GW_HWNDNEXT = 2
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const GW_OWNER = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

' INI File ----------------------------------
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' Returns version information about a specified file ------------------------------
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

' Returns selected version information from the
' specified version-information resource ----------------------------
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

' -------------------------------------------
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

' -------------------------------------------
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

' --------------------------------------
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' --------------------------------------
Public Declare Function GetUSERNAME Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

' --------------------------------------
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' --------------------------------------
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' --------------------------------------
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

' mdlMemory ---------------------------------
Public Declare Function PdhOpenQuery Lib "PDH.DLL" (ByVal Reserved As Long, ByVal dwUserData As Long, ByRef hQuery As Long) As PDH_STATUS
Public Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As PDH_STATUS
Public Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As PDH_STATUS
Public Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUS)

Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'CPU Information (Using: Windows Performance Data Helper DLL)
Public Enum PDH_STATUS
    PDH_CSTATUS_VALID_DATA = &H0
    PDH_CSTATUS_NEW_DATA = &H1
End Enum

Public Type CounterInfo
    hCounter As Long
    strName As String
End Type

' Registry access
' mdlRegistry -----------------------------------
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A
Public Const SPI_SETNONCLIENTMETRICS As Long = &H2A
Public Const SMTO_ABORTIFHUNG As Long = &H2

Public Type SECURITY_ATTRIBUTES
   nLength                 As Long
   lpSecurityDescriptor    As Long
   bInheritHandle          As Long
End Type

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
    
Public Enum TipeReg
    REG_SZ = 1                         ' Unicode nul terminated string
    reg_binarY = 3                     ' Free form binary
    REG_DWORD = 4                      ' 32-bit number
End Enum

Public Const ERROR_SUCCESS = 0&

' Ensures that the common control
' dynamic-link library (DLL) is loaded -----------------------------
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

' Maps the specified executable module into
' the address space of the calling process -------------------------
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long

