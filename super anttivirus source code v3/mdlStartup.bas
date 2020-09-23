Attribute VB_Name = "mdlStartup"
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hkey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Const KEY_QUERY_VALUE = &H1
Private Const MAX_PATH = 260

Private Enum RegDataTypes
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_DWORD = 4                      ' 32-bit number
End Enum

Private Enum RegistryKeys
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Enum ValKey
    Values = 0
    Keys = 1
End Enum

Private Type ByteArray
  FirstByte As Byte
  ByteBuffer(255) As Byte
End Type

Dim baData As ByteArray
Dim REG As New clsRegistry

Private Function OpenKey(RegistryKey As RegistryKeys, Optional SubKey As String) As Long
    If OpenKey <> 0 Then RegCloseKey (OpenKey)
    RegOpenKeyEx RegistryKey, SubKey, 0, KEY_QUERY_VALUE, OpenKey
End Function

Private Function GetCount(RegisteryKeyHandle As Long, ValuesOrKeys As ValKey) As Long
    If ValuesOrKeys = Keys Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, GetCount, 0, 0, 0, 0, 0, 0, 0
    If ValuesOrKeys = Values Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, GetCount, 0, MAX_PATH + 1, 0, 0
End Function

Private Function EnumKey(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    EnumKey = Space(MAX_PATH + 1)
    RegEnumKey RegisteryKeyHandle, KeyIndex, EnumKey, MAX_PATH + 1
    EnumKey = Trim(EnumKey)
End Function

Private Function EnumValue(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    Dim lBufferLen As Long, I As Integer
    For I = 0 To 255
      baData.ByteBuffer(I) = 0
    Next
    lBufferLen = 255
    EnumValue = Space(MAX_PATH + 1)
    RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, 0, lValNameLen, lValLen, 0, 0
    RegEnumValue RegisteryKeyHandle, KeyIndex, EnumValue, MAX_PATH + 1, 0, 0, baData.FirstByte, lBufferLen
    EnumValue = Trim(EnumValue)
End Function

Private Function DeleteValue(RegisteryKeyHandle As Long, KeyName As String) As Long
    DeleteValue = RegDeleteValue(RegisteryKeyHandle, KeyName)
End Function

Private Function SetValue(RegisteryKeyHandle As RegistryKeys, SubRegistryKey As String, KeyName As String, newValue As String, Optional DataType As RegDataTypes)
    Dim lRetVal As Long
    lRetVal = OpenKey(RegisteryKeyHandle, SubRegistryKey)
    If DataType = 0 Then DataType = REG_SZ
    RegSetValueEx lRetVal, KeyName, 0, DataType, newValue, LenB(StrConv(SubKeyValue, vbFromUnicode))
End Function

Private Function GetKeyValue(hkey As Long, KeyName As String) As String
    Dim I As Long
    Dim rc As Long
    
    Dim hDepth As Long
    Dim sKeyVal As String
    Dim lKeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hkey, KeyName, 0, lKeyValType, tmpVal, KeyValSize)
    GetKeyValue = Trim(tmpVal)
    
End Function

Public Function GetAllRun()
On Error Resume Next
    
    Dim hkey As Long
    Dim lCount As Long
    Dim I As Long
    
    frmScanVirus.lstStartup.Clear
    frmScanVirus.List5.Clear
    frmScanVirus.List6.Clear

    Select Case frmScanVirus.cboStartup.Text
        Case "HKLM - Run"
        hkey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKLM - Run"
        Next I
    
    Case "HKLM - RunOnce"
        hkey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKLM - RunOnce"
        Next I
    
    Case "HKLM - RunOnceEx"
        hkey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKLM - RunOnceEx"
        Next I
    
    Case "HKLM - RunServices"
        hkey = OpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKLM - RunServices"
        Next I
    
    Case "HKCU - Run"
        hkey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKCU - Run"
        Next I
        
    Case "HKCU - PoliciesRun"
        hkey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKCU - PoliciesRun"
        Next I
    
    Case "HKCU - RunOnce"
        hkey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKCU - RunOnce"
        Next I
    
    Case "HKCU - Windows"
        hkey = OpenKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows")
        lCount = GetCount(hkey, Values)
        For I = 0 To lCount - 1
            frmScanVirus.lstStartup.AddItem EnumValue(hkey, I)
            frmScanVirus.List5.AddItem GetKeyValue(hkey, EnumValue(hkey, I))
            frmScanVirus.List6.AddItem "HKCU - Windows"
        Next I

    Dim fso As New FileSystemObject
    Dim sFolder As Folder
    Dim sFiles As Files
    Dim sFile As file
    
    Case "Scheduled Task"
    Set sFolder = fso.GetFolder("C:\Windows\Tasks")
    Set sFiles = sFolder.Files
    If sFiles.count > 0 Then
        For Each sFile In sFiles
            frmScanVirus.lstStartup.AddItem (sFile.name)
            frmScanVirus.List5.AddItem sFile.path
            frmScanVirus.List6.AddItem "Scheduled Task"
        Next
    End If

    Case "User Startup"
    Dim strUserProfile As String
    strUserProfile = Environ$("UserProfile") & "\Start Menu\Programs\Startup"
    Set sFolder = fso.GetFolder(strUserProfile)
    Set sFiles = sFolder.Files
    If sFiles.count > 0 Then
        For Each sFile In sFiles
            frmScanVirus.lstStartup.AddItem (sFile.name)
            frmScanVirus.List5.AddItem sFile.path
            frmScanVirus.List6.AddItem "User Startup"
        Next
    End If
    
    Case "All Users Startup"
    Set sFolder = fso.GetFolder("C:\Documents and Settings\All Users\Start Menu\Programs\Startup")
    Set sFiles = sFolder.Files
    If sFiles.count > 0 Then
        For Each sFile In sFiles
            frmScanVirus.lstStartup.AddItem (sFile.name)
            frmScanVirus.List5.AddItem sFile.path
            frmScanVirus.List6.AddItem "All Users Startup"
        Next
    End If
    Case Else
        frmScanVirus.txtPathStartup.Text = "Please choose startup location ..."
    End Select
End Function

Public Function ClearAuto()
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    
    On Error Resume Next
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        Kill drv.DriveLetter & ":\autorun.inf"
    Next
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
End Function

Public Function ClearAutorun()

    Dim I As Long
    Dim tmp As Long
    Dim fso As New FileSystemObject
        
    Select Case frmScanVirus.List6.Text
                Case "HKLM - RunServices"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices", frmScanVirus.lstStartup.Text
                Case "HKLM - Run"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", frmScanVirus.lstStartup.Text
                Case "HKCU - Run"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", frmScanVirus.lstStartup.Text
                Case "HKCU - PoliciesRun"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", frmScanVirus.lstStartup.Text
                Case "HKLM - RunOnce"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnce", frmScanVirus.lstStartup.Text
                Case "HKLM - RunOnceEx"
                    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunOnceEx", frmScanVirus.lstStartup.Text
                Case "HKCU - RunOnce"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", frmScanVirus.lstStartup.Text
                Case "HKCU - Windows"
                    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", frmScanVirus.lstStartup.Text
                Case Else
                    fso.DeleteFile frmScanVirus.txtPathStartup.Text, True
    End Select
    
    Set fso = Nothing
    Call GetAllRun
End Function



