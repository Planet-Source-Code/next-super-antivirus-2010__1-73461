Attribute VB_Name = "mdlGlobalFunction"
Option Explicit

' Constans
Public Const APP_VERSION = "3 Beta"
Public Const CURRENT_BUILD = "15 September 2010"
Public Const ENGINE_VERSION = "Engine 3 Beta"
Public Const APP_PROGRAM = "Av Super Protector"
Public Const TWEAK_REG = "2.1"
Public Const PROCESSESES = "2.1"

' Variabels
Public StopScan As Boolean
Public PauseScan As Boolean
Public Buffering As Boolean
Public DeleteAll As Boolean
Public Heutrue As Boolean
Public RegistryFix As Boolean
Public TipeHeuristic As String
Public ANSAVEnable As Boolean
Public ExternalEnable As Boolean
Public InternalEnable As Boolean
Public Unhide As Boolean
Public IconCompare As Boolean
Public DisBuffer As Boolean
Public RAREnable As Boolean
Public JumlahBuffer As Long
Public JumlahFile As Long
Public VirusDetected As Long
Public Akurat As Integer
Public VirusCleaned As Long
Public Where As String
Public ext As String
Public Persen As Integer
Public StopButton As Boolean
Public SignTemp As String
Public lTuan As Long

Public VersiAnsav As String
Public JumlahSignature As String
Public Namaku As String
Public t As Single
Public strUserCom As String
Public ComName As Long
Public PCName As String

Public isStop  As Boolean
Public myFileCol As New Collection

' Procedure untuk TabStrips XP
Sub ControlTabs(ByVal oForm As Form, ByVal oTabStrip As TabStrip, ByVal iTabIndex As Integer)

    Dim iTabCnt, iCntFrames As Integer
    On Error GoTo Error_ControlTabs

    iTabCnt = (oTabStrip.Tabs.count - 1)
    iTabIndex = (iTabIndex - 1)

    For iCntFrames = 0 To iTabCnt
        oForm.Frame1(iCntFrames).Visible = False
        If iCntFrames = iTabIndex Then oForm.Frame1(iCntFrames).Visible = True
    Next iCntFrames
    
    Exit Sub
    
Error_ControlTabs:
    Select Case Err.Number
        Case 340: MsgBox "Frame index(" & iCntFrames & ") doesn't exist in the ControlArray." & _
                  vbCrLf & "Make sure you add this frame to your TabStrip control.", _
                  vbCritical, "Error " & Err.Number
        Case Else: MsgBox Err.Description, vbCritical, "Error " & Err.Number
    End Select

End Sub
' End procedure TabStrips

Public Function NameOfTheComputer(MachineName As String) As Long

    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    
End Function

Public Function GetUserCom() As String

    GetUserCom = Environ$("username")
    ComName = NameOfTheComputer(PCName)

    frmScanVirus.StatusBar1.Panels(4).Text = "Hello : " + GetUserCom
    frmScanVirus.StatusBar1.Panels(5).Text = "Registered To : " + PCName
    
End Function

'Sub UserName()
'
'    strUserCom = String(100, Chr$(0))
'    GetUSERNAME strUserCom, 100
'    strUserCom = Left$(strUserCom, InStr(strUserCom, Chr$(0)) - 1)
'    ComName = NameOfTheComputer(PCName)
'
'    frmScanVirus.StatusBar1.Panels(3).Text = "Hello : " + strUserCom
'    frmScanVirus.StatusBar1.Panels(4).Text = "Registered To : " + PCName
'
'End Sub

Sub Tunggu(ByVal scd As Single)
Dim isStop As Boolean
On Error Resume Next
    Dim mulai As Variant
    mulai = Timer
    Do While Timer < mulai + scd
      DoEvents
      If isStop = True Then Exit Do
    Loop
End Sub

Function nPath(mypath As String) As String
If Right(mypath, 1) = "\" Then
   nPath = mypath
Else
   nPath = mypath & "\"
End If
End Function

Function TempWindow() As String
Dim buff As String
buff = String(255, 0)
GetTempPath 255, buff
TempWindow = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

Function MyWindowDir() As String
Dim buff As String
buff = String(255, 0)
GetWindowsDirectory buff, 255
MyWindowDir = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

Function MyWindowSys() As String
Dim buff As String
buff = String(255, 0)
GetSystemDirectory buff, 255
MyWindowSys = nPath(Left(buff, InStr(1, buff, Chr(0)) - 1))
End Function

Private Function CekIsExist(ID As String) As Boolean
On Error GoTo salah
   Dim tmp As String
   tmp = myFileCol(ID)
   CekIsExist = True
   Exit Function
salah:
End Function

Function KillProcessByName(nName As String, Optional hID As Long = 0) As Long
On Error Resume Next
If IsWinNT Then
    If nName = "" Then Exit Function
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32, R As Long
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    R = Process32First(hSnapShot, uProcess)
    Do While R
        If LCase(Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))) = LCase(nName) Then
           If hID = 0 Then
              KillProcessById uProcess.th32ProcessID
              KillProcessById uProcess.th32ParentProcessID
           Else
              If uProcess.th32ProcessID = hID Then
                 KillProcessById uProcess.th32ProcessID
                 KillProcessById uProcess.th32ParentProcessID
              End If
           End If
        End If
        R = Process32Next(hSnapShot, uProcess)
        DoEvents
    Loop
    CloseHandle hSnapShot
End If
End Function

Function TerminateExeName(ExeName As String) ' As Long
On Error GoTo ErrHandle
    
    Dim uProcess As PROCESSENTRY32
    Dim lProc As Long, hProcSnap As Long
    Dim ExePath As String
    Dim hPID As Long, hExit As Long
    Dim I As Integer

    uProcess.dwSize = Len(uProcess)
    hProcSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    lProc = Process32First(hProcSnap, uProcess)
    Do While lProc
        I = InStr(1, uProcess.szExeFile, Chr$(0))
        ExePath = UCase$(Left$(uProcess.szExeFile, I - 1))
        If UCase$(GetFileName(ExePath)) = UCase$(ExeName) Then
            hPID = OpenProcess(1&, -1&, uProcess.th32ProcessID)
            hExit = TerminateProcess(hPID, 0&)
            Call CloseHandle(hPID)
        End If
        lProc = Process32Next(hProcSnap, uProcess)
    Loop
    Call CloseHandle(hProcSnap)
    Exit Function
    
ErrHandle:
End Function

Function KillProcessById(p_lngProcessId As Long) As Long
On Error Resume Next
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    Dim hToken As Long
    Dim hProcess As Long
    Dim tp As TOKEN_PRIVILEGES

    If IsWinNT Then
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or _
            TOKEN_QUERY, hToken) = 0 Then
            CloseHandle hToken
        End If
        If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
            CloseHandle hToken
        End If
        tp.PrivilegeCount = 1
        tp.Attributes = SE_PRIVILEGE_ENABLED
        If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, _
           ByVal 0&) = 0 Then
            CloseHandle hToken
        End If
    End If
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    KillProcessById = lngReturn
End Function

Public Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Private Function GetSpecialfolder(CSIDL As Long) As String
On Error Resume Next
    Dim R As Long, path As String
    Dim IDL As ITEMIDLIST
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = 0 Then
        path = Space$(512)
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal path)
        GetSpecialfolder = Left$(path, InStr(path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Function ReplacePathSystem(np As String) As String
On Error Resume Next
Dim buff As String
buff = Replace(np, "\??\", "", , , vbTextCompare)
buff = Replace(buff, "\\?\", "", , , vbTextCompare)
buff = Replace(buff, "\SystemRoot\", MyWindowDir, , , vbTextCompare)
buff = Replace(buff, "%systemroot%", MyWindowDir, , , vbTextCompare)
buff = Replace(buff, "\\", "\", , , vbTextCompare)
ReplacePathSystem = buff
End Function

Public Function SetSuspendResumeThread(lvwProc As ListView, _
    ItemProcessID As Integer, SuspendNow As Boolean) _
    As Long
    
    Dim Thread() As THREADENTRY32, hPID As Long, _
        hThread As Long, I As Long
    
    hPID = lvwProc.SelectedItem.SubItems(ItemProcessID)
    
    Thread32_Enum Thread(), hPID
    
    For I = 0 To UBound(Thread)
        If Thread(I).th32OwnerProcessID = hPID Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, _
                False, (Thread(I).th32ThreadID))
            If SuspendNow Then
                SuspendThread hThread
            Else
                ResumeThread hThread
            End If
            CloseHandle hThread
        End If
    Next I
End Function

Public Function Thread32_Enum(ByRef Thread() As THREADENTRY32, Optional ByVal lProcessID As Long) As Long
On Error GoTo VB_Error
    ReDim Thread(0)
    
    Dim THREADENTRY32 As THREADENTRY32
    Dim hSnapShot As Long
    Dim lThread As Long
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, lProcessID)  ': 'If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot ::: INVALID_HANDLE_VALUE failed", sLocation, "Thread32_Enum")
    
    THREADENTRY32.dwSize = Len(THREADENTRY32)
    If Thread32First(hSnapShot, THREADENTRY32) = False Then
        Thread32_Enum = -1
        Exit Function
    Else
        ReDim Thread(lThread)
        Thread(lThread) = THREADENTRY32
    End If
    
    Do
        If Thread32Next(hSnapShot, THREADENTRY32) = False Then
            Exit Do
        Else
            lThread = lThread + 1
            ReDim Preserve Thread(lThread)
            Thread(lThread) = THREADENTRY32
        End If
    Loop
    Thread32_Enum = lThread
    
Exit Function
VB_Error:
Resume Next
End Function

Function GetUserNameA(sID As Long) As String
If IsWinNT Then

    On Error Resume Next
    Dim retname As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    LookupAccountSid vbNullString, sID, retname, 255, retdomain, 255, 0
    GetUserNameA = Left$(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & Left$(retname, InStr(retname, vbNullChar) - 1)
End If
End Function

Sub GetWTSProcesses(coll As Collection)
On Error Resume Next
Dim Retval As Long
Dim count As Long
Dim I As Integer
Dim lpBuffer As Long
Dim P As Long
Dim udtProcessInfo As WTS_PROCESS_INFO

If IsWinNT Then
    Retval = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, count)
    If Retval Then
       P = lpBuffer
         For I = 1 To count
             CopyMemory udtProcessInfo, ByVal P, LenB(udtProcessInfo)
             coll.Add GetUserNameA(udtProcessInfo.pUserSid), "#" & udtProcessInfo.ProcessID
             P = P + LenB(udtProcessInfo)
         Next I
         WTSFreeMemory lpBuffer   'Free your memory buffer
     End If
End If
End Sub

Public Function GetPriority(PID As Long)
Dim hWnd As Long, pri As Long
    hWnd = OpenProcess(PROCESS_QUERY_INFORMATION, False, PID)
    pri = GetPriorityClass(hWnd)
    CloseHandle hWnd
    GetPriority = pri
End Function

Public Sub ShutDownNT(Force As Boolean)
On Error Resume Next
    Dim Ret As Long
    Dim FLAGS As Long
    FLAGS = EWX_SHUTDOWN
    If Force Then FLAGS = FLAGS + EWX_FORCE
    If IsWinNT Then EnableShutDown
    ExitWindowsEx FLAGS, 0
End Sub

Public Sub RebootNT(Force As Boolean)
On Error Resume Next
    Dim Ret As Long
    Dim FLAGS As Long
    FLAGS = EWX_REBOOT
    If Force Then FLAGS = FLAGS + EWX_FORCE
    If IsWinNT Then EnableShutDown
    ExitWindowsEx FLAGS, 0
End Sub

Public Sub LogOffNT(Force As Boolean)
On Error Resume Next
    Dim Ret As Long
    Dim FLAGS As Long
    FLAGS = EWX_LOGOFF
    If Force Then FLAGS = FLAGS + EWX_FORCE
    ExitWindowsEx FLAGS, 0
End Sub

Private Sub EnableShutDown()
On Error Resume Next
    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES32
    Dim mNewPriv As TOKEN_PRIVILEGES32
    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID
    ' enable shutdown privilege for the current application
    AdjustTokenPrivileges32 hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)
End Sub

Function GetAttribute(ByVal sFilePath As String) As String
        
    Select Case GetFileAttributes(sFilePath)
        Case 1: GetAttribute = "R": Case 2: GetAttribute _
            = "H": Case 3: GetAttribute = "RH": Case 4: _
            GetAttribute = "S": Case 5: GetAttribute = _
            "RS": Case 6: GetAttribute = "HS": Case 7: _
            GetAttribute = "RHS"
        '-------------------------------------------------'
        Case 32: GetAttribute = "A": Case 33: GetAttribute _
            = "RA": Case 34: GetAttribute = "HA": Case 35: _
            GetAttribute = "RHA": Case 36: GetAttribute = _
            "SA": Case 37: GetAttribute = "RSA": Case 38: _
            GetAttribute = "HSA": Case 39: GetAttribute = _
            "RHSA"
        '-------------------------------------------------'
        Case 128: GetAttribute = "Normal"
        '-------------------------------------------------'
        Case Else: GetAttribute = "N/A"
    End Select

End Function

Function Hex2Str(inHexStr As String) As String
On Error GoTo salah
    If Trim(inHexStr) <> "" Then
       Dim I As Integer, buff As String
       For I = 1 To Len(inHexStr) Step 2
          buff = buff & Chr(Val("&H" & Mid(inHexStr, I, 2)))
       Next I
       Hex2Str = buff
    End If
Exit Function
salah:
End Function

Function GetMySetting(nApp As String, nKey As String, Optional nDefault As String)
On Error GoTo salah
Dim Ret As String
Ret = String(255, 0)
GetPrivateProfileString nApp, nKey, nDefault, Ret, 255, nPath(App.path) & "config.cfg"
GetMySetting = Left(Ret, InStr(1, Ret, Chr(0), vbTextCompare) - 1)
Exit Function
salah:
End Function

Function SetMySetting(nApp As String, nKey As String, nVal As String)
On Error GoTo salah
WritePrivateProfileString nApp, nKey, nVal, nPath(App.path) & "config.cfg"
Exit Function
salah:
End Function

Public Sub KeepOnTop(F As Form)

    Const SWP_NOMOVE   As Long = 2
    Const SWP_NOSIZE   As Long = 1
    Const HWND_TOPMOST As Long = -1

    SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Public Sub BeepNow()
    Dim I As Integer
    For I = 1500 To 2000 Step 100
        Beep I, 20
    Next I
End Sub

Public Sub Finish()
    Beep 1000, 80
End Sub

Public Function M_Scan(A As Boolean)
    frmScanVirus.picScan.Enabled = A
    frmScanVirus.picScan.Enabled = A
End Function

Public Function M_Options(A As Boolean)
    frmScanVirus.picOptions.Enabled = A
    frmScanVirus.picOptions.Visible = A
End Function

Public Function M_Process(A As Boolean)
    frmScanVirus.picProcess.Enabled = A
    frmScanVirus.picProcess.Visible = A
End Function

Public Function M_Startup(A As Boolean)
    frmScanVirus.picStartup.Enabled = A
    frmScanVirus.picStartup.Visible = A
End Function

Public Function M_Tweak(A As Boolean)
    frmScanVirus.picTweak.Enabled = A
    frmScanVirus.picTweak.Visible = A
End Function

'Public Function M_About(A As Boolean)
'
'frmScanVirus.picAbout.Enabled = A
'frmScanVirus.picAbout.Visible = A
'
'End Function

Function KillBundleFile(fName As String, Mark As String, pos As Long, posAdd As Long) As Boolean
On Error GoTo salah
Dim lSize As Long

lSize = FileLen(fName) - (pos + posAdd)
If MatchFile(fName, Mark, pos) Then
   If ModifyFromFile(fName, pos, lSize) Then
      KillBundleFile = True
   Else
      KillBundleFile = False
      SetAttr fName, vbNormal + vbArchive
      FileDie fName
   End If
Else
  KillBundleFile = False
  SetAttr fName, vbNormal + vbArchive
  FileDie fName
End If
Exit Function
salah:
KillBundleFile = False

End Function

Function FileDie(nFileName As String) As Boolean
 On Error GoTo salah
 SetAttr nFileName, vbArchive + vbNormal
 Kill nFileName
 FileDie = True
 Exit Function
salah:

End Function

Function AddToColFileType(ID As String, col As Collection)
On Error GoTo salah
   Dim buff As String
   col.Add ID, "#" & ID
   AddToColFileType = True
   Exit Function
salah:
End Function

Function file_getTitle(FileName As String) As String
    Dim buffer() As String
    If InStr(1, FileName, ".", vbTextCompare) > 0 Then
       buffer = Split(FileName, ".")
       If UBound(buffer) > 0 Then
          file_getTitle = buffer(UBound(buffer))
       End If
    End If
End Function

Function MatchFile(fName As String, Mark As String, Optional PosFile As Long = -1) As Boolean
On Error GoTo salah
    Dim I As Integer
    Dim hHex() As String
    Dim tmp As String
    hHex() = Split(Mark, " ")
    
    Dim data() As Byte
    ReDim data(UBound(hHex)) As Byte
    
    If PosFile > 0 Then
       Open fName For Binary Access Read As #1
           Get #1, PosFile, data
       Close #1
       For I = 0 To UBound(data)
            tmp = tmp & String(2 - Len(Hex(data(I))), "0") & Hex(data(I)) & " "
       Next I
       tmp = IIf(Right(tmp, 1) = " ", Left(tmp, Len(tmp) - 1), tmp)
       If tmp = Mark Then
          MatchFile = True
       End If
    Else
    End If
    Exit Function
salah:
Close #1

End Function

Function ModifyFromFile(fName As String, starpos As Long, size As Long, Optional newName As String = "") As Boolean
On Error GoTo salah
    Dim data() As Byte
    ReDim data(size) As Byte
    
    Open fName For Binary Access Read As #1
        Get #1, starpos, data
    Close #1
    
    If newName = "" Then
       SetAttr fName, vbNormal + vbArchive
       Kill fName
       Open fName For Binary As #1
           Put #1, , data
       Close #1
       ModifyFromFile = True
    Else
       Open newName For Binary As #1
           Put #1, , data
       Close #1
       ModifyFromFile = True
    End If
    Exit Function
salah:
Close #1

End Function


