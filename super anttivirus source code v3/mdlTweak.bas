Attribute VB_Name = "mdlTweak"
'Av Super Protector - A N T I V I R U S
'Current Build This Module 9 Agustus 2008
'Version 1.1 Beta
'Copyright (c) 2008 Moh Aly Shodiqin (fidly)
'Av Super Protector Software Studio
'----------------------------------------------

Option Explicit

Public Const rExplorer = "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
            rWapp = "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", _
            rInternet = "Software\Policies\Microsoft\Internet Explorer\Restrictions", _
            rSystem = "Software\Microsoft\Windows\CurrentVersion\Policies\System", _
            rNetwork = "Software\Microsoft\Windows\CurrentVersion\Policies\Network", _
            rDesktop = "Control Panel\Desktop", _
            rAdvanced = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

Dim REG As New clsRegistry

Sub GetApp()
    On Error Resume Next
    
    Dim I As Integer, Isi As String, tmp
    
    With frmScanVirus
    For I = 0 To .chkT.count - 1
       Isi = Trim(.chkT(I).Tag)
       Select Case I
            Case 0, 1, 7, 8 To 11
                tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rSystem, Isi)
                tmp = REG.GetSettingLong(HKEY_LOCAL_MACHINE, rSystem, Isi)
            Case 2 To 5, 12, 13, 17 To 20
                tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rExplorer, Isi)
                tmp = REG.GetSettingLong(HKEY_LOCAL_MACHINE, rExplorer, Isi)
            Case 6
                tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rDesktop, Isi)
            Case 14
                tmp = REG.GetSettingLong(HKEY_CURRENT_USER, rAdvanced, Isi)
            Case 15
                tmp = REG.GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi)
                If Trim(tmp) <> 1 Then
                    tmp = 0
                Else
                    tmp = 1
                End If
            Case 16
                tmp = REG.GetSettingLong(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi)
                If Trim(tmp) = 0 Then
                    tmp = 1
                Else
                    tmp = 0
                End If
        End Select
            .chkT(I).value = Val(tmp)
       DoEvents
    Next I
    End With

End Sub

Sub SaveApp()
    On Error Resume Next
    
    Dim I As Integer, Isi As String
    
    With frmScanVirus
    For I = 0 To .chkT.count - 1
       Isi = Trim(.chkT(I).Tag)
       Select Case I
            Case 0, 1, 7, 8 To 11
                CekReg .chkT(I).value, HKEY_CURRENT_USER, rSystem, Isi, 1
                CekReg .chkT(I).value, HKEY_LOCAL_MACHINE, rSystem, Isi, 1
            Case 2 To 5, 12, 13, 17 To 20
                CekReg .chkT(I).value, HKEY_CURRENT_USER, rExplorer, Isi, 1
                CekReg .chkT(I).value, HKEY_LOCAL_MACHINE, rExplorer, Isi, 1
            Case 6
                CekReg .chkT(I).value, HKEY_CURRENT_USER, rDesktop, Isi, 1
            Case 14
                CekReg .chkT(I).value, HKEY_CURRENT_USER, rAdvanced, Isi, 1
            Case 15
                If .chkT(15).value = 1 Then
                    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi, 1
                Else
                    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi, 2
                End If
            Case 16
                If .chkT(16).value = Isi Then
                    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi, 1
                Else
                    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", Isi, 1
                End If
       End Select
       DoEvents
    Next I
    End With

End Sub

Function CekReg(Nm As Boolean, Root As Long, path As String, value As String, Tipe As Byte)
    On Error Resume Next
    
    If Nm = True Then
       Select Case Tipe
              Case 1
                    REG.SaveSettingLong Root, path, value, 1
              Case 2
                    REG.SaveSettingByte Root, path, value, 1
              Case 3
                    REG.SaveSettingString Root, path, value, 1
      End Select
    Else
       REG.DeleteValue Root, path, value
    End If
    
End Function

Public Function FixRegistry()

    ComName = NameOfTheComputer(PCName)
    strUserCom = GetUserCom()

    On Error Resume Next
    
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "piffile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "cmdfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "scrfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    REG.SaveSettingString HKEY_CLASSES_ROOT, "regfile\shell\open\command", "", "regedit.exe %1"
    
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "HideFileExt", 1
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Hidden", 2
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "ShowSuperHidden", 0
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", "0"
    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "@shell32.dll,-30508", "Hide protected operating system files (Recommended)"
    REG.SaveSettingString HKEY_USERS, "S-1-5-21-1417001333-1060284298-725345543-500\Software\Microsoft\Windows\ShellNoRoam\MUICache", "@shell32.dll,-30508", "Hide protected operating system files (Recommended)"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden", "Text", "@shell32.dll,-30508"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden", "Bitmap", "%SystemRoot%\system32\SHELL32.dll,4"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\SHOWALL", "Text", "@shell32.dll,-30500"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\SHOWALL", "Type", "radio"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", strUserCom
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", PCName

'    ------------------------------- Reg W32/Xeror--------------------------------------
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden", "Type", "checkbox"
    REG.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations"
    REG.SaveSettingString HKEY_USERS, "S-1-5-21-839522115-2052111302-2147137731-500\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "LastKey", "My Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden"
    REG.DeleteValue HKEY_USERS, "S-1-5-21-839522115-2052111302-2147137731-500\Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\WINDOWS\system32\ctfmon.exe"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "ctfmon.exe"
'    ------------------------------ Akhir Reg W32/Xeror----------------------------------

'    ------------------------------- Reg 4k51k4--------------------------------------
'   Disable CMD
    REG.DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    REG.DeleteKey HKEY_USERS, "S-1-5-21-1547161642-1343024091-725345543-500\Software\Policies\Microsoft\Windows\System"

'   Disable System Restore
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"
    REG.SaveSettingLong HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    REG.SaveSettingString HKEY_CURRENT_USER, "Control Panel\Desktop\", "SCRNSAVE.EXE", "C:\WINDOWS\System32\logon.scr"
    REG.DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp"

'   Show Full Path at Address Bar
    REG.SaveSettingLong HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState", "FullPathAddress", 1

'   atur registy agar virus dapat berjalan pada saat login
'    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & GetUserAktif & GetLocalSettingsUser & "\Application Data\WINDOWS\CSRSS.EXE"
'    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & GetLocalSettingsUser & "\Application Data\WINDOWS\LSASS.EXE"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe "
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", MyWindowSys & "userinit.exe"

    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "4k51k4"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "MSMSGS"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Service" & strUserCom
    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Logon" & strUserCom
    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "System Monitoring"
'    ------------------------------ Akhir Reg 4k51k4----------------------------------

'    ------------------------------ Reg Worm32.walpaper----------------------------------
    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "ssvchost"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Print Epson"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Walpaper"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\lutil\FMR", "svchost"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "Software\lutil\FMR", "Register"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Group Policy", "svchost"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Group Policy", "AppMgmt"

    Kill GetSpecialfolder(CSIDL_MYPICTURES) & "\XxXLove.exe"
    Kill GetSpecialfolder(CSIDL_LOCAL_APPDATA) & "\MakeLove.exe"
    Kill GetSpecialfolder(CSIDL_STARTUP) & "\Fuckme.com"
    Kill GetSpecialfolder(CSIDL_APPDATA) & "\TuGas.exe"
'    ------------------------------ Akhir Reg Worm32.walpaper----------------------------------

'    ------------------------------ Reg Virus Amburadul.Hokage Killer----------------------------------

    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PaRaY_VM"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ConfigVir"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NviDiaGT"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "NarmonVirusAnti"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AVManager"
    REG.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title"

    REG.DeleteValue HKEY_LOCAL_MACHINE, " SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msconfig.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\rstrui.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\wscript.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\mmc.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Install.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\procexp.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\msiexec.exe"
    REG.DeleteValue HKEY_CLASSES_ROOT, "exefile", "NeverShowExt"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\taskkill.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\cmd..exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\tasklist.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\HokageFile.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Rin.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\Obito.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\KakashiHatake.exe"
    REG.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\HOKAGE4.exe"
    
'    ------------------------------ Akhir Reg Virus Amburadul.Hokage Killer----------------------------------

'    ------------------------------ Reg Flu_Ikan----------------------------------
'    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "NV Hostname", ""
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "kebodohan"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "pemalas"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "mulut_besar"
    REG.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "otak_udang"

    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Start Page", "http://www.microsoft.com/isapi/redir.dll?prd={SUB_PRD}&clcid={SUB_CLSID}&pver={SUB_PVER}&ar=home"

'    Buat agar tidak bisa masuk Safe Mode
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmboot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmio.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\dmload.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\sermouse.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\sr.sys", "", "FSFilter System Recovery"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\vga.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Minimal\vgasave.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\dmboot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\dmiot.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpcdd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpdd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\rdpwd.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\sermouse.sys", "", "Driver"
'    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\sr.sys", "", "FSFilter System Recovery"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\tdpipe.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\tdtcp.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\vga.sys", "", "Driver"
    REG.SaveSettingString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\Network\vgasave.sys", "", "Driver"

'    Ganti Nama User
'    REG.DeleteKey HKEY_CURRENT_USER, "Software\Microsoft\MS Setup (ACME)\User Info"
'    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\MS Setup (ACME)\User Info", "DefCompany"
'    REG.SaveSettingString HKEY_CURRENT_USER, "Software\Microsoft\MS Setup (ACME)\User Info", "DefName"
'    REG.DeleteKey HKEY_USERS, "S-1-5-21-2025429265-527237240-725345543-1003\Software\Microsoft\MS Setup (ACME)\User Info"
'    REG.DeleteValue HKEY_USERS, "S-1-5-21-2025429265-527237240-725345543-1003\Software\Microsoft\MS Setup (ACME)\User Info", "DefCompany"
'    REG.DeleteValue HKEY_USERS, "S-1-5-21-2025429265-527237240-725345543-1003\Software\Microsoft\MS Setup (ACME)\User Info", "DefName"
'    ------------------------------ Akhir Reg Flu_Ikan----------------------------------

    DoEvents
    
End Function
