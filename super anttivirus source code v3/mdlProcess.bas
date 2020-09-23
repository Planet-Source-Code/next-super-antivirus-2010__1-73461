Attribute VB_Name = "mdlProcess"
Option Explicit

Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE

Dim myProcess  As Collection

Sub ViewProcess()
On Error Resume Next
 Dim hSnapShot As Long, uProcess As PROCESSENTRY32
 Dim namafile As String, lngModules(1 To 200) As Long
 Dim strModuleName As String, Xproses As Long
 Dim enumerasi As Long, strProcessName As String
 Dim lngSize As Long
 Dim lngReturn  As Long
 Dim hFile As String
 Dim C As New clsCRC32
 Dim memUsage As PROCESS_MEMORY_COUNTERS

    frmScanVirus.lstView.ListItems.Clear
    frmScanVirus.ImageList3.ListImages.Clear
    
    Set myProcess = Nothing
    Set myProcess = New Collection
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    enumerasi = Process32First(hSnapShot, uProcess)
    lngSize = 500
    strModuleName = Space(MAX_PATH)
    Dim data(10) As String
    
    Dim col As New Collection
    Dim pos As Long
    pos = 1
    GetWTSProcesses col 'Get user name
        
    Do While enumerasi
    DoEvents
        Xproses = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, uProcess.th32ProcessID)
        lngReturn = GetModuleFileNameExA(Xproses, lngModules(1), strModuleName, lngSize)
        strProcessName = ReplacePathSystem(Left(strModuleName, lngReturn))
        If strProcessName <> "" Then
           namafile = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
           
           Dim myUsername As String
           If col.count > 0 Then
              On Error Resume Next
              myUsername = col("#" & uProcess.th32ProcessID)
           End If
           
           hFile = strProcessName
           Dim H As VERHEADER, buff As String
            GetVerHeader hFile, H
            With H
                 buff = "File Version: " & H.FileVersion & vbCrLf
            End With
            
           Dim lst As ListItem
           DrawIconFromFile strProcessName, "#" & namafile
           
           Set lst = frmScanVirus.lstView.ListItems.Add(, , LCase(namafile), , "#" & namafile)
               lst.SubItems(1) = strProcessName
               lst.SubItems(2) = myUsername
               lst.SubItems(3) = H.FileDescription
               lst.SubItems(4) = FileLen(strProcessName) \ 1024 & " KB"
               lst.SubItems(5) = uProcess.th32ProcessID
               lst.SubItems(6) = uProcess.pcPriClassBase
               lst.SubItems(7) = uProcess.cntThreads
               lst.SubItems(8) = GetAttribute(strProcessName)
               lst.SubItems(9) = GetBasePriority(uProcess.th32ProcessID)
               lst.SubItems(10) = C.FileChecksum(strProcessName)
               lst.SubItems(11) = Format(GetMemory(uProcess.th32ProcessID) \ 1024, "###,###") & " KB"
               lst.Tag = uProcess.th32ProcessID
            
           pos = pos + 1
        End If

        enumerasi = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
    Set col = Nothing
    frmScanVirus.StatusBar1.Panels(2).Text = "Processes: " & pos
End Sub

Sub DrawIconFromFile(FileName As String, hkey As String)
On Error Resume Next
    Dim mIcon As Long, Cnt As Long, sIcon As Long
    frmScanVirus.picIcon.Cls
    Set frmScanVirus.picIcon.Picture = Nothing
    If ExtractIconEx(FileName, 0, mIcon, sIcon, 1) > 0 Then
       DrawIconEx frmScanVirus.picIcon.hdc, 0, 0, sIcon, 0, 0, 0, 0, DI_IMAGE
       DestroyIcon mIcon
       frmScanVirus.ImageList3.ListImages.Add , hkey, frmScanVirus.picIcon.Image
    Else
       frmScanVirus.ImageList3.ListImages.Add , hkey, frmScanVirus.ImgIcon.Picture
    End If
    
End Sub

Sub GetIconFromFile(FileName As String, hdc As PictureBox)
On Error Resume Next
    Dim mIcon As Long, Cnt As Long
    hdc.Cls
    Set hdc.Picture = Nothing
    If ExtractIconEx(FileName, 0, mIcon, ByVal 0&, 1) > 0 Then
       DrawIconEx hdc.hdc, 0, 0, mIcon, 0, 0, 0, 0, DI_NORMAL
       DestroyIcon mIcon
    Else
       frmScanVirus.picIcon.PaintPicture frmScanVirus.ImgIcon.Picture, 0, 0
    End If
End Sub

Private Function GetBasePriority(ReadPID As Long) As String
Dim hPID As Long
    
    hPID = OpenProcess(PROCESS_QUERY_INFORMATION, 0, ReadPID)
    
    Select Case GetPriorityClass(hPID)
        Case 32: GetBasePriority = "Normal"
        Case 64: GetBasePriority = "Idle"
        Case 128: GetBasePriority = "High"
        Case 256: GetBasePriority = "Realtime"
        Case Else: GetBasePriority = "N/A"
    End Select
    
    CloseHandle hPID
End Function

Public Function ScanProcess(Jawab As Boolean)

    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    Dim namafile As String, lModule(1 To 200) As Long
    Dim strModuleName As String, xProcess As Long
    Dim enumerasi As Long, strProcessName As String
    Dim lSize As Long, lReturn As Long
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    enumerasi = Process32First(hSnapShot, uProcess)
    lSize = 500
    strModuleName = Space(MAX_PATH)
    Dim data(1) As String
    
    Do While enumerasi
        xProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, uProcess.th32ProcessID)
        lReturn = GetModuleFileNameExA(xProcess, lModule(1), strModuleName, lSize)
        strProcessName = ReplacePathSystem(Left(strModuleName, lReturn))
        If strProcessName <> "" Then
            namafile = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
            frmStartup.lblProcess = namafile
            ScanVirus strProcessName, frmStartup, frmScanVirus.lstDetection
            DrawIconFromFile strProcessName, "#" & namafile
        End If
        
        enumerasi = Process32Next(hSnapShot, uProcess)
    Loop
    
    CloseHandle hSnapShot
       
    If Jawab = True Then
        MsgBox "Scanning Process finished", vbInformation + vbOKOnly, "Scanning Process"
    End If
    
End Function
