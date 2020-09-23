Attribute VB_Name = "mdlScanEngine"
Option Explicit

Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public VirType As String
Public nameDetect As String
Public VirSize As String

Dim CRC32 As New clsCRC32
Dim Heuristic As New clsHeuristic

Public VirusName As New Collection
Public VirusSign As New Collection

Public IconName As New Collection
Public IconSign As New Collection

Dim PatternCount As Integer
Dim PatternVirus(1000) As String
Dim VirusDetect As String

Private Function FindVirus(path As String, SearchStr As String, FileCount As Long, DirCount As Long)
On Error Resume Next

Dim FileName As String
Dim FindFiles As Long
Dim NAMA_DIRECTORY As String
Dim DIR_NAMES() As String
Dim nDir As Long
Dim I As Long
Dim buff As Long

If Right(path, 1) <> "\" Then path = path & "\"
    nDir = 0
ReDim DIR_NAMES(nDir)
    NAMA_DIRECTORY = Dir(path, vbDirectory Or vbHidden Or vbSystem Or vbReadOnly)

Do While Len(NAMA_DIRECTORY) > 0
    If (NAMA_DIRECTORY <> ".") And (NAMA_DIRECTORY <> "..") Then
        If GetAttr(path & NAMA_DIRECTORY) Or FILE_ATTRIBUTE_DIRECTORY Then
            DIR_NAMES(nDir) = NAMA_DIRECTORY
            DirCount = DirCount + 1
            nDir = nDir + 1
    
            ReDim Preserve DIR_NAMES(nDir)
        End If
    End If
    NAMA_DIRECTORY = Dir()
Loop
If StopScan = True Then Exit Function
    While PauseScan = True
        DoEvents
    Wend
    FileName = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive)

While Len(FileName) <> 0
    If StopScan = True Then Exit Function
        While PauseScan = True
            DoEvents
        Wend
    FindFiles = FindFiles + FileLen(path & FileName)
    FileCount = FileCount + 1
    
    '----------------- module Buffering -----------------
    If Buffering = True Then
        frmScanVirus.txtStatus.ForeColor = &H0&
        frmScanVirus.txtStatus.Text = "STATUS : [ Buffering, please wait ... ]"
        JumlahBuffer = FileCount
    Else
        If Unhide = True Then
                SetAttr path, vbNormal And vbDirectory
                SetAttr path & FileName, vbNormal
        End If
        frmScanVirus.lblFileName.Caption = path & FileName
        Persen = FileCount / JumlahBuffer * 100
        frmScanVirus.ProgScan.value = Persen
        JumlahFile = FileCount
        frmScanVirus.lblFileScan.Caption = ": " & JumlahFile
        frmScanVirus.lblPercen.Caption = Persen & " %" & " " & "Completed."
        Namaku = ": " & FileName
        ScanVirus frmScanVirus.lblFileName.Caption, frmScanVirus, frmScanVirus.lstDetection

        Dim cek As Long
        With frmScanVirus.lstDetection.ListItems
            For cek = 1 To .count
                If .Item(cek).Index <> 0 Then
                    .Item(cek).Checked = False
                End If
            Next cek
        End With

     '----------------- module scanning -----------------------
    End If
    DoEvents
    FileName = Dir()
Wend
If nDir > 0 Then
    For I = 0 To nDir - 1
        FindVirus = FindVirus + FindVirus(path & DIR_NAMES(I) & "\", SearchStr, FileCount, DirCount)
    Next I
    DoEvents
End If
End Function

Public Function Looping()
Dim SearchPath As String, FindStr As String
Dim FileSize As Long
Dim NumFiles As Long, NumDirs As Long

SearchPath = Where
FindStr = ext
FileSize = FindVirus(SearchPath, FindStr, NumFiles, NumDirs)
DoEvents
    
End Function

Public Function DeleteIt(whereit As String)

SetFileAttributes whereit, FILE_ATTRIBUTE_NORMAL
TerminateExeName GetFileName(whereit)
Kill (whereit)
If IsFileExist(whereit) = True Then
    Call MsgBox("File can't be deleted!", vbCritical + vbOKOnly, "Error Detected !")
End If
End Function

Private Function IsFileExist(ByVal sPath As String) As Boolean
    
If PathFileExists(sPath) = 1 And PathIsDirectory(sPath) = 0 Then
    IsFileExist = True
Else
    IsFileExist = False
End If
    
End Function

Public Function GetFileName(sFilename As String) As String

Dim buffer As String

buffer = String(255, 0)
GetFileTitle sFilename, buffer, Len(buffer)
buffer = StripNulls(buffer)
GetFileName = buffer
    
End Function

Public Function StripNulls(OriginalStr As String) As String

If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If

StripNulls = OriginalStr
End Function

Public Function ScanVirus(Dimana As String, FormMe As Form, lDetection As ListView)

Dim Where As String

Where = Dimana

     If IsVirus(Dimana) Then
        With frmScanVirus.lstDetection.ListItems.Add(, , VirusDetect)
            .SubItems(1) = Dimana
            .SubItems(2) = Format(FileLen(Where) / 1024, "###,####") & " KB"
            .SubItems(3) = GetAttribute(Dimana)
            If VirusDetect = TipeHeuristic Then
                .SubItems(4) = "Virus Suspected with heuristic icon!"
            End If
            VirusDetect = VirusDetect
            VirusDetected = VirusDetected + 1
            frmScanVirus.lblVirusDetected.Caption = ": " & VirusDetected
            If DeleteAll = True Then
                DeleteIt (Dimana)
                VirusCleaned = VirusCleaned + 1
                frmScanVirus.lblVirusClean.Caption = ": " & VirusCleaned
            End If
                LogFile "Virus found  " & Dimana & "     Detected As: " & VirusDetect
        End With
    End If
End Function

Public Function IsVirus(strFileName As String) As Boolean
    On Error Resume Next
    
    Dim strCRC As String
    Dim I As Long
    Dim j As Long
    Dim lv As ListItems
    
    strCRC = GetChecksum(strFileName)
    For I = 1 To VirusSign.count
        If strCRC = VirusSign.Item(I) Then
            IsVirus = True
            VirSize = Int((FileLen(strFileName) / 1024) * 100 + 0.5) / 100 & " KB" 'FileLen(strFile) & " Bytes"
            VirusDetect = VirusName.Item(I)
            Exit Function
        End If
    Next I
    
    If Heuristic.CekHeuristic(strFileName) = True Then
        IsVirus = True
        VirusDetect = TipeHeuristic
        Exit Function
    End If
        
End Function

Public Function ProcedureScan()
    Dim RegistryFix As Boolean
        
    Tutup
    frmScanVirus.lblExt.Caption = ext
    If DisBuffer = False Then
        frmScanVirus.ProgScan.TextStyle = CustomText
        GoTo NonBuffer
    End If
    
    frmScanVirus.ProgScan.TextStyle = PBNoneText
    frmScanVirus.tmrStatus.Enabled = True
        
    Buffering = True
        Looping 'for buffering
        frmScanVirus.txtStatus.ForeColor = &HFF0000
    Buffering = False
    
NonBuffer:
        frmScanVirus.txtStatus.Text = "STATUS : [ Scanning files ... ]"
        Looping 'for scanning
        If RegistryFix = True Then
            frmScanVirus.txtStatus.Text = "STATUS : [ FIXING REGISTRY ]"
        End If
        frmScanVirus.tmrStatus = False
        If MsgBox("Scanning Complete." & vbCrLf & vbCrLf & _
            " - Scanned File(s)  : " & JumlahFile & vbCrLf & _
            " - Object Detected : " & VirusDetected & vbCrLf & _
            " - Object Fixed / Cleaned : " & VirusCleaned & vbCrLf & _
            "  " & vbCrLf & _
            " - Total Scanning : " & frmScanVirus.lblTimeValue.Caption, vbInformation + vbOKOnly, APP_PROGRAM) = vbOK Then
            StopScan = True
            frmScanVirus.cmdStop.Enabled = False
            frmScanVirus.cmdPause.Enabled = False
            Buka
        End If
End Function

Public Function ExternalDatabase()
Dim nf As Integer
Dim cVDF As String
Dim cPattern As String

cVDF = App.path + "\TCM.VDB"
nf = FreeFile
Open cVDF For Append As #nf
    cPattern = SignTemp
    If cPattern = "" Then
        Call MsgBox(APP_PROGRAM & "Error Handle" & vbCrLf & "Error Detected :" & vbCrLf & vbCrLf & " - " & "Error Make External Virus DB", vbCritical + vbOKOnly, "Critical Error")
    Else
        Print #nf, cPattern
    End If
Close #nf
    Call MsgBox("Virus Database updated successfully!", vbOKOnly + vbInformation, "Add Virus Signature")
End Function

Public Function LoadExternalDatabase(Pertamakali As Boolean)

    On Error Resume Next
    Dim dbFile As String
    
    dbFile = App.path & "\TCM.VDB"
    
    If IsFileExist(dbFile) = False Then
        Exit Function
    End If
    
    Dim sData As String
    Open dbFile For Binary Access Read As #1
        sData = String(LOF(1), Chr(0))
        Get #1, , sData
    Close #1
    
    Dim strArray() As String
    strArray = Split(sData, vbCrLf)
    
    Dim I As Long, j As Long
    
    If Pertamakali = True Then
        For I = 0 To UBound(strArray)
            Dim cVirus() As String
            cVirus = Split(strArray(I), ";")
            VirusName.Add cVirus(0)
            VirusSign.Add cVirus(1)
        Next I
    Else
        VirusName.Add VirType & UCase(frmSignature.txtVirusName.Text)
        VirusSign.Add frmSignature.lblCRCSTR.Caption
    End If

End Function

Private Function Tutup()

StopScan = False
StopButton = True
With frmScanVirus
    .cmdScan.Enabled = False
    .cmdBrowse.Enabled = False
    .lstDetection.Enabled = False
End With

End Function

Public Function Buka()

StopScan = True
StopButton = False
    With frmScanVirus
        .cmdScan.Enabled = True
        .cmdBrowse.Enabled = True
        .lstDetection.Enabled = True
        .tmrStatus.Enabled = False
        .txtStatus.ForeColor = &H80000012
        .txtStatus.Text = "STATUS : [ Waiting For Instructions ]"
        If VirusDetected <> 0 Then
            .txtStatus.ForeColor = &HFF&
            .txtStatus.Text = "STATUS : Infected." & " [ Go to (tweak registry) menu to fix your system / registry ]"
        Else
            .txtStatus.ForeColor = &H80000012
            .txtStatus.Text = "STATUS : [ Waiting For Instructions ]"
        End If
        .lblFileName.Caption = "||----- "
        .lblFileScan = ": 0"
        .ProgScan.value = "0"
        .lblExt.Caption = "-"
        .lblTimeValue.Caption = ": 00:00:00"
        .lblPercen.Caption = "0" & " %" & " " & "Completed."
        JumlahBuffer = 0
        JumlahFile = 0
        VirusDetected = 0
        VirusCleaned = 0
End With
End Function

Public Sub VirusAlert()
    Dim I As Integer
    
    For I = 1500 To 2000 Step 100
        Beep I, 20
    Next I
End Sub

Public Sub ScanFinish()
    Beep 1800, 50
    Sleep 20
    Beep 1800, 100
End Sub
