Attribute VB_Name = "mdlFileFunction"
Option Explicit

Private Const HWND_NOTOPMOST = -2

Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
    
Type VERHEADER
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    Comments As String
    LegalTradeMarks As String
    PrivateBuild As String
    SpecialBuild As String
End Type

'Get icon
Private shinfo As SHFILEINFO, sshinfo As SHFILEINFO
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hdcDest&, ByVal X&, ByVal Y&, ByVal FLAGS&) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Enum IconRetrieve
    ricnLarge = 32
    ricnSmall = 16
End Enum

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private SIconInfo As SHFILEINFO

Public Function GetVerHeader(ByVal fPN$, ByRef oFP As VERHEADER)
On Error Resume Next
Dim lngBufferlen&, lngDummy&, lngRc&, lngVerPointer&, lngHexNumber&, I%
Dim bytBuffer() As Byte, bytBuff(255) As Byte, strBuffer$, strLangCharset$, strVersionInfo(11) As String, strTemp$
 If Dir(fPN$, vbHidden + vbArchive + vbNormal + vbReadOnly + vbSystem) = "" Then
    oFP.CompanyName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.FileDescription = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.FileVersion = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.InternalName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.LegalCopyright = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.OrigionalFileName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.ProductName = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.ProductVersion = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.Comments = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.LegalTradeMarks = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.PrivateBuild = "The file """ & GetShortPath(fPN) & """ N/A"
    oFP.SpecialBuild = "The file """ & GetShortPath(fPN) & """ N/A"
    Exit Function
 End If
   lngBufferlen = GetFileVersionInfoSize(fPN$, lngDummy)
    If lngBufferlen > 0 Then
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(fPN$, 0&, lngBufferlen, bytBuffer(0))
       If lngRc <> 0 Then
        lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
         If lngRc <> 0 Then
          MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
           lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
            strLangCharset = Hex(lngHexNumber)
             Do While Len(strLangCharset) < 8
              strLangCharset = "0" & strLangCharset
             Loop
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             strVersionInfo(8) = "Comments"
             strVersionInfo(9) = "LegalTrademarks"
             strVersionInfo(10) = "PrivateBuild"
             strVersionInfo(11) = "SpecialBuild"
            For I = 0 To 11
               strBuffer = String$(255, 0)
               strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(I)
               lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   lstrcpy strBuffer, lngVerPointer
                   strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                   strVersionInfo(I) = strBuffer
                Else
                  strVersionInfo(I) = "N/A"
                   End If
            Next I
          End If
       End If
    End If
     For I = 0 To 11
      If Trim(strVersionInfo(I)) = "" Then strVersionInfo(I) = "N/A"
     Next I
    oFP.CompanyName = strVersionInfo(0)
    oFP.FileDescription = strVersionInfo(1)
    oFP.FileVersion = strVersionInfo(2)
    oFP.InternalName = strVersionInfo(3)
    oFP.LegalCopyright = strVersionInfo(4)
    oFP.OrigionalFileName = strVersionInfo(5)
    oFP.ProductName = strVersionInfo(6)
    oFP.ProductVersion = strVersionInfo(7)
    oFP.Comments = strVersionInfo(8)
    oFP.LegalTradeMarks = strVersionInfo(9)
    oFP.PrivateBuild = strVersionInfo(10)
    oFP.SpecialBuild = strVersionInfo(11)
End Function

Public Sub RetrieveIcon(fName As String, DC As PictureBox, icnSize As IconRetrieve)
    Dim hImgLarge As Long  'the handle to the system image list
        
    If icnSize = ricnLarge Then
        hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC.hdc, 0, 0, ILD_TRANSPARENT)
    End If
    
End Sub

Private Function GetShortPath(strFileName As String) As String
    Dim lngRes&, strPath$: strPath = String$(MAX_PATH, 0)
    lngRes = GetShortPathName(strFileName, strPath, MAX_PATH)
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Sub ShowProps(FileName As String, OwnerhWnd As Long)
On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = App.hInstance
        .lpIDList = 0
    End With
     ShellExecuteEx SEI
End Sub

Public Function GetFilePath(ByVal sPath As String) _
    As String
    
    Dim I As Integer
    
    For I = Len(sPath) To 1 Step -1
        If Mid$(sPath, I, 1) = "\" Then
            GetFilePath = Mid$(sPath, 1, I)
            Exit For
        End If
    Next I
End Function

Public Function GetPathType(ByVal path As String) As String
    
    Dim FileInfo As SHFILEINFO, lngRet As Long
    
    lngRet = SHGetFileInfo(path, 0, FileInfo, Len(FileInfo), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME)
    If lngRet = 0 Then GetPathType = Trim$(GetFileExtension(path) & " File"): Exit Function
    GetPathType = Left$(FileInfo.szTypeName, InStr(1, FileInfo.szTypeName, vbNullChar) - 1)
    
End Function

Public Function GetFileExtension(ByVal path As String) As String
    
    Dim intRet As Integer: intRet = InStrRev(path, ".")
    
    If intRet = 0 Then Exit Function
    GetFileExtension = UCase(Mid$(path, intRet + 1))
    
End Function

Public Function FileParsePath(sPathname As String, bRetFile As Boolean, bExtension As Boolean) As String
    Dim sEditArray() As String
    sEditArray = Split(sPathname, "\", -1)
    If bRetFile = True Then
        Dim sFilename As String
        sFilename = sEditArray(UBound(sEditArray))
        If bExtension = True Then
            FileParsePath = sFilename
        Else
            sEditArray = Split(sFilename, ".vir", -1)
            FileParsePath = sEditArray(LBound(sEditArray))
        End If
    Else
        Dim sPathnameA As String
        Dim I As Integer
        For I = 0 To UBound(sEditArray) - 1
            sPathnameA = sPathnameA & sEditArray(I) & "\"
        Next
        FileParsePath = sPathnameA
    End If
    On Error GoTo 0
End Function

Public Function GetSpecialfolder(CSIDL As sFolder) As String
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function

Public Function GenerateRandomTitle() As String

    Dim sTitle() As Variant
    
    sTitle = Array("a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "I", _
        "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    Randomize
    
    GenerateRandomTitle = sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * _
        UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * _
        UBound(sTitle)) & sTitle(Rnd * UBound(sTitle)) & sTitle(Rnd * UBound(sTitle))
    GenerateRandomTitle = EncryptText(GenerateRandomTitle)
        

End Function

Private Function EncryptText(ByVal sText As String) As String
    
    Dim intLen As Integer
    Dim sNewText As String
    Dim sChar As String
    Dim I As Integer
    
    sChar = ""
    intLen = Len(sText)
    For I = 1 To intLen
        sChar = Mid(sText, I, 1)
        Select Case Asc(sChar)
            Case 65 To 90: sChar = Chr(Asc(sChar) + 127)
            Case 97 To 122: sChar = Chr(Asc(sChar) + 121)
            Case 48 To 57: sChar = Chr(Asc(sChar) + 196)
            Case 32: sChar = Chr(32)
        End Select
        sNewText = sNewText + sChar
    Next I
    EncryptText = sNewText
    
    Exit Function
    
End Function

Public Sub AlwaysOnTop(ByVal hwndOwner As Long, ByVal SetOnTop As Boolean)

    If SetOnTop Then
        SetWindowPos hwndOwner, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
    Else
        SetWindowPos hwndOwner, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    End If
    
End Sub



