VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeuristic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long

Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

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

Dim I As Integer
Dim j As Integer


Public Function CekHeuristic(Alamatfile As String) As Boolean
On Error GoTo keluar
Dim tmp As String
    
    CekHeuristic = False
    If (Right(Alamatfile, 3) = "exe") Then
        If CekIconBinary(Alamatfile) = True Then CekHeuristic = True
    End If
    Exit Function
    
keluar:
    CekHeuristic = False
End Function

Private Function CekIconBinary(PathFile As String) As Boolean
Dim q As Integer
Dim IconIDNow As String
   
    IconIDNow = CalcIcon(PathFile)
        For q = 1 To IconSign.count
            If IconSign.Item(q) = IconIDNow Then
                TipeHeuristic = IconName.Item(q) & ".Variant"
                CekIconBinary = True
                Exit Function
            End If
        Next q
        CekIconBinary = False
    
End Function

Private Function CalcBinary(ByVal lpFileName As String, ByVal lpByteCount As Long, Optional ByVal StartByte As Long = 0) As String

Dim Bin() As Byte
Dim ByteSum As Long
Dim I As Long

ReDim Bin(lpByteCount) As Byte

Open lpFileName For Binary As #1
    If StartByte = 0 Then
        Get #1, , Bin
    Else
        Get #1, StartByte, Bin
    End If
Close #1
For I = 0 To lpByteCount
    ByteSum = ByteSum + Bin(I) ^ 2
Next I

CalcBinary = Hex$(ByteSum)

End Function
Private Function GetIconCount(ByVal lpFileName As String) As Long
    Dim Icon As Long
    Dim hIcon As Long
    
    hIcon = ExtractIcon(App.hInstance, lpFileName, Icon)
    Do Until hIcon <= 0
        Icon = Icon + 1
        hIcon = ExtractIcon(App.hInstance, lpFileName, Icon)
    Loop
    GetIconCount = Icon
    
End Function

Private Function CalcIcon(ByVal lpFileName As String) As String
Dim PicPath As String
Dim ByteSum As String
Dim IconExist As Long
Dim hIcon As Long

IconExist = ExtractIconEx(lpFileName, 0, ByVal 0&, hIcon, 1)
If IconExist <= 0 Then
    IconExist = ExtractIconEx(lpFileName, 0, hIcon, ByVal 0&, 1)
    If IconExist <= 0 Then Exit Function
End If

frmScanVirus.sIcon.BackColor = vbWhite
DrawIconEx frmScanVirus.sIcon.hdc, 0, 0, hIcon, 0, 0, 0, 0, DI_NORMAL
DestroyIcon hIcon

PicPath = "C:\" & GetFileName(lpFileName) & ".tmp"
SavePicture frmScanVirus.sIcon.Image, PicPath

ByteSum = CalcBinary(PicPath, FileLen(PicPath))
DeleteFile PicPath

CalcIcon = ByteSum
End Function

