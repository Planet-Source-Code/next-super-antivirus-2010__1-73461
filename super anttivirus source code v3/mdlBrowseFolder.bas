Attribute VB_Name = "mdlBrowseFolder"
Option Explicit

Public Function BrowseForFolder(ByVal lnghWnd As Long, _
    ByVal strPrompt As String) As String
    On Error GoTo ehBrowseForFolder
    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo
    With udtBI
        .lnghWnd = lnghWnd
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_NEWDIALOGSTYLE + BIF_EDITBOX
    End With
    lngIDList = SHBrowseForFolder(udtBI)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, 0)
        lngResult = SHGetPathFromIDList(lngIDList, _
            strPath)
        Call CoTaskMemFree(lngIDList)
        intNull = InStr(strPath, vbNullChar)
            If intNull > 0 Then
                strPath = Left(strPath, intNull - 1)
            End If
    End If
    BrowseForFolder = strPath
    Exit Function
ehBrowseForFolder:
    BrowseForFolder = Empty
End Function

Public Function ShowOpen(hwnd As Long, Optional Title As String = "Open", Optional extFile As String = "All files|*.*") As String
    Dim OFName As OPENFILENAME
    extFile = Replace(extFile, "|", Chr(0))
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = extFile
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = App.path
    OFName.lpstrTitle = Title
    OFName.FLAGS = 0
    If GetOpenFileNameEx(OFName) Then
       ShowOpen = Left(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
    Else
       ShowOpen = ""
    End If
End Function


