Attribute VB_Name = "mdlSystemTray"
    '******************************************************************************
    'Systray Module
    '
    'Mark Mokoski
    'markm@cmtelephone.com
    'www.cmtelephone.com
    '
    '6-NOV-2004
    '
    'Put App in SysTray, remove App from SysTray, Form on top, Balloon ToolTip code
    '
    'See Systray Form Code.txt in the ZIP file for form add-in's to make it all work
    '
    'Also see Microsoft Knowledge base http://support.microsoft.com/default.aspx?scid=kb;en-us;149276
    'for more information.
    '
    'This code is based on the Microsoft Knowledge Base code.
    '******************************************************************************

Public Sub SystrayOn(frm As Form, IconTooltipText As String)

    'Adds Icon to SysTray

        With vbTray
            .cbSize = Len(vbTray)
            .hWnd = frm.hWnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .szTip = Trim(IconTooltipText$) & vbNullChar
            .hIcon = frm.Icon
        End With
    
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    
End Sub

Public Sub SystrayOff(frm As Form)

    'Removes Icon from SysTray

        With vbTray
            .cbSize = Len(vbTray)
            .hWnd = frm.hWnd
            .uID = vbNull
        End With
    
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
    
End Sub

Public Sub ChangeSystrayToolTip(frm As Form, IconTooltipText As String)

    'Changes the SysTray Balloon Tool Tip Text

        With vbTray
            .cbSize = Len(vbTray)
            .hWnd = frm.hWnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .szTip = Trim(IconTooltipText$) & vbNullChar
            .hIcon = frm.Icon
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)
    
End Sub

Public Sub FormOnTop(frm As Form)

    'Puts your form ontop of all the other windows!
    Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)

End Sub

Public Sub PopupBalloon(frm As Form, Message As String, Title As String, Optional balType As TypeBallon = NIIF_INFO)

    'Set a Balloon tip on Systray

    'Call RemoveBalloon(frm), This removes any current Balloon Tip that is active.
    'If you want Balloon Tips to "Stack up" and display in sequence
    'after each times out (or you click on the Balloon Tip to clear it),
    'comment out the Call below.

    Call RemoveBalloon(frm)

        With vbTray
            .cbSize = Len(vbTray)
            .hWnd = frm.hWnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = frm.Icon
            .dwState = 0
            .dwStateMask = 0
            .szInfo = Message & Chr(0)
            .szInfoTitle = Title & Chr(0)
            'Choose the message icon below, NIIF_NONE, NIIF_WARNING, NIIF_ERROR, NIIF_INFO
            .dwInfoFlags = balType
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub

Public Sub RemoveBalloon(frm As Form)

    'Kill any current Balloon tip on screen for referenced form
  
        With vbTray
            .cbSize = Len(vbTray)
            .hWnd = frm.hWnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = frm.Icon
            .dwState = 0
            .dwStateMask = 0
            .szInfo = Chr(0)
            .szInfoTitle = Chr(0)
            .dwInfoFlags = NIIF_NONE
        End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub




