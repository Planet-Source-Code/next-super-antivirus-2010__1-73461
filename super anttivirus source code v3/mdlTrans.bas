Attribute VB_Name = "mdlTrans"
Option Explicit

Public Enum TransType
  LWA_OPAQUE = 0
  LWA_COLORKEY = 1
  LWA_ALPHA = 2
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Private Const zFormOrPictBoxStr = "Must pass in the name of either a Form or a PictureBox."

Public Function isTransparent(zForm As Form) As TransType
  On Local Error Resume Next
  Dim vTrans As Byte, ALPHA As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.hWnd, cKey, vTrans, ALPHA
  If Err Then
    isTransparent = -1
  Else
    isTransparent = ALPHA
  End If
End Function

Public Function GetTrans(zForm As Form) As Long
  On Local Error Resume Next
  Dim vTrans As Byte, ALPHA As TransType, cKey As Long
  GetLayeredWindowAttributes zForm.hWnd, cKey, vTrans, ALPHA
  If ALPHA = LWA_ALPHA Then
    GetTrans = vTrans
  ElseIf ALPHA = LWA_COLORKEY Then
    GetTrans = cKey
  Else
    GetTrans = -1
  End If
  If Err Then
    GetTrans = -1
  End If
End Function

Public Function FadeIn(zForm As Form, Optional ByVal Final As Byte = 255, Optional ByVal vStep As Single = 2) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean, VarTmp As Single
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 0
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 0
    vTrans = 0
  End If
  If vTrans > Final Then
    FadeIn = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  VarTmp = vTrans
  While VarTmp < Final
    DoEvents
    VarTmp = VarTmp + vStep
    If VarTmp > Final Then VarTmp = Final
    SetTrans zForm, CByte(VarTmp)
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Err Then
    FadeIn = False
  Else
    FadeIn = True
  End If
End Function

Public Function FadeOut(zForm As Form, Optional ByVal Final As Byte = 0, Optional ByVal vStep As Single = 2) As Boolean
  On Local Error Resume Next
  Dim vTrans As Long, ZFE As Boolean, VarTmp As Single
  vTrans = isTransparent(zForm)
  If vTrans <> LWA_ALPHA Then SetTrans zForm, 255
  vTrans = GetTrans(zForm)
  If vTrans = -1 Then
    SetTrans zForm, 255
    vTrans = 255
  End If
  If vTrans < Final Then
    FadeOut = False
    Exit Function
  End If
  If zForm.Visible = False Then zForm.Show
  ZFE = zForm.Enabled
  If ZFE = True Then zForm.Enabled = False
  VarTmp = vTrans
  While VarTmp > Final
    DoEvents
    VarTmp = VarTmp - vStep
    If VarTmp < Final Then VarTmp = Final
    SetTrans zForm, CByte(VarTmp)
  Wend
  If ZFE = True Then zForm.Enabled = True
  If Final = 0 Then zForm.Hide
  If Err Then
    FadeOut = False
  Else
    FadeOut = True
  End If
End Function

Public Function SetTrans(zForm As Form, Optional ByVal vTrans As Byte = 127) As Boolean
  On Local Error Resume Next
  Dim msg As Long
  msg = GetWindowLong(zForm.hWnd, GWL_EXSTYLE)
  msg = msg Or WS_EX_LAYERED
  SetWindowLong zForm.hWnd, GWL_EXSTYLE, msg
  SetLayeredWindowAttributes zForm.hWnd, 0, vTrans, LWA_ALPHA
  If Err Then
    SetTrans = False
  Else
    SetTrans = True
  End If
End Function

Public Function MakeTransparent(hWnd As Long, Perc As Integer) As Long
    Dim msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
        MakeTransparent = 1
    Else
        msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
        SetWindowLong hWnd, GWL_EXSTYLE, msg
        SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
        MakeTransparent = 0
    End If
    If Err Then
        MakeTransparent = 2
    End If
End Function




