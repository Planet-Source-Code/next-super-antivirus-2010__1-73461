VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1635
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1635
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFadeout 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5475
      Top             =   75
   End
   Begin VB.CheckBox chkUnclose 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   1365
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5025
      Top             =   75
   End
   Begin SuperProtector.ProgressBar progLoad 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      Value           =   0
      Theme           =   8
      TextStyle       =   4
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "U11D ProgressBar"
      TextEffectColor =   16777215
      TextEffect      =   5
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblProcess 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label lblPercen 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 % Completed."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2100
      TabIndex        =   4
      Top             =   420
      Width           =   1365
   End
   Begin VB.Label lblLoad 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait... Configuring Environment."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   735
      Width           =   5790
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "V3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4005
      TabIndex        =   2
      Top             =   1320
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2010  Ermal Gjermeni Softwares"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   1050
      Width           =   5940
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lAlpha As Integer

Private Sub chkUnclose_Click()
    If chkUnclose.value = 1 Then Exit Sub
End Sub

Private Sub Form_Activate()

    lblVersion = APP_VERSION
    App.Title = "Av Super Protector"
    App.TaskVisible = False
    Where = GetSpecialfolder(CSIDL_STARTUP Or CSIDL_APPDATA)
    ext = "*.*"
    Buffering = False
  '  ProcedureScan

    Looping
'    FadeIn Me
    lAlpha = 255
    Tunggu 0.5
    If tmrLoad.Enabled = False Then
        tmrLoad.Enabled = True
    Else
        tmrLoad.Enabled = False
    End If
End Sub


Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then
        MsgBox "already runing in your system.", vbExclamation, "Av Super Protector"
        End
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If chkUnclose.value = 0 Then Exit Sub
    Cancel = 1
End Sub



Private Sub tmrFadeout_Timer()
    If lAlpha > 0 Then
        DoEvents
        lAlpha = lAlpha - 5
        MakeTransparent Me.hWnd, lAlpha
    Else
        lAlpha = 0
        frmScanVirus.Show
        tmrFadeout.Enabled = False
        Unload Me
    End If
End Sub

Private Sub tmrLoad_Timer()
    With progLoad
        If .value < 100 Then
            DoEvents
            .value = .value + 1
            If .value = 30 Then
                lblPercen.Caption = progLoad.value & " % Completed."
                lblLoad.Caption = "Please wait Av Super Protector."
                Tunggu 2
            End If
            If .value = 60 Then
                lblPercen.Caption = progLoad.value & " % Completed."
                lblLoad.Caption = "Configuring Database"
                If Dir$(App.path & "\ansavcore.dll", vbArchive Or vbNormal Or vbHidden Or vbReadOnly Or vbSystem) = "" Then
                    If MsgBox(APP_PROGRAM & vbCrLf & "Error Detected :" & vbCrLf & vbCrLf & " - ansavcore.dll not found", vbCritical + vbOKOnly, "Critical Error") = vbOK Then
                        ANSAVEnable = False
                    End If
                Else
                    Shell "regsvr32 /s" & "ansavcore.dll", vbHide
                    ANSAVEnable = True
                End If
                Tunggu 2
            End If
            If .value = 90 Then
                lblPercen.Caption = progLoad.value & " % Completed."
                lblLoad.Caption = "Scanning Processes And Startup"
                ScanProcess False
                Tunggu 2
                lblProcess.Caption = ""
            End If
            If .value = 100 Then
                lblPercen.Caption = progLoad.value & " % Completed."
                lblLoad.Caption = "Building Main Application."
                Tunggu 1
                tmrLoad.Enabled = False
                tmrFadeout.Enabled = True
                Tunggu 2.5
                With frmScanVirus
                    .Refresh
                    .Enabled = True
                    .Show
                End With
            End If
        End If
    End With
    With progLoad
       If .value < 30 Then
          .value = .value + 3
          lblPercen.Caption = .value & " % Completed."
          lblLoad.Caption = "Please wait Av Super Protector is configuring environment."
           Tunggu 0.1
        ElseIf .value = 30 Then
            lblPercen.Caption = .value & " % Completed."
            .value = 50
        ElseIf .value >= 50 And .value < 60 Then
            .value = .value + 2
            lblPercen.Caption = .value & " % Completed."
            lblLoad.Caption = "Configuring Database"
            HitDatabase
            Tunggu 0.1
       ElseIf .value = 60 Then
           Tunggu 1
            .value = 80
            lblPercen.Caption = .value & " % Completed."
            lblLoad.Caption = "Scanning Processes And Startup"
            ScanProcess False
            Tunggu 1
            lblProcess.Caption = ""
        ElseIf .value = 80 Then
            .value = 100
            lblPercen.Caption = .value & " % Completed."
            lblLoad.Caption = "Building Main Application."
            Tunggu 2
        ElseIf .value = 100 Then
            tmrLoad.Enabled = False
            FadeOut Me
           tmrFadeout.Enabled = True
           Unload Me
            Tunggu 2
            With frmScanVirus
                .Refresh
                .Enabled = True
                .Show
            End With
        End If
    End With
End Sub

