VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "- About"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SuperProtector.ShapeButton cmdClose 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   6480
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   661
      ButtonStyle     =   7
      ButtonStyleColors=   3
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Back to Main Menu"
      Picture         =   "frmAbout.frx":11185
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   4815
      Width           =   1860
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ": -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   2400
      TabIndex        =   13
      Top             =   4815
      Width           =   2985
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Engine"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   5055
      Width           =   1860
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Virus Signature"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   5295
      Width           =   1860
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ": -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   5055
      Width           =   2625
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ": -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   2400
      TabIndex        =   9
      Top             =   5295
      Width           =   2385
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ": -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   2400
      TabIndex        =   8
      Top             =   5535
      Width           =   2505
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ": -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   2400
      TabIndex        =   7
      Top             =   5775
      Width           =   2625
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Processes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   5535
      Width           =   2100
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tweak Registry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   5775
      Width           =   1980
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Build"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   4575
      Width           =   1860
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ": -"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   2400
      TabIndex        =   3
      Top             =   4575
      Width           =   2745
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Informations :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This software is provided 'as-is' free of charge without any warranty of any kind."
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
      Left            =   0
      TabIndex        =   1
      Top             =   6120
      Width           =   5925
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long

Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_MEMORY = &H4

Private Sub cmdClose_Click()
    Unload Me
    Set frmAbout = Nothing
    frmScanVirus.Enabled = True
    frmScanVirus.Show
End Sub

Private Sub Form_Load()
    With Me
        .lblInfo(0) = ": " & APP_VERSION
        .lblInfo(1) = ": " & ENGINE_VERSION
        .lblInfo(2) = ": " & frmScanVirus.lstVirus.ListItems.count & " Viruses"
        .lblInfo(3) = ": " & PROCESSESES
        .lblInfo(4) = ": " & TWEAK_REG
        .lblInfo(5) = ": " & CURRENT_BUILD
    End With
'    ControlTabs Me, tbsAbout, tbsAbout.SelectedItem.Index
'    txtVersion.Text = LoadResString(102)
End Sub

Private Sub picAbout_Click()

End Sub

Private Sub tbsAbout_Click()

    picAbout.Cls
    picVersion.Cls
    
    ControlTabs Me, tbsAbout, tbsAbout.SelectedItem.Index
    If tbsAbout.SelectedItem.Caption = "About" Then
        picAbout.Visible = True
        picVersion.Visible = False
    Else
        tbsAbout.SelectedItem.Caption = "Version History"
        picAbout.Visible = False
        picVersion.Visible = True
    End If
End Sub
