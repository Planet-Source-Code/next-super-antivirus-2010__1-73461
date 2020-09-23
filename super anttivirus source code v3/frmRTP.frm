VERSION 5.00
Begin VB.Form frmRTP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scan Message"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SuperProtector.ShapeButton cmdQua 
      Height          =   375
      Left            =   150
      TabIndex        =   11
      Top             =   2325
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Quarantine"
      FocusRect       =   0   'False
      Picture         =   "frmRTP.frx":058A
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Virus Scan Message "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4965
      Begin VB.TextBox txtCRC 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "-"
         Top             =   1635
         Width           =   3090
      End
      Begin VB.TextBox txtDetected 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "-"
         Top             =   1320
         Width           =   3090
      End
      Begin VB.TextBox txtDir 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "-"
         Top             =   1005
         Width           =   3090
      End
      Begin VB.TextBox txtDateTime 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "-"
         Top             =   690
         Width           =   3090
      End
      Begin VB.TextBox txtMessage 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "-"
         Top             =   375
         Width           =   3090
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CRC32"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   1635
         Width           =   1440
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Detected As"
         Height          =   240
         Left            =   225
         TabIndex        =   4
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Directory"
         Height          =   240
         Left            =   225
         TabIndex        =   3
         Top             =   1005
         Width           =   1440
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Date And Time"
         Height          =   240
         Left            =   225
         TabIndex        =   2
         Top             =   690
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         Height          =   240
         Left            =   225
         TabIndex        =   1
         Top             =   375
         Width           =   1440
      End
   End
   Begin SuperProtector.ShapeButton cmdCure 
      Height          =   375
      Left            =   1815
      TabIndex        =   12
      Top             =   2325
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Cure"
      FocusRect       =   0   'False
      Picture         =   "frmRTP.frx":0B24
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
   Begin SuperProtector.ShapeButton cmdSkip 
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   2325
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Skip"
      FocusRect       =   0   'False
      Picture         =   "frmRTP.frx":10BE
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
End
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
