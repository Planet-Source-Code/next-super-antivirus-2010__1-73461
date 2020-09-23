VERSION 5.00
Begin VB.Form frmGPL 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "- Tips For Virus Detection"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGPL 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7725
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmGPL.frx":0000
      Top             =   0
      Width           =   13125
   End
   Begin SuperProtector.ShapeButton cmdClose 
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   7920
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   661
      ButtonStyle     =   7
      ButtonStyleColors=   3
      CaptionAlignment=   3
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Back to Main Menu"
      Picture         =   "frmGPL.frx":04F9
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
Attribute VB_Name = "frmGPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
    frmScanVirus.Show
    
End Sub

Private Sub Form_Load()
'    txtGPL.Text = LoadResString(101)
End Sub

