VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Left            =   4275
      Top             =   900
   End
   Begin SuperProtector.ProgressBar ProgWait 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
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
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If tmrWait.Enabled = False Then
        tmrWait.Enabled = True
        tmrWait.Interval = 1
    End If
End Sub

Private Sub tmrWait_Timer()
    If ProgWait.value < 25 Then
        ProgWait.value = ProgWait.value + 1
    ElseIf ProgWait.value = 25 Then
        ProgWait.value = 55
    ElseIf ProgWait.value >= 55 And ProgWait.value < 65 Then
        ProgWait.value = ProgWait.value + 2
    ElseIf ProgWait.value = 65 Then
        ProgWait.value = 75
    ElseIf ProgWait.value = 75 Then
        ProgWait.value = 85
    ElseIf ProgWait.value = 85 Then
        ProgWait.value = 100
    ElseIf ProgWait.value = 100 Then
        tmrWait.Enabled = False
        Unload Me
    End If
End Sub
