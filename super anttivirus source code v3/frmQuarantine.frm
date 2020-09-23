VERSION 5.00
Begin VB.Form frmQuarantine 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3105
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SuperProtector.ShapeButton cmdClean 
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   7
      CaptionAlignment=   3
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   11907757
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Clean Quarantine"
      Picture         =   "frmQuarantine.frx":0000
      Enabled         =   0   'False
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
   Begin VB.FileListBox flbQua 
      Height          =   1650
      Left            =   150
      Pattern         =   "*.Vir"
      TabIndex        =   1
      Top             =   1200
      Width           =   3990
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1965
      Left            =   0
      ScaleHeight     =   1965
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   1200
      Width           =   4665
   End
   Begin SuperProtector.ShapeButton cmdRefresh 
      Height          =   375
      Left            =   2325
      TabIndex        =   3
      Top             =   150
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Refresh"
      Picture         =   "frmQuarantine.frx":059A
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
   Begin SuperProtector.ShapeButton cmdRestore 
      Height          =   375
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   11907757
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Restore"
      Picture         =   "frmQuarantine.frx":0C6C
      Enabled         =   0   'False
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
   Begin SuperProtector.ShapeButton cmdBack 
      Height          =   375
      Left            =   2325
      TabIndex        =   5
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ButtonStyle     =   7
      CaptionAlignment=   3
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Back to Main Menu"
      Picture         =   "frmQuarantine.frx":1206
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
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmQuarantine.frx":17A0
      Top             =   0
      Width           =   10725
   End
   Begin VB.Menu mnuQua 
      Caption         =   "Quarantine"
      Visible         =   0   'False
      Begin VB.Menu mnuClean 
         Caption         =   "Clean Selected"
         Index           =   1
      End
      Begin VB.Menu mnuClean 
         Caption         =   "Clean All Object"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmQuarantine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Seal As New clsHuffman

Private Sub cmdBack_Click()
    Unload Me
    frmScanVirus.Enabled = True
End Sub

Private Sub cmdClean_Click()
    PopupMenu Me.mnuQua
End Sub

Private Sub cmdRefresh_Click()
    cmdRestore.Enabled = False
    cmdClean.Enabled = False
    flbQua.Refresh
End Sub

Private Sub cmdRestore_Click()
    Dim Alamatku As String
    
    If flbQua.FileName = "" Then
        MsgBox "File not found or file not selected.", vbExclamation, "/Quarantine"
    Else
        If MsgBox("Are you sure restore this file?", vbQuestion + vbYesNo, "/Warning") = vbYes Then
            Alamatku = FileParsePath(App.path & "\Quarantine\" & "\" & flbQua.List(flbQua.ListIndex), False, False) & FileParsePath(App.path & "\Quarantine\" & "\" & flbQua.List(flbQua.ListIndex), True, False)
            If Seal.DecodeFile(App.path & "\Quarantine\" & "\" & flbQua.List(flbQua.ListIndex), Alamatku) = False Then
                Call MsgBox("Virus Seal Invalid !", vbOKOnly, "AL127 Worm Cleaner")
                Exit Sub
            End If
            LogFile "Restore from quarantine folder  " & flbQua.FileName
            DeleteIt (App.path & "\Quarantine\" & "\" & flbQua.List(flbQua.ListIndex))
            flbQua.Refresh
        End If
    End If
End Sub

Private Sub flbQua_Click()
    If flbQua.Selected(flbQua.ListIndex) Then
        cmdRestore.Enabled = True
        cmdClean.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    Me.Caption = "- Quarantine"
    flbQua = App.path & "\Quarantine\"
End Sub

Private Sub mnuClean_Click(Index As Integer)
    Select Case Index
        Case 1: CleanSelected
        Case 2: CleanAll
    End Select
End Sub

Private Sub CleanAll()
    If flbQua.FileName = "" Then
        MsgBox "File not found or file not selected.", vbExclamation, "/Quarantine"
        Exit Sub
    ElseIf flbQua.FileName <> "" Then
        If MsgBox("Are you sure clean all object?", vbQuestion + vbYesNo, "/Warning") = vbYes Then
            Kill App.path & "\Quarantine\" & "*.vir"
            MsgBox "All object has been cleaned.", vbInformation, "/Quarantine"
            flbQua.Refresh
            cmdClean.Enabled = False
        End If
    End If
End Sub

Private Sub CleanSelected()
    If flbQua.FileName = "" Then
        MsgBox "File not found or file not selected.", vbExclamation, "/Quarantine"
    Else
        LogFile "Clean from quarantine folder   " & flbQua.FileName
        DeleteIt (App.path & "\Quarantine\" & "\" & flbQua.List(flbQua.ListIndex))
        flbQua.Refresh
    End If
End Sub
