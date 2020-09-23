VERSION 5.00
Begin VB.Form frmSetAttrib 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "- Set Attributes"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SuperProtector.ShapeButton cmdControlAttrib 
      Height          =   390
      Index           =   0
      Left            =   225
      TabIndex        =   9
      Top             =   2550
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   688
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Browse"
      Picture         =   "frmSetAttrib.frx":0000
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
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "ARCHIVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   4
      Top             =   1260
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "READ-ONLY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "HIDDEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1575
      TabIndex        =   2
      Top             =   1260
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1575
      TabIndex        =   1
      Top             =   1500
      Width           =   1275
   End
   Begin SuperProtector.ShapeButton cmdControlAttrib 
      Height          =   390
      Index           =   1
      Left            =   1740
      TabIndex        =   10
      Top             =   2550
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   688
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   11907757
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Start"
      Picture         =   "frmSetAttrib.frx":059A
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
   Begin SuperProtector.ShapeButton cmdControlAttrib 
      Height          =   390
      Index           =   2
      Left            =   3255
      TabIndex        =   11
      Top             =   2550
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   688
      ButtonStyle     =   7
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   11907757
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Stop"
      Picture         =   "frmSetAttrib.frx":0B34
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
   Begin SuperProtector.ShapeButton cmdControlAttrib 
      Height          =   390
      Index           =   3
      Left            =   4770
      TabIndex        =   12
      Top             =   2550
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   688
      ButtonStyle     =   7
      CaptionAlignment=   3
      PictureAlignment=   1
      BackColor       =   14211288
      BackColorPressed=   15715986
      BackColorHover  =   16243621
      BorderColor     =   9408398
      BorderColorPressed=   6045981
      BorderColorHover=   11632444
      Caption         =   "Back To Main"
      HandPointer     =   -1  'True
      Picture         =   "frmSetAttrib.frx":10CE
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00565656&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00565656&
      Height          =   15
      Index           =   0
      Left            =   225
      Top             =   2400
      Width           =   5835
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SET ATTRIBUTE FOR FILE / FOLDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   975
      Width           =   2865
   End
   Begin VB.Label lblFileName 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "||----"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   5835
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmSetAttrib.frx":1668
      Top             =   0
      Width           =   10725
   End
   Begin VB.Label lblvalue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ": 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   5475
      TabIndex        =   7
      Top             =   1260
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00565656&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00565656&
      Height          =   15
      Index           =   1
      Left            =   225
      Top             =   1875
      Width           =   5835
   End
   Begin VB.Label lblDir 
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
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   240
      TabIndex        =   6
      Top             =   1950
      Width           =   5895
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FILE(S)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4500
      TabIndex        =   5
      Top             =   1260
      Width           =   915
   End
End
Attribute VB_Name = "frmSetAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Av Super Protector - A N T I V I R U S


Option Explicit

Dim var_file As Long
Dim var_dir As Long
Dim var_filed As Long

Private Sub cmdControlAttrib_Click(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0
            Where = BrowseForFolder(Me.hWnd, "Select Driver or Folder to set Attribute")
            If Len(Where) > 0 Then
                lblDir = Where
                cmdControlAttrib(1).Enabled = True
            End If
        Case 1:
            cmdControlAttrib(0).Enabled = False
            cmdControlAttrib(1).Enabled = False
            cmdControlAttrib(2).Enabled = True
            cmdControlAttrib(3).Enabled = False
            StopScan = False
            Check1(0).Enabled = False
            Check1(1).Enabled = False
            Check1(2).Enabled = False
            Check1(3).Enabled = False
            LoopingAttrib
            StopScan = True
            cmdControlAttrib(0).Enabled = True
            cmdControlAttrib(1).Enabled = True
            cmdControlAttrib(2).Enabled = False
            cmdControlAttrib(3).Enabled = True
            Check1(0).Enabled = True
            Check1(1).Enabled = True
            Check1(2).Enabled = True
            Check1(3).Enabled = True
        Case 2
            StopScan = True
        Case 3
            Unload Me
            frmScanVirus.Show
            
            frmScanVirus.Enabled = True
    End Select
End Sub

Private Function FindFilesAttrib(path As String, SearchStr As String, FileCount As Long, DirCount As Long)
    On Error Resume Next
    
    Dim FileName As String
    Dim FindFiles As Long
    Dim NAMA_DIRECTORY As String
    Dim DIR_NAMES() As String
    Dim nDir As Long
    Dim I As Long
    Dim buff As Long
    Dim hRes As FILE_ATTRIBUTE
    
    If Check1(0).value Then hRes = hRes + FILE_ATTRIBUTE_ARCHIVE
    If Check1(1).value Then hRes = hRes + FILE_ATTRIBUTE_READONLY
    If Check1(2).value Then hRes = hRes + FILE_ATTRIBUTE_HIDDEN
    If Check1(3).value Then hRes = hRes + FILE_ATTRIBUTE_SYSTEM
    
    If Right(path, 1) <> "\" Then path = path & "\"
        nDir = 0
    ReDim DIR_NAMES(nDir)
        NAMA_DIRECTORY = Dir(path, vbDirectory Or vbHidden Or vbSystem Or vbReadOnly)
    
    Do While Len(NAMA_DIRECTORY) > 0
        If (NAMA_DIRECTORY <> ".") And (NAMA_DIRECTORY <> "..") Then
            If GetAttr(path & NAMA_DIRECTORY) Or FILE_ATTRIBUTE_DIRECTORY Then
                DIR_NAMES(nDir) = NAMA_DIRECTORY
                DirCount = DirCount + 1
                nDir = nDir + 1
        
                ReDim Preserve DIR_NAMES(nDir)
            End If
        End If
        NAMA_DIRECTORY = Dir()
    Loop
    If StopScan = True Then Exit Function
        While PauseScan = True
            DoEvents
        Wend
        FileName = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbArchive)
    
    While Len(FileName) <> 0
        If StopScan = True Then Exit Function
            While PauseScan = True
                DoEvents
            Wend
            
            FindFiles = FindFiles + FileLen(path & FileName)
            FileCount = FileCount + 1
            JumlahBuffer = FileCount
                
            lblFileName.Caption = "||---- " & path & FileName
            JumlahFile = FileCount
            lblvalue(1).Caption = ": " & JumlahFile
            
        DoEvents
        FileName = Dir()
    Wend
    
    If nDir > 0 Then
        For I = 0 To nDir - 1
            FindFilesAttrib = FindFilesAttrib + FindFilesAttrib(path & DIR_NAMES(I) & "\", SearchStr, FileCount, DirCount)
            SetFileAttributes path & DIR_NAMES(I), hRes
        Next I
        DoEvents
    End If
End Function

Public Function LoopingAttrib()
    Dim SearchPath As String, FindStr As String
    Dim FileSize As Long
    Dim NumFiles As Long, NumDirs As Long
    
    SearchPath = Where
    FindStr = ext
    FileSize = FindFilesAttrib(SearchPath, FindStr, NumFiles, NumDirs)
    DoEvents
End Function

Private Sub Form_Load()
    Me.Caption = "- Set Attributes"
End Sub
