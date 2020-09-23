VERSION 5.00
Begin VB.Form frmSignature 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "- Virus Signature"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "frmSignature.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Suspected Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   150
      TabIndex        =   7
      Top             =   825
      Width           =   3540
      Begin VB.Label lblFileName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1275
         TabIndex        =   16
         Top             =   390
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "File Size      :"
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
         Left            =   300
         TabIndex        =   12
         Top             =   705
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CRC-STR    :"
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
         Left            =   300
         TabIndex        =   11
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "File name    :"
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
         Left            =   300
         TabIndex        =   10
         Top             =   390
         Width           =   915
      End
      Begin VB.Label lblFileSize 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1275
         TabIndex        =   9
         Top             =   705
         Width           =   2040
      End
      Begin VB.Label lblCRCSTR 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1275
         TabIndex        =   8
         Top             =   1020
         Width           =   2040
      End
   End
   Begin VB.TextBox txtVirusName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3900
      TabIndex        =   5
      Top             =   1215
      Width           =   1815
   End
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSignature.frx":058A
      Left            =   3900
      List            =   "frmSignature.frx":05A3
      TabIndex        =   4
      Top             =   1890
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2925
      Width           =   5565
   End
   Begin SuperProtector.ShapeButton cmdOpen 
      Height          =   375
      Left            =   4455
      TabIndex        =   0
      Top             =   2400
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "Open"
      FocusRect       =   0   'False
      Picture         =   "frmSignature.frx":05F0
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
   Begin SuperProtector.ShapeButton cmdClear 
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   3375
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "Clear"
      FocusRect       =   0   'False
      Picture         =   "frmSignature.frx":0B8A
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
   Begin SuperProtector.ShapeButton cmdAddDB 
      Height          =   375
      Left            =   1665
      TabIndex        =   2
      Top             =   3375
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "Update Sign"
      FocusRect       =   0   'False
      Picture         =   "frmSignature.frx":1124
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
   Begin SuperProtector.ShapeButton cmdClose 
      Height          =   375
      Left            =   3855
      TabIndex        =   6
      Top             =   3375
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "Back to main menu"
      FocusRect       =   0   'False
      Picture         =   "frmSignature.frx":16BE
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
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning : CRC-STR (AntiPolymorphic) maybe cause false alarm. Choose only file above 1 KB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000033FF&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   3555
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Virus Type"
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
      Left            =   3900
      TabIndex        =   14
      Top             =   1575
      Width           =   840
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Virus Name"
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
      Left            =   3900
      TabIndex        =   13
      Top             =   900
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Picture         =   "frmSignature.frx":1C58
      Top             =   0
      Width           =   7620
   End
End
Attribute VB_Name = "frmSignature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddDB_Click()
    Dim I As Long
    
    If txtVirusName.Text = "" Then
        Call MsgBox("Please add virus name !", vbExclamation + vbOKOnly, "Generate Virus Signature")
        Exit Sub
    End If
    If MsgBox("Are you sure you want to" & vbCrLf & "add virus signature to this file?", vbQuestion + vbYesNo, "Add Virus Signature") = vbYes Then
        Select Case cboType.ListIndex
            Case 0
                VirType = "WORM."
            Case 1
                VirType = "TH."
            Case 2
                VirType = "WGEN."
            Case 3
                VirType = "VGEN."
            Case 4
                VirType = "W32."
            Case 5
                VirType = "BSC."
        End Select
        SignTemp = VirType & UCase((txtVirusName.Text)) + ";" + (lblCRCSTR.Caption)
          
        ExternalDatabase
        LoadExternalDatabase (False)
        HitDatabase
        ClearStatus
    End If
End Sub

Private Sub ClearStatus()
    txtPath.Text = ""
    txtVirusName.Text = ""
    cboType.Text = ""
    lblFileName.Caption = " "
    lblFileSize.Caption = " "
    lblCRCSTR.Caption = " "
    cmdAddDB.Enabled = False
    txtVirusName.Enabled = False
    cboType.Enabled = False
End Sub

Private Sub cmdClear_Click()
    ClearStatus
End Sub

Private Sub cmdClose_Click()
    Unload Me
    frmScanVirus.Enabled = True
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo ErrHandle
    Dim sFilename As String
    
    sFilename = ShowOpen(Me.hWnd, "Generate Virus Signature", "Suspected File|*.exe;*.com;*.vbs;*.bat;*.cmd;*.ocx;*.dll;*.scr;*.inf;*.*")
    SetFileAttributes sFilename, FILE_ATTRIBUTE_NORMAL
    If sFilename <> "" And (Int((FileLen(sFilename) / 1024) * 100 + 0.5) / 100) >= 1 Then
        txtPath.Text = sFilename
        lblFileName.Caption = Mid$(txtPath.Text, InStrRev(txtPath.Text, Chr$(92)) + 1)
        lblFileSize.Caption = Int((FileLen(txtPath.Text) / 1024) * 100 + 0.5) / 100 & " KB" 'Format(FileLen(Text7.Text) / 1024, "###,####") & " KB"
        lblCRCSTR.Caption = GetChecksum(txtPath.Text)
        txtVirusName.Text = ""
        cmdAddDB.Enabled = True
        txtVirusName.Enabled = True
        cboType.Enabled = True
    Else
        MsgBox "Generate Virus Signature Invalid!" & vbCrLf & "Minimum accepted file size is 1 KB", vbOKOnly + vbExclamation, APP_PROGRAM
    End If
    
ErrHandle:

End Sub
