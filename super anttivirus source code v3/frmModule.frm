VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModule 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "- Module Info"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   5340
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   1500
      Width           =   8115
      Begin VB.PictureBox picAdvanced 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5265
         Left            =   0
         ScaleHeight     =   5265
         ScaleWidth      =   7965
         TabIndex        =   6
         Top             =   0
         Width           =   7965
         Begin VB.Frame fraFileInfor 
            BackColor       =   &H00FFFFFF&
            Caption         =   "File Information :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   150
            TabIndex        =   39
            Top             =   2175
            Width           =   7665
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   10
               Left            =   150
               Picture         =   "frmModule.frx":0000
               Top             =   2700
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   9
               Left            =   150
               Picture         =   "frmModule.frx":058A
               Top             =   2460
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   8
               Left            =   150
               Picture         =   "frmModule.frx":0B14
               Top             =   2220
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   7
               Left            =   150
               Picture         =   "frmModule.frx":109E
               Top             =   1980
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   6
               Left            =   150
               Picture         =   "frmModule.frx":1628
               Top             =   1740
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   5
               Left            =   150
               Picture         =   "frmModule.frx":1BB2
               Top             =   1500
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   4
               Left            =   150
               Picture         =   "frmModule.frx":213C
               Top             =   1260
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   3
               Left            =   150
               Picture         =   "frmModule.frx":26C6
               Top             =   1020
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   2
               Left            =   150
               Picture         =   "frmModule.frx":2C50
               Top             =   780
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   1
               Left            =   150
               Picture         =   "frmModule.frx":31DA
               Top             =   540
               Width           =   240
            End
            Begin VB.Image ImgInfor 
               Height          =   240
               Index           =   0
               Left            =   150
               Picture         =   "frmModule.frx":3764
               Top             =   300
               Width           =   240
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "File Type"
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
               Height          =   240
               Index           =   0
               Left            =   450
               TabIndex        =   61
               Top             =   300
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Company"
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
               Height          =   240
               Index           =   1
               Left            =   450
               TabIndex        =   60
               Top             =   540
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Description"
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
               Height          =   240
               Index           =   2
               Left            =   450
               TabIndex        =   59
               Top             =   780
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Version"
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
               Height          =   240
               Index           =   3
               Left            =   450
               TabIndex        =   58
               Top             =   1020
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Internal Name"
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
               Height          =   240
               Index           =   4
               Left            =   450
               TabIndex        =   57
               Top             =   1260
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Copyright"
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
               Height          =   240
               Index           =   5
               Left            =   450
               TabIndex        =   56
               Top             =   1500
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Trademark"
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
               Height          =   240
               Index           =   6
               Left            =   450
               TabIndex        =   55
               Top             =   1740
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Original File Name"
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
               Height          =   240
               Index           =   7
               Left            =   450
               TabIndex        =   54
               Top             =   1980
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Product Name"
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
               Height          =   240
               Index           =   8
               Left            =   450
               TabIndex        =   53
               Top             =   2220
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Product Version"
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
               Height          =   240
               Index           =   9
               Left            =   450
               TabIndex        =   52
               Top             =   2460
               Width           =   1365
            End
            Begin VB.Label lblInfor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Comments"
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
               Height          =   240
               Index           =   10
               Left            =   450
               TabIndex        =   51
               Top             =   2700
               Width           =   1365
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   2175
               TabIndex        =   50
               Top             =   300
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   1
               Left            =   2175
               TabIndex        =   49
               Top             =   540
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   2
               Left            =   2175
               TabIndex        =   48
               Top             =   780
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   3
               Left            =   2175
               TabIndex        =   47
               Top             =   1020
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   4
               Left            =   2175
               TabIndex        =   46
               Top             =   1260
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   5
               Left            =   2175
               TabIndex        =   45
               Top             =   1500
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   6
               Left            =   2175
               TabIndex        =   44
               Top             =   1740
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   7
               Left            =   2175
               TabIndex        =   43
               Top             =   1980
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   8
               Left            =   2175
               TabIndex        =   42
               Top             =   2220
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   9
               Left            =   2175
               TabIndex        =   41
               Top             =   2460
               Width           =   5340
            End
            Begin VB.Label lblFileInfor 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   10
               Left            =   2175
               TabIndex        =   40
               Top             =   2700
               Width           =   5340
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FFFFFF&
            Height          =   1890
            Left            =   150
            TabIndex        =   37
            Top             =   150
            Width           =   7665
            Begin MSComctlLib.ListView lvwMod 
               Height          =   1500
               Left            =   150
               TabIndex        =   38
               Top             =   225
               Width           =   7365
               _ExtentX        =   12991
               _ExtentY        =   2646
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               SmallIcons      =   "ilsMod"
               ForeColor       =   0
               BackColor       =   16777215
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   13
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Data"
                  Object.Width           =   4304
                  ImageIndex      =   1
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Path"
                  Object.Width           =   7544
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Type"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Company Name"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Description"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Version"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Internal Name"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Copyright"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Trademark"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Original File Name"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   10
                  Text            =   "Product Name"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   11
                  Text            =   "Product Version"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   12
                  Text            =   "Comment"
                  Object.Width           =   2540
               EndProperty
            End
         End
         Begin VB.Label lblNotFound 
            BackStyle       =   0  'Transparent
            Caption         =   "Cannot open file. Module not found."
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
            Height          =   255
            Left            =   510
            TabIndex        =   62
            Top             =   4868
            Width           =   3015
         End
         Begin VB.Image imgEmpty 
            Height          =   240
            Left            =   150
            Picture         =   "frmModule.frx":3CEE
            Top             =   4875
            Width           =   240
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5340
      Index           =   0
      Left            =   75
      TabIndex        =   3
      Top             =   1500
      Width           =   8115
      Begin VB.PictureBox picGeneral 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5265
         Left            =   75
         ScaleHeight     =   5265
         ScaleWidth      =   7965
         TabIndex        =   4
         Top             =   0
         Width           =   7965
         Begin SuperProtector.ShapeButton cmdClose 
            Height          =   465
            Left            =   5700
            TabIndex        =   64
            Top             =   4650
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   820
            ButtonStyle     =   7
            PictureAlignment=   1
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   "Back to Main Menu"
            Picture         =   "frmModule.frx":4278
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
         Begin SuperProtector.ShapeButton cmdFileProperties 
            Height          =   465
            Left            =   5700
            TabIndex        =   65
            Top             =   4050
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   820
            ButtonStyle     =   7
            PictureAlignment=   1
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   "File Properties"
            Picture         =   "frmModule.frx":4812
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
            Height          =   2250
            Left            =   5625
            Picture         =   "frmModule.frx":4DAC
            Top             =   1200
            Width           =   2250
         End
         Begin VB.Label lblStartedOn 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   36
            Top             =   3795
            Width           =   2955
         End
         Begin VB.Label lblDateModif 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   35
            Top             =   4050
            Width           =   2955
         End
         Begin VB.Label lblPID 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   34
            Top             =   2265
            Width           =   2955
         End
         Begin VB.Label lblThreads 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   33
            Top             =   2775
            Width           =   2955
         End
         Begin VB.Label lblMemory 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   32
            Top             =   3030
            Width           =   2955
         End
         Begin VB.Label lblPriority 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   31
            Top             =   3285
            Width           =   2955
         End
         Begin VB.Label lblBase 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   30
            Top             =   2520
            Width           =   2955
         End
         Begin VB.Label lblCreated 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   29
            Top             =   1755
            Width           =   4005
         End
         Begin VB.Label lblAttributes 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   28
            Top             =   1500
            Width           =   2955
         End
         Begin VB.Label lblSize 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   27
            Top             =   1245
            Width           =   2955
         End
         Begin VB.Label lblType 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   26
            Top             =   990
            Width           =   2955
         End
         Begin VB.Label lblFile 
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
            Height          =   255
            Left            =   1815
            TabIndex        =   25
            Top             =   225
            Width           =   6045
         End
         Begin VB.Label lblLocation 
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
            Height          =   480
            Left            =   1815
            TabIndex        =   24
            Top             =   480
            Width           =   6045
         End
         Begin VB.Label lblCopyright 
            BackStyle       =   0  'Transparent
            Caption         =   "© Microsoft Corporation. All rights reserved."
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
            Height          =   555
            Left            =   240
            TabIndex        =   23
            Top             =   4650
            Width           =   4995
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Base Priority"
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
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   22
            Top             =   2520
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Priority"
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
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   21
            Top             =   3285
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Memory"
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
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   20
            Top             =   3030
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Threads"
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
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   19
            Top             =   2775
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Process ID"
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
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   18
            Top             =   2265
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
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
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "File"
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Modified"
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
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   15
            Top             =   4050
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Started On"
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
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   14
            Top             =   3795
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
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
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   990
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
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
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Top             =   1245
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Attributes"
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
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   1500
            Width           =   1080
         End
         Begin VB.Label lblEmpty 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Created"
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
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   10
            Top             =   1755
            Width           =   1080
         End
      End
   End
   Begin VB.Frame fraBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   8220
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1290
      Left            =   0
      Picture         =   "frmModule.frx":8975
      ScaleHeight     =   1290
      ScaleWidth      =   8265
      TabIndex        =   2
      Top             =   375
      Width           =   8265
      Begin VB.PictureBox picIco 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   300
         ScaleHeight     =   480
         ScaleWidth      =   450
         TabIndex        =   63
         Top             =   225
         Width           =   510
      End
      Begin MSComctlLib.ImageList ilsMod 
         Left            =   7575
         Top             =   525
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label lblModuleUsed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   4500
         TabIndex        =   67
         Top             =   150
         Width           =   3615
      End
      Begin VB.Label lblModUsage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5475
         TabIndex        =   66
         Top             =   412
         Width           =   2640
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "Microsoft Corporation"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1050
         TabIndex        =   9
         Top             =   405
         Width           =   6375
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "6.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1050
         TabIndex        =   8
         Top             =   660
         Width           =   6375
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Explorer"
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
         Height          =   255
         Left            =   1050
         TabIndex        =   7
         Top             =   150
         Width           =   6375
      End
   End
   Begin MSComctlLib.TabStrip tabFileInfo 
      Height          =   6690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   11800
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(256) As Byte
End Type

Private Const LVM_FIRST = &H1000

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function OpenFile Lib "kernel32.dll" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Dim GetIco As New clsGetIconFile

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdFileProperties_Click()

    'frmWait.Show vbModal
    On Error Resume Next
    Dim I As Integer
    
    For I = 1 To frmScanVirus.lstView.ListItems.count
      If frmScanVirus.lstView.ListItems(I).Selected Then
         ShowProps frmScanVirus.lstView.ListItems(I).SubItems(1), Me.hWnd
      End If
    Next I
    
End Sub

Private Sub Form_Activate()

    MakeInfo
    GetModuleProcessID frmScanVirus.lstView, 5, lvwMod, ilsMod
    
End Sub

Private Sub Form_Load()

    ControlTabs Me, tabFileInfo, tabFileInfo.SelectedItem.Index

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmScanVirus.Show
Me.Hide

End Sub

Private Sub lvwMod_Click()
    
    FillList
    
End Sub



Private Sub tabFileInfo_Click()

    picGeneral.Cls
    picAdvanced.Cls
    
    ControlTabs Me, tabFileInfo, tabFileInfo.SelectedItem.Index
    If tabFileInfo.SelectedItem.Caption = "Advanced" Then
        picGeneral.Visible = False
        picAdvanced.Visible = True
        CheckModule
    Else
        tabFileInfo.SelectedItem.Caption = "General"
        picGeneral.Visible = True
        picAdvanced.Visible = False
        lblModuleUsed.Visible = False
        lblModUsage.Visible = False
        MakeInfo
    End If
    
End Sub

Private Sub FillList()

'Dim i As Integer
'
'lblFileInfor(0).Caption = ": " & lvwMod.SelectedItem
'For i = 0 To 8
'    lblFileInfor(i).Caption = ": " & lvwMod.SelectedItem.SubItems(i)
'Next i

    lblFileInfor(0).Caption = ": " & lvwMod.SelectedItem.SubItems(2)
    lblFileInfor(1).Caption = ": " & lvwMod.SelectedItem.SubItems(3)
    lblFileInfor(2).Caption = ": " & lvwMod.SelectedItem.SubItems(4)
    lblFileInfor(3).Caption = ": " & lvwMod.SelectedItem.SubItems(5)
    lblFileInfor(4).Caption = ": " & lvwMod.SelectedItem.SubItems(6)
    lblFileInfor(5).Caption = ": " & lvwMod.SelectedItem.SubItems(7)
    lblFileInfor(6).Caption = ": " & lvwMod.SelectedItem.SubItems(8)
    lblFileInfor(7).Caption = ": " & lvwMod.SelectedItem.SubItems(9)
    lblFileInfor(8).Caption = ": " & lvwMod.SelectedItem.SubItems(10)
    lblFileInfor(9).Caption = ": " & lvwMod.SelectedItem.SubItems(11)
    lblFileInfor(10).Caption = ": " & lvwMod.SelectedItem.SubItems(12)
End Sub

Private Sub CheckModule()

    Dim I As Integer
    Dim strFile As String
    Dim hVer As VERHEADER
    Dim fso As New FileSystemObject
    Dim FileInfo As file

    strFile = frmScanVirus.lstView.SelectedItem.SubItems(1)
    Set FileInfo = fso.GetFile(strFile)
    GetVerHeader strFile, hVer

    If lvwMod.ListItems.count = 0 Then
        For I = lblInfor.LBound To lblInfor.UBound
            lblInfor.Item(I).Visible = False
            lblFileInfor.Item(I).Visible = False
        Next I
            Frame1(0).Visible = False
            Frame2.Visible = False
            fraFileInfor.Visible = False
            lvwMod.Visible = False
    Else
        For I = lblInfor.LBound To lblInfor.UBound
            lblInfor.Item(I).Visible = True
            lblFileInfor.Item(I).Visible = True
        Next I
            Frame2.Visible = True
            lblModuleUsed.Visible = True
            lblModUsage.Visible = True
            lblModuleUsed.Caption = "Modules Used By : " & " [ " & FileInfo.ShortName & " ]"
            lblModUsage.Caption = "Total Usage : " & lvwMod.ListItems.count & " Modules"
            Frame1(1).Visible = True
            fraFileInfor.Visible = True
            lblFileInfor(0).Caption = ": " & GetPathType(strFile)
            lblFileInfor(1).Caption = ": " & hVer.CompanyName
            lblFileInfor(2).Caption = ": " & hVer.FileDescription
            lblFileInfor(3).Caption = ": " & hVer.FileVersion
            lblFileInfor(4).Caption = ": " & hVer.InternalName
            lblFileInfor(5).Caption = ": " & hVer.LegalCopyright
            lblFileInfor(6).Caption = ": " & hVer.LegalTradeMarks
            lblFileInfor(7).Caption = ": " & hVer.OrigionalFileName
            lblFileInfor(8).Caption = ": " & hVer.ProductName
            lblFileInfor(9).Caption = ": " & hVer.ProductVersion
            lblFileInfor(10).Caption = ": " & hVer.Comments
        lvwMod.Visible = True
    End If
End Sub

Private Sub MakeInfo()

    On Error Resume Next

    Dim strFile As String
    Dim hPID As Long
    Dim hVer As VERHEADER
    Dim hIcoExt As Long, hIcoDraw As Long
    Dim fso As New FileSystemObject
    Dim FileInfo As file

    picIco.Cls

    strFile = frmScanVirus.lstView.SelectedItem.SubItems(1)
    Set FileInfo = fso.GetFile(strFile)

    GetVerHeader strFile, hVer

    lblDescription = hVer.FileDescription
    lblCompany = hVer.CompanyName
    lblVersion = hVer.FileVersion
    lblCopyright = hVer.LegalCopyright
    lblFile = ": " & FileInfo.ShortName ' GetFileName(strFile)
    lblLocation = ": " & FileInfo.ParentFolder ' GetFilePath(strFile)
    lblType = ": " & GetPathType(strFile)
    lblSize = ": " & Format(GetSizeOfFile(strFile) \ 1024, "###,###") & " KB"
    lblAttributes = ": " & GetAttribute(strFile)
    lblCreated = ": " & FormatDateTime(FileDateTime(strFile), vbLongDate) & ", " & FormatDateTime(FileDateTime(strFile), vbLongTime)
    lblPID = ": " & frmScanVirus.lstView.SelectedItem.SubItems(5)
    lblBase = ": " & frmScanVirus.lstView.SelectedItem.SubItems(6)
    lblThreads = ": " & frmScanVirus.lstView.SelectedItem.SubItems(7)
    lblMemory = ": " & frmScanVirus.lstView.SelectedItem.SubItems(11)
    lblPriority = ": " & frmScanVirus.lstView.SelectedItem.SubItems(9)
    'lblStartedOn = ": " + CStr(ST.wMonth) + "/" + CStr(ST.wDay) + "/" + CStr(ST.wYear) + " " + CStr(ST.wHour) + ":" + CStr(ST.wMinute) + "." + CStr(ST.wSecond)
    lblStartedOn = ": " & FileInfo.DateLastAccessed
    lblDateModif = ": " & FileInfo.DateLastModified

    hIcoExt = ExtractIcon(Me.hWnd, strFile, 0)
    hIcoDraw = DrawIcon(picIco.hdc, 0, 0, hIcoExt)

End Sub

Private Function GetModuleProcessID(lvwProc As ListView, ItemProcID As Integer, lvwModule As ListView, ilsModule As ImageList)

    On Error Resume Next

    Dim ExePath As String
    Dim uProcess As MODULEENTRY32
    Dim hSnapShot As Long
    Dim hPID As Long
    Dim lMod As Long
    Dim intLVW As Integer
    Dim I As Integer
    Dim lvwItem As ListItem
    Dim hVer As VERHEADER
    Dim strFile As String

    hPID = lvwProc.SelectedItem.SubItems(ItemProcID)

    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, hPID)
    uProcess.dwSize = Len(uProcess)
    lMod = Module32First(hSnapShot, uProcess)

    lvwModule.ListItems.Clear
    ilsModule.ListImages.Clear

    I = 0

    Do While lMod
        I = I + 1
        ExePath = uProcess.szExePath
        GetVerHeader ExePath, hVer
        strFile = Left(uProcess.szModule, InStr(uProcess.szModule, Chr(0)) - 1)
        ilsModule.ListImages.Add I, , GetIco.Icon(ExePath, SmallIcon)
        Set lvwItem = lvwMod.ListItems.Add(, , strFile, , 2)
            lvwItem.SubItems(1) = GetFilePath(StripNulls(ExePath))
            lvwItem.SubItems(2) = GetPathType(ExePath)
            lvwItem.SubItems(3) = hVer.CompanyName
            lvwItem.SubItems(4) = hVer.FileDescription
            lvwItem.SubItems(5) = hVer.FileVersion
            lvwItem.SubItems(6) = hVer.InternalName
            lvwItem.SubItems(7) = hVer.LegalCopyright
            lvwItem.SubItems(8) = hVer.LegalTradeMarks
            lvwItem.SubItems(9) = hVer.OrigionalFileName
            lvwItem.SubItems(10) = hVer.ProductName
            lvwItem.SubItems(11) = hVer.ProductVersion
            lvwItem.SubItems(12) = hVer.Comments
        lMod = Module32Next(hSnapShot, uProcess)
    Loop

    Call CloseHandle(hSnapShot)

End Function

Private Function GetSizeOfFile(ByVal PathFile As String) As Long

    Dim hFile As Long, OFS As OFSTRUCT

    hFile = OpenFile(PathFile, OFS, 0)
    GetSizeOfFile = GetFileSize(hFile, 0)

    Call CloseHandle(hFile)

End Function
