VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Mania V1.2"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6585
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3840
      Pattern         =   "*.mid;*.MID"
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer TmrMusic 
      Interval        =   16000
      Left            =   3720
      Top             =   4920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   4575
      Begin VB.Image Image7 
         Height          =   480
         Left            =   2280
         Picture         =   "FrmMain.frx":0ECA
         ToolTipText     =   "..Msg. Coder Room.."
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   1560
         Picture         =   "FrmMain.frx":0FA7
         Stretch         =   -1  'True
         ToolTipText     =   "..Picture Convertor Room.."
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   3960
         Picture         =   "FrmMain.frx":12B1
         Stretch         =   -1  'True
         ToolTipText     =   "..LOG Room.."
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   520
         Left            =   3120
         Picture         =   "FrmMain.frx":15BB
         Stretch         =   -1  'True
         ToolTipText     =   "..Password Generator Room.."
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   840
         Picture         =   "FrmMain.frx":2485
         Stretch         =   -1  'True
         ToolTipText     =   "..Files Room.."
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrmFiles 
      BackColor       =   &H00C00000&
      Caption         =   "Files Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   6375
      Begin VB.PictureBox shpbar 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   4080
         ScaleHeight     =   375
         ScaleWidth      =   975
         TabIndex        =   85
         Top             =   3480
         Width           =   975
      End
      Begin LockMania.chameleonButton Command1 
         Height          =   375
         Left            =   3960
         TabIndex        =   60
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Open File To Unlock"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":334F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdOpen 
         Height          =   375
         Left            =   1680
         TabIndex        =   59
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Open File To Lock"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":336B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdOpen1 
         Height          =   375
         Left            =   4920
         TabIndex        =   58
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Open"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":3387
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdClear 
         Height          =   495
         Left            =   4320
         TabIndex        =   57
         ToolTipText     =   "Click Here To Cancel Every Thing"
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Clear"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":33A3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdStop 
         Height          =   495
         Left            =   2160
         TabIndex        =   56
         ToolTipText     =   "Click Here To Stop the procces."
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Stop"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":33BF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdDo 
         Height          =   495
         Left            =   120
         TabIndex        =   55
         ToolTipText     =   "Click here 2 Start Our journy"
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Lock"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":33DB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Hide Key"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Left            =   3960
         TabIndex        =   18
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox TxtKeyf 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   9
         PasswordChar    =   "*"
         TabIndex        =   11
         ToolTipText     =   "u can enter a key of 9 numbers  without reapeating any number"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox TxtToFile 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "The new File Path is displayed Here"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox TxtFile 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "The File That u Want to lock or unlock is displayed here."
         Top             =   960
         Width           =   5415
      End
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblPer 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Progress In Percentage :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label LblStat 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ReadY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   6135
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000080&
         Caption         =   "Key :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000080&
         Caption         =   "New File Path :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "File :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   4080
         Top             =   3480
         Width           =   2175
      End
   End
   Begin VB.Frame FrmMsg 
      BackColor       =   &H00C00000&
      Caption         =   "Msg. Coder Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6375
      Begin LockMania.chameleonButton CmdMsgClr 
         Height          =   495
         Left            =   4680
         TabIndex        =   63
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Clear All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":33F7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdMsgDec 
         Height          =   495
         Left            =   4680
         TabIndex        =   62
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Decrypt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":3413
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton CmdMsgEnc 
         Height          =   495
         Left            =   4680
         TabIndex        =   61
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Encrypt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":342F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtKey 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox TxtMsg 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000080&
         Caption         =   "Orders :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000080&
         Caption         =   "Key : But Remeber (longer Key means      .Longer Result)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   4335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000080&
         Caption         =   "Write Your Message Here :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         X1              =   4560
         X2              =   4560
         Y1              =   720
         Y2              =   3840
      End
      Begin VB.Image Image8 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C00000&
      Caption         =   "Help Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   79
      Top             =   960
      Width           =   6375
      Begin LockMania.chameleonButton cmdhlp 
         Height          =   375
         Left            =   5040
         TabIndex        =   84
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "GO >>>"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":344B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txthlp 
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   83
         Top             =   1560
         Width           =   6135
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   82
         Text            =   "Combo2"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Image Image15 
         Height          =   615
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "How Do I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   81
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Please Choose A Category And Then Click Go:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   80
         Top             =   480
         Width           =   4965
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Log Viewer Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   6375
      Begin LockMania.chameleonButton Command4 
         Height          =   735
         Left            =   4920
         TabIndex        =   66
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         BTYPE           =   3
         TX              =   "Save Log As txt File"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":3467
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton Command3 
         Height          =   375
         Left            =   4920
         TabIndex        =   65
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Clear Log"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":3483
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton Command2 
         Height          =   375
         Left            =   4920
         TabIndex        =   64
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":349F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtReport 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   20
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Orders :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   4920
         TabIndex        =   86
         Top             =   840
         Width           =   1020
      End
      Begin VB.Line Line13 
         BorderColor     =   &H0000FFFF&
         X1              =   4800
         X2              =   4800
         Y1              =   840
         Y2              =   3840
      End
      Begin VB.Image Image11 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Log  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C00000&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   49
      Top             =   960
      Width           =   6375
      Begin VB.Label Label24 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Believe it Or Not :    But It Was Made In Egypt  !!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   975
         Left            =   4080
         TabIndex        =   54
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Image Image14 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   4080
         Picture         =   "FrmMain.frx":34BB
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email Me For Support :             a_kotb2003@yahoo.com ""Or""   kotbcorp@gmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   5895
      End
      Begin VB.Line Line14 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   3840
         X2              =   3840
         Y1              =   1800
         Y2              =   3720
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lock Mania 1.2 By Ahmed Kotb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Credits :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FrmMain.frx":35E7
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   120
         TabIndex        =   50
         Top             =   2040
         Width           =   3375
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Key Generator Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   43
      Top             =   960
      Width           =   6375
      Begin LockMania.chameleonButton Command16 
         Height          =   495
         Left            =   3960
         TabIndex        =   68
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Copy"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":3683
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin LockMania.chameleonButton Command15 
         Height          =   495
         Left            =   960
         TabIndex        =   67
         Top             =   2160
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Generate"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":369F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   510
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   3240
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   44
         Text            =   "Combo1"
         Top             =   960
         Width           =   3135
      End
      Begin VB.Image Image13 
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000FFFF&
         X1              =   5280
         X2              =   5280
         Y1              =   600
         Y2              =   3720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Step 3 : Just Watch You Password Generated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   47
         Top             =   2880
         Width           =   4890
      End
      Begin VB.Line Line11 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   5040
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Step 2 : Hit ""Generate Key"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   46
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Line Line10 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   5040
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Step 1 : Please Choose The key Generator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   4530
      End
   End
   Begin VB.Frame FrmPic 
      BackColor       =   &H00C00000&
      Caption         =   "Picture Convertor Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   6375
      Begin VB.OptionButton Option2 
         BackColor       =   &H000080FF&
         Caption         =   "&Extract Files From  Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000080FF&
         Caption         =   "&Convert Files To Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Frame FrmPicAdd 
         BackColor       =   &H00C00000&
         Caption         =   "Convert Files To Pictures"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   3135
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   4335
         Begin LockMania.chameleonButton Command5 
            Height          =   375
            Left            =   2520
            TabIndex        =   74
            Top             =   2520
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Open"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":36BB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command8 
            Height          =   375
            Left            =   1800
            TabIndex        =   73
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Clear All"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":36D7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command10 
            Height          =   375
            Left            =   3240
            TabIndex        =   72
            Top             =   2520
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Stop"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":36F3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command9 
            Height          =   375
            Left            =   3240
            TabIndex        =   71
            Top             =   2040
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Start"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":370F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command7 
            Height          =   495
            Left            =   3240
            TabIndex        =   70
            Top             =   1320
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Remove File"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":372B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command6 
            Height          =   375
            Left            =   3240
            TabIndex        =   69
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Add File"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":3747
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox TxtPicFile 
            BackColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   2520
            Width           =   2415
         End
         Begin VB.ListBox List1 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   2895
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   3120
            X2              =   4200
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   3120
            X2              =   4200
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   3120
            X2              =   4200
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   3120
            X2              =   3120
            Y1              =   600
            Y2              =   3000
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            Caption         =   "Picture File :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   28
            Top             =   2040
            Width           =   1410
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   4200
            X2              =   4200
            Y1              =   600
            Y2              =   3000
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            Caption         =   "Files :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame FrmPicExt 
         BackColor       =   &H00C00000&
         Caption         =   "Extract Files From A Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   3135
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   4335
         Begin LockMania.chameleonButton Command14 
            Height          =   375
            Left            =   3120
            TabIndex        =   78
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Clear All"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":3763
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command13 
            Height          =   495
            Left            =   3120
            TabIndex        =   77
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Stop"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":377F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command12 
            Height          =   495
            Left            =   3120
            TabIndex        =   76
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Start"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":379B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin LockMania.chameleonButton Command11 
            Height          =   255
            Left            =   3000
            TabIndex        =   75
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "Open"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmMain.frx":37B7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtpicextoutput 
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtpicpass 
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox TxtPicfile2 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   720
            Width           =   2775
         End
         Begin VB.Line Line9 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   3000
            X2              =   3000
            Y1              =   1200
            Y2              =   3000
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            Caption         =   "File Name :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   1185
         End
         Begin VB.Line Line8 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   3000
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            Caption         =   "Password :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Line Line7 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   4200
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00000080&
            Caption         =   "Picture File :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1290
         End
      End
      Begin VB.Image Image12 
         Height          =   615
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lplpictotper 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4560
         TabIndex        =   35
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Total File Percentage :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4560
         TabIndex        =   34
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "File Percentage :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   4560
         TabIndex        =   33
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label Lblpicper 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4560
         TabIndex        =   32
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label LblPicStat 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   4560
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "Progress Data :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   4560
         TabIndex        =   30
         Top             =   960
         Width           =   1650
      End
   End
   Begin VB.Image Image10 
      Height          =   735
      Left            =   120
      Picture         =   "FrmMain.frx":37D3
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6375
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   4800
      Picture         =   "FrmMain.frx":3FA4
      Stretch         =   -1  'True
      ToolTipText     =   "..About.."
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   5640
      Picture         =   "FrmMain.frx":4E6E
      Stretch         =   -1  'True
      ToolTipText     =   "..Quit.."
      Top             =   5040
      Width           =   720
   End
   Begin VB.Menu mnufile 
      Caption         =   "Main"
      WindowList      =   -1  'True
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnumainhelp 
         Caption         =   "Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnusls 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "Tools"
      Begin VB.Menu mnufiletool 
         Caption         =   "File Locker Tool"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnupictool 
         Caption         =   "Picture Convertor Tool"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnumsgtool 
         Caption         =   "Msg. Coder Tool"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnukeytool 
         Caption         =   "Key Generator Tool"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnulogtool 
         Caption         =   "Log Viewer Tool"
         Shortcut        =   ^L
      End
      Begin VB.Menu sls2 
         Caption         =   "-"
      End
      Begin VB.Menu CmdMusic 
         Caption         =   "Stop Music"
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnupophelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnupopabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LstFile As String
Dim D
Dim O
Dim fpps As String
Dim Result, aa
Dim Quant
Dim Chars
Dim UQuant As Single
Dim MsTest As Integer
Public DevBarVal1
Private Function barval(TheVal As Long)
shpbar.Width = (TheVal / DevBarVal1)
End Function
Private Sub Check1_Click()
If Check1.Value = 1 Then TxtKeyf.PasswordChar = "*"
If Check1.Value = 0 Then TxtKeyf.PasswordChar = ""
TxtKeyf.SetFocus
End Sub

Public Sub CmdClear_Click()
CmdDo.Enabled = False
CmdOpen1.Enabled = False
TxtFile.Text = ""
TxtKeyf.Text = ""
TxtToFile.Text = ""
LblStat.Caption = ""
LblPer.Caption = ""
CmdStop.Enabled = False
CmdClear.Enabled = True
CmdOpen.Enabled = True
Command1.Enabled = True
TxtKeyf.Locked = False
LblStat.Caption = "ReadY"
shpbar.Width = 1
End Sub

Private Sub CmdDo_Click()
If TxtToFile.Text = "" Then MsgBox "Please Enter New File Path", vbCritical: Exit Sub
If Dir(TxtToFile.Text) <> "" Then
MsgBox "Please The New File Path is Already Exists So Please Change It", vbCritical
CmdOpen1.Enabled = True
Exit Sub
ElseIf TxtKeyf.Text = "" Then
MsgBox "please Write A key", vbCritical, "Missing Key"
Exit Sub
End If
ClearStrs
LblStat.Caption = "Gathering Data..."
CmdOpen.Enabled = False
Command1.Enabled = False
CmdOpen1.Enabled = False
CmdDo.Enabled = False
CmdStop.Enabled = True
CmdClear.Enabled = False
Select Case CmdDo.Caption
Case "&Lock File"
 StopFor = False
 If checkpassword(TxtKeyf.Text) = False Then MsgBox KeyError, vbCritical, "ERROR :": CmdDo.Enabled = True: CmdStop.Enabled = False: LblStat.Caption = "Invalid Key": Exit Sub
 LblStat.Caption = "Locking File...Please Wait..."
 q = RecsNum(FileLen(TxtFile.Text), TxtKeyf.Text)
 If q <> 0 Then
 Ext = Right(TxtFile.Text, 3)
 DOOPERATION q, TxtKeyf.Text, True, TxtFile.Text, TxtToFile.Text
 Write2Log (Now & vbCrLf & "Lock Mania File Tool Used To lock File.." & vbCrLf & "File Name :" & TxtFile.Text & vbCrLf & "New File Path :" & TxtToFile.Text & vbCrLf & Report)
 MsgBox "Operation Done", vbInformation
 LblStat.Caption = "ReadY"
 If Len(Dir(TMPFILE)) <> 0 Then Kill TMPFILE
 ElseIf q = 0 Then
 LblStat.Caption = "Optimizing Data..."
 FileCopy TxtFile.Text, TMPFILE
 Open TMPFILE For Binary As #11
 For i = 1 To Int(Remain)
 M = 255
 Put 11, LOF(11) + 1, M
 Next i
 Close #11
 q = RecsNum(FileLen(TMPFILE), TxtKeyf.Text)
 'MsgBox q
 Ext = Right(TxtFile.Text, 3)
 LblStat.Caption = "Locking File...Please Wait..."
 DOOPERATION q, TxtKeyf.Text, True, TMPFILE, TxtToFile.Text
 
 LblStat.Caption = "Deleting Temporary Files.."
 If Len(Dir(TMPFILE)) <> 0 Then Kill TMPFILE
 End If
 'Close
 Write2Log (Now & vbCrLf & "Lock Mania File Tool Used To lock File.." & vbCrLf & "File Name :" & TxtFile.Text & vbCrLf & "New File Path :" & TxtToFile.Text & vbCrLf & Report)
 MsgBox "Operation Done", vbInformation
 LblStat.Caption = "Ready"
Case "&UnLock File" '######################## unlock file
 LblStat.Caption = "Gathering Data..."
 StopFor = False
 PsWrd = ""
 Open TxtFile.Text For Binary As #4
 Get 4, LOF(4), M
 tmplenkey = Chr(M)
 'MsgBox "key=" & tmplenkey
 Select Case tmplenkey
 Case 1 To 9
 LenKey = Int(tmplenkey)
 Case Else
 MsgBox "This File Is Corrupted", vbCritical, "Sorry..."
 Close #4
 CmdClear_Click
 Exit Sub
 End Select
 For i = 1 To LenKey
 Get 4, i, M
 M = M - 150
 PsWrd = PsWrd & Chr(M)
 Next i
 Close #4
 'MsgBox PsWrd
 If TxtKeyf.Text <> PsWrd Then MsgBox "Acces Denied" & vbCrLf & "WRONG PASSWORD", vbCritical, "BAD PASSWORD": CmdClear.Enabled = True: CmdDo.Enabled = True: CmdStop.Enabled = False: LblStat.Caption = "Wrong Password": Exit Sub
 Open TxtFile.Text For Binary As #4
 For i = 1 To 3
 Get 4, Len(TxtKeyf.Text) + i, M
 Ext = Ext & Chr(M)
 Next i
q = (FileLen(TxtFile.Text) - Len(TxtKeyf.Text) - 3 - 1) / Len(TxtKeyf.Text)
'MsgBox q
Close #4
 LblStat.Caption = "Unlocking File...Please Wait"
 DOOPERATION q, TxtKeyf.Text, False, TxtFile.Text, TMPFILE
 If StopFor = True Then Exit Sub
 LblStat.Caption = "Processing New File..Please Wait"
 Open TMPFILE For Binary As #11
 For i = LOF(11) - 7 To LOF(11)
 Get 11, i, M
 Rmvstr = Rmvstr & Chr(M)
 Next i
'MsgBox "Rmvstr=" & Rmvstr & vbCrLf & Len(Rmvstr)
 RmvVal = InStr(1, Rmvstr, Chr(255))
 If RmvVal = 0 Then
 Close #11
 Name TMPFILE As TxtToFile.Text & "." & Ext
 MsgBox "Operation Done", vbInformation
 Write2Log (Now & vbCrLf & "Lock Mania File Tool Used To UnLock File.." & vbCrLf & "File Name :" & TxtFile.Text & vbCrLf & "New File Path :" & TxtToFile.Text & vbCrLf & Report)
 CmdClear_Click
 Exit Sub
 Else
 RmvVal = 8 - RmvVal
 RmvVal = LOF(11) - RmvVal - 1
 LblStat.Caption = "Writing Data...Patience Please"
 Open TxtToFile.Text & "." & Ext For Binary As #12
 For i = 1 To RmvVal
 Get 11, i, M
 Put 12, i, M
 DoEvents
 LblPer.Caption = Int(LOF(12) / LOF(11) * 100)
 If StopFor = True Then
   msg = MsgBox("Are YOu Sure That You Want TO Stop Operations ?", vbYesNo + vbExclamation, "STOP!!")
 Select Case msg
  Case vbYes
    Close #11
    Close #12
    StopOperation
    Exit Sub
  Case vbNo
    StopFor = False
  End Select
  End If
 Next i
 Close #11
 Close #12
 LblStat.Caption = "Deleting Temporary Files..."
'On Error Resume Next
 Kill TMPFILE
 Write2Log (Now & vbCrLf & "Lock Mania File Tool Used To UnLock File.." & vbCrLf & "File Name :" & TxtFile.Text & vbCrLf & "New File Path :" & TxtToFile.Text & vbCrLf & Report)
 MsgBox "Operation Done", vbInformation
 End If
End Select
'################### clear
CmdClear_Click
' ################## END
End Sub

Private Sub cmdhlp_Click()
Select Case Combo2.ListIndex
Case 0
Image15.Picture = Image1.Picture
txthlp.Text = "Locking A File is So Easy Here is the Steps" & vbCrLf & "Steps :" & vbCrLf & "1.Go to Files Room Then Click Open button File For Lock." & vbCrLf & "2.Select the file u want 2 lock then click open button " & vbCrLf & "3. Click Open button to select the new file path" & vbCrLf & "4.Enter A key up to 9 numbers" & vbCrLf & "5. Click Lock Then Wait Untill The File is locked"
Case 1
Image15.Picture = Image1.Picture
txthlp.Text = "UnLocking A File is So Easy Here is the Steps" & vbCrLf & "Steps :" & vbCrLf & "1.Go to Files Room Then Click Open button File For UnLock." & vbCrLf & "2.Select the file u want 2 Unlock" & vbCrLf & "3. Click Open button to select the new file path" & vbCrLf & "4.Enter the file key" & vbCrLf & "5. Click UnLock Then Wait Untill The File is Unlocked"
Case 2
Image15.Picture = Image4.Picture
txthlp.Text = "Key Generator Utility helps u 2 generate passwords" & vbCrLf & "Steps :" & vbCrLf & "1.Choose the password length." & vbCrLf & "2.Click Generate." & vbCrLf & "3.that is it & u can Copy the result by clicking Copy button"
Case 3
Image15.Picture = Image6.Picture
txthlp.Text = "please Read The Steaps Carefully" & vbCrLf & "1.open pictures Conv room" & vbCrLf & "2.Click Convert Files To A picture." & vbCrLf & "3.Click Add to add files to the list u can do it more than 1 time." & vbCrLf & "4.u can remove a file by clicking remove button" & vbCrLf & "5.Click open button to Choosethe picture then click start" & "7.the program will give u a password for each file u can click yes and then open notepad or wordpad and choose paste" & vbCrLf & "8.the program will ask u if u want to delete the files u r free!!"
Case 4
Image15.Picture = Image6.Picture
txthlp.Text = "Steps :" & vbCrLf & "1.open pictures Conv room" & vbCrLf & "2.click Extract file from a picture" & vbCrLf & "3.click open 2 choose the picture" & vbCrLf & "3.enter file password then enter a file name then click Start"
Case 5
Image15.Picture = Image7.Picture
txthlp.Text = "to Encrypt A message " & vbCrLf & "1.Enter ur message then enter a key then click Encrypt" & vbCrLf & "To Decrypt A message" & vbCrLf & "1.Enter the encrypted message then enter the key then click Decrypt"
End Select
End Sub

Private Sub CmdMsgClr_Click()
TxtMsg.Text = ""
TxtKey.Text = ""
End Sub

Private Sub CmdMsgDec_Click()
If TxtKey.Text = "" Then MsgBox "Please Write A key", vbExclamation, "Key is Missing": Exit Sub
CmdMsgEnc.Enabled = False
CmdMsgDec.Enabled = False
CmdMsgClr.Enabled = False
TxtMsg.Text = ToStr(TxtMsg, TxtKey)
CmdMsgClr.Enabled = True
CmdMsgEnc.Enabled = True
CmdMsgDec.Enabled = True
End Sub

Private Sub CmdMsgEnc_Click()
If TxtKey.Text = "" Then MsgBox "Please Write A key", vbExclamation, "Key is Missing": Exit Sub
CmdMsgEnc.Enabled = False
CmdMsgDec.Enabled = False
CmdMsgClr.Enabled = False
TxtMsg.Text = ToHex(TxtMsg, TxtKey)
CmdMsgClr.Enabled = True
CmdMsgEnc.Enabled = True
CmdMsgDec.Enabled = True
End Sub
Private Sub CmdMusic_Click()
Select Case CmdMusic.Caption
Case "Stop Music"
StopMusic
TmrMusic.Enabled = False
CmdMusic.Caption = "Resume Music"
Case "Resume Music"
MsgBox "Music Will Be Resumed After Few Seconds", vbInformation, "Music :"
TmrMusic.Enabled = True
CmdMusic.Caption = "Stop Music"
CmdMusic.Enabled = False
End Select
End Sub


Private Sub CmdOpen_Click()
Dlg1.Filter = "All Files(*.*)|*.*"
TxtFile.Text = Dlg1.FileOpen
If TxtFile.Text <> "" Then
CmdDo.Caption = "&Lock File"
CmdDo.Enabled = True
CmdOpen1.Enabled = True
End If
End Sub

Private Sub CmdOpen1_Click()
Dim ex As String
Select Case CmdDo.Caption
Case "&Lock File"
Dlg1.Filter = "Lock Mania Encrypted File(*.LMF)|*.LMF"
Dlg1.DefaultExtension = "*.LMF"
TxtToFile.Text = Dlg1.FileSave
If TxtToFile.Text = TxtFile.Text And Right(TxtToFile.Text, 3) = Right(TxtFile.Text, 3) Then
MsgBox "You Can not Use the same File The application will change the file name as " & TxtToFile.Text & ".LMF" & "  if that doesnt Satisfy you then Choose Another One", vbCritical, "Choose Another File"
TxtToFile.Text = TxtToFile.Text & ".LMF"
End If
If TxtToFile.Text <> "" Then
If InStr(1, TxtToFile.Text, ".LMF") = 0 Then TxtToFile.Text = TxtToFile.Text & ".LMF"
End If
Case "&UnLock File"
Dlg1.Filter = "All Files(*.*)|*.*"
Dlg1.DefaultExtension = ""
TxtToFile.Text = Dlg1.FileSave
If TxtToFile.Text = TxtFile.Text And Right(TxtToFile.Text, 3) = Right(TxtFile.Text, 3) Then
MsgBox "You Can not Use the same File Please Choose Another One", vbCritical, "Choose Another File"
TxtToFile.Text = ""
End If
End Select
ex = Dir(TxtToFile.Text)
If ex <> "" Then MsgBox "You Can not Use A file That is Already Exists..." & vbCrLf & "Please Choose Another One", vbCritical: TxtToFile.Text = "": Exit Sub
End Sub
Private Sub CmdStop_Click()
StopFor = True
End Sub

Private Sub Command1_Click()
Dlg1.Filter = "Lock Mania Encrypted File(*.LMF)|*.LMF"
TxtFile.Text = Dlg1.FileOpen

If TxtFile.Text <> "" Then
CmdDo.Caption = "&UnLock File"
CmdDo.Enabled = True
CmdOpen1.Enabled = True
End If
End Sub

Private Sub Command10_Click()
StopDO = True
End Sub

Private Sub Command11_Click()
LblPicStat.Caption = "Status : " & "Opening Client Picture File"
Dlg1.Filter = "ALL PICTURE FORMATS(*.jpg;*.bmp;*.gif)|*.jpg;*.bmp;*.gif|BITMAPS(*.bmp)|*.bmp;*.BMP)|JPEG PICS(*.jpg)|*.jpg;*.JPG)"
TxtPicfile2.Text = Dlg1.FileOpen
Command12.Enabled = True
LblPicStat.Caption = "Status : Ready"
If TxtPicfile2.Text = "" Then Command12.Enabled = False
End Sub

Private Sub Command12_Click()
If Command10.Enabled = True Then MsgBox "You Cant Add & Extract Files in the Same Time Please Wait Until The Extracted File Finished Then Come and Some Files.", vbCritical: Exit Sub
If txtpicpass.Text = "" Or txtpicextoutput.Text = "" Then
MsgBox "One Of the password or the filename spaces is empty", vbCritical
Exit Sub
End If
TmpToFilePath = InStrRev(TxtPicfile2.Text, "\", -1)
TmpToFilePath = Left(TxtPicfile2.Text, TmpToFilePath)
TmpToFilePath = TmpToFilePath & txtpicextoutput.Text
If InStr(1, txtpicpass.Text, "@") = 0 Or InStr(1, txtpicpass.Text, "(") = 0 Or InStr(1, txtpicpass.Text, ")") = 0 Or Right(txtpicpass.Text, 1) <> ")" Then
MsgBox "invalid Password", vbCritical
Exit Sub
Else
msg = MsgBox("Are You Sure From Your Password ?" & vbCrLf & "Remeber if it is Wrong The Program Will Produce A Damaged File And It May Cause Errors", vbInformation + vbYesNo, "Last Warning :")
If msg = vbYes Then
Command11.Enabled = False
Command13.Enabled = False
Command14.Enabled = False
MsgBox extractfile(TxtPicfile2.Text, TmpToFilePath, txtpicpass.Text), vbInformation, "MISSION ACCOMPLISHED"
Command11.Enabled = True
Command13.Enabled = True
Command14.Enabled = True
Command14_Click
LblPicStat.Caption = "Status : Ready"
End If
End If
End Sub

Public Sub Command14_Click()
TxtPicfile2.Text = ""
txtpicpass.Text = ""
txtpicextoutput.Text = ""
LblStat.Caption = "READY"
Lblpicper.Caption = ""
lplpictotper.Caption = ""
Command11.Enabled = True
Command12.Enabled = False
Command13.Enabled = False
End Sub

Private Sub Command15_Click()
Result = ""
If Combo1.Text = "" Then
MsgBox "Select the password Length please"
Exit Sub
End If
Randomize
UQuant = Combo1.Text
Quant = 0
Do Until UQuant = Quant
Run:
aa = Int(Rnd * 122) + 1
If aa < 48 Or _
aa > 57 And _
aa < 65 Or _
aa > 90 And _
aa < 97 Or _
aa > 122 Then
GoTo Run
End If
Chars = Chr(aa)
Result = Result & Chars
Quant = Quant + 1
Loop
Text1.Text = Result
Write2Log (Now & vbCrLf & "Key Generator Used To Generate a key")
End Sub

Private Sub Command16_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
MsgBox "The Key Was Send To Clipboard", vbInformation, "Data Transfer Complete.."
End Sub

Private Sub Command2_Click()
txtReport.Text = loadLog
End Sub

Private Sub Command3_Click()
msg = MsgBox("Are You Sure That you want To Clear Log ?", vbYesNo + vbExclamation, "CLEAR>>")
If msg = vbYes Then
Clearlog
txtReport.Text = loadLog
End If
End Sub

Private Sub Command4_Click()
Dim dest As String
Dlg1.Filter = "Txt Files(*.txt)|*.txt"
Dlg1.DefaultExtension = "*.txt"
dest = Dlg1.FileSave
If dest <> "" Then
Command2_Click
Open dest For Binary As #45
For logi = 1 To Len(txtReport.Text)
LogM = Asc(Mid(txtReport.Text, logi, 1))
Put #45, logi, LogM
Next logi
Close #45
MsgBox "Log Saved To File :" & dest, vbInformation, "Log Saved"
End If
End Sub
Private Sub Command5_Click()
Command5.Enabled = False
LblPicStat.Caption = "Status : " & "Opening Client Picture File"
Dlg1.Filter = "ALL PICTURE FORMATS(*.jpg;*.bmp;*.gif)|*.jpg;*.bmp;*.gif|BITMAPS(*.bmp)|*.bmp;*.BMP)|JPEG PICS(*.jpg)|*.jpg;*.JPG)"
TxtPicFile.Text = Dlg1.FileOpen
If TxtPicFile.Text <> "" Then Command9.Enabled = True
Command5.Enabled = True
LblPicStat.Caption = "Status : ReadY"
End Sub
Private Sub Command6_Click()
LstFile = Dlg1.FileOpen
If LstFile = "" Then Exit Sub
If LstFile <> "" Then
List1.AddItem LstFile
End If
End Sub

Private Sub Command7_Click()
If List1.ListIndex <> -1 Then
List1.RemoveItem List1.ListIndex
End If
End Sub

Public Sub Command8_Click()
List1.Clear
TxtPicFile.Text = ""
Lblpicper.Caption = ""
lplpictotper.Caption = ""
LblPicStat.Caption = "Ready"
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = False
Command10.Enabled = False
End Sub

Private Sub Command9_Click()
If Command13.Enabled = True Then MsgBox "You Cant Add & Extract Files in the Same Time Please Wait Until The Extracted File Finished Then Come and Some Files.", vbCritical: Exit Sub
StopDO = False
LblPicStat.Caption = "Making BackUp"
MsgBox "The Program Will Make A BakeUp For The Client Picture Becouse If You Stop The Process The Picture Will Be Deleted And Replaced With The Bakeup", vbInformation, "BakeUP"
FileCopy TxtPicFile.Text, BKP
LblPicStat.Caption = "Status : Checking Data..."
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = True
If List1.ListCount = 0 Or TxtPicFile.Text = "" Then
MsgBox "You Might Have Forgot to Add Files Or To Choose A Picture To Convert Files To It", vbCritical, "Data Error :"
LblPicStat.Caption = ""
Command10.Enabled = False
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command9.Enabled = True
Exit Sub
End If
TotalWrittenBytes = 0
'step 1
LblPicStat.Caption = "Status : Checking Files..."
For Pici = 0 To List1.ListCount - 1
D = ""
D = Dir(List1.List(Pici))
If D = "" Then
MsgBox "File : " & List1.List(i) & "  Was Not Found Please Remove It.", vbCritical
Exit Sub
End If
Next Pici
'step 2
LblPicStat.Caption = "Status : Gathering Data..."
For Pici = 0 To List1.ListCount - 1
O = FileLen(List1.List(Pici))
TotalFileLen = TotalFileLen + O
Next Pici
LblPicStat.Caption = "Status : Starting..."

picReport = "Picture Convertor Report" & vbCrLf & "This Report is very important becouse it contain files password" & vbCrLf & "Client Picture File : " & TxtPicFile.Text

For Pici = 0 To List1.ListCount - 1        'adding files
fpps = ""
LblPicStat.Caption = "Adding File : " & List1.List(Pici)
fpps = AddFile(TxtPicFile.Text, List1.List(Pici))
picReport = picReport & vbCrLf & "File : " & List1.List(Pici) & " / password : " & fpps
DoEvents
If StopDO = True Then
Command8_Click
Exit Sub
End If
Next Pici
msg = MsgBox("Do You Want To Delete Source Files ? ", vbYesNo, "TAKING PERMISSION...")
If msg = vbYes Then
LblPicStat.Caption = "Status : Deleting Source Files"
For Pici = 0 To List1.ListCount - 1
Kill List1.List(Pici)
Next Pici
End If
LblPicStat.Caption = "Status :ReadY"
Command8_Click
msg = MsgBox(picReport & vbCrLf & vbCrLf & "Press Yes To Copy That Report Or No To Continue", vbInformation + vbYesNo, "Report :")
If msg = vbYes Then
Clipboard.SetText picReport
MsgBox "Data Copied , Please paste it on word proccesor as notepad or MicroSoft Word ", vbInformation, "DONE:"
End If
If Dir(BKP) <> "" Then Kill BKP
End Sub

Private Sub Form_Activate()
If MsTest = 0 Then
On Error GoTo error:
FrmMain.File1.Path = App.Path & "\MusicSys\"
Randomize
song = File1.Path & "\" & File1.List(Int((File1.ListCount - 1) * Rnd) + 2)
FileCopy song, "C:\tmp.dat"
PlayMidiFile "c:\tmp.dat"
Exit Sub
error:
MsgBox "Music Folder Was Not Found...Music Will Not Be Played", vbCritical
TmrMusic.Enabled = False
CmdMusic.Enabled = False
MsTest = 1
Write2Log Now & vbCrLf & "Music Folder Was Not Found"
End If
End Sub
Private Sub Form_Load()
For i = 1 To 9
Combo1.AddItem i
Next i
Combo2.AddItem "Lock File ?"
Combo2.AddItem "Unlock File ?"
Combo2.AddItem "Generate A Key ?"
Combo2.AddItem "Add Files To A Picture ?"
Combo2.AddItem "Extract Files From A Picture ?"
Combo2.AddItem "Encrypt And Decrypt A message"
Combo1.ListIndex = 7
Combo2.ListIndex = 0
cmdhlp_Click

Image8.Picture = Image7.Picture
Image9.Picture = Image1.Picture
Image11.Picture = Image5.Picture
Image12.Picture = Image6.Picture
Image13.Picture = Image4.Picture
Image1_Click
MsTest = 0

Write2Log Now & vbCrLf & "Program Started"
txtReport.Text = loadLog
DevBarVal1 = (100 / Shape1.Width)
barval (100)
GradientPlus shpbar, 50, 1, 1, 200, 300, 100
barval (0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
If CmdStop.Enabled = True Then
msg = MsgBox("Are u sure that u want to Stop Operations And Quit ?", vbExclamation + vbYesNo, "Stop & Quit")
If msg = vbYes Then
StopOperation
StopMusic
Write2Log Now & vbCrLf & "Program Closed"
MsgBox "Thank U 4 using My Applications" & vbCrLf & "a_kotb2003@yahoo.com", vbInformation
End
Else
Cancel = 1
Exit Sub
End If
End If
'#################################
If Command10.Enabled = True Or Command13.Enabled = True Then
msg = MsgBox("Are u sure that u want to Stop Operations And Quit ?", vbExclamation + vbYesNo, "Stop & Quit")
If msg = vbYes Then
StopPicOperations False
StopPicOperations True
StopMusic
Write2Log Now & vbCrLf & "Program Closed"
MsgBox "Thank U 4 using My Applications" & vbCrLf & "a_kotb2003@yahoo.com" & vbCrLf & "Kotbcorp@gmail.com", vbInformation
End
Else
Cancel = 1
Exit Sub
End If
End If
'#################################
msg = MsgBox("Are u sure that u want to quit ?", vbExclamation + vbYesNo, "Quit")
If msg = vbYes Then
StopMusic
Write2Log Now & vbCrLf & "Program Closed"
MsgBox "Thank U 4 using My Applications" & vbCrLf & "a_kotb2003@yahoo.com", vbInformation
End
Else
Cancel = 1
Exit Sub
End If
End Sub



Private Sub Image1_Click()
FrmFiles.ZOrder 0
End Sub

Private Sub Image2_Click()
If LblStat.Caption <> "ReadY" Then
msg = MsgBox("Are u sure that u want to Stop Operations And Quit ?", vbExclamation + vbYesNo, "Stop & Quit")
If msg = vbYes Then
StopOperation
End
Else
Exit Sub
End If
End If
msg = MsgBox("Are u sure that u want to quit ?", vbExclamation + vbYesNo, "Quit")
If msg = vbYes Then
StopMusic
MsgBox "Thank U 4 using My Applications" & vbCrLf & "a_kotb2003@yahoo.com" & vbCrLf & "kotbcorp@gmail.com", vbInformation
End
End If
End Sub

Private Sub Image3_Click()
Me.PopupMenu mnupop
End Sub
Private Sub Image4_Click()
Frame3.ZOrder 0
End Sub

Private Sub Image5_Click()
Frame2.ZOrder 0
End Sub

Private Sub Image6_Click()
FrmPic.ZOrder 0
End Sub

Private Sub Image7_Click()
FrmMsg.ZOrder 0
End Sub

Private Sub LblPer_Change()
If LblPer.Caption <> "" Then
barval (LblPer.Caption)
End If
End Sub
Private Sub LblStat_change()
Report = Report & LblStat.Caption & vbCrLf
End Sub


Private Sub mnuabout_Click()
Frame4.ZOrder 0
End Sub

Private Sub mnuexit_Click()
Image2_Click
End Sub

Private Sub mnufiletool_Click()
Image1_Click
End Sub

Private Sub mnukeytool_Click()
Image4_Click
End Sub

Private Sub mnulogtool_Click()
Image5_Click
End Sub

Private Sub mnumainhelp_Click()
Frame5.ZOrder 0
End Sub

Private Sub mnumsgtool_Click()
Image7_Click
End Sub

Private Sub mnupictool_Click()
Image6_Click
End Sub

Private Sub mnupopabout_Click()
Frame4.ZOrder 0
End Sub

Private Sub mnupophelp_Click()
Frame5.ZOrder 0
End Sub

Private Sub Option1_Click()
FrmPicAdd.ZOrder 0
End Sub

Private Sub Option2_Click()
FrmPicExt.ZOrder 0
End Sub

Private Sub TmrMusic_Timer()
CmdMusic.Enabled = True
Randomize
song = File1.Path & "\" & File1.List(Int((File1.ListCount - 1) * Rnd) + 2)
StopMusic
FileCopy song, "C:\Tmp.dat"
PlayMidiFile "c:\Tmp.dat"
End Sub

