VERSION 5.00
Begin VB.Form FrmStart 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   6285
   ClientTop       =   3210
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         ScaleHeight     =   375
         ScaleWidth      =   1575
         TabIndex        =   4
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   6615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   4965
         Left            =   120
         Picture         =   "FrmStart.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   6960
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p
Dim bar_value As Integer
Dim Se As String
Dim sm As Byte
Dim Myloc As String
Private Sub Form_Activate()
Dim str As String
Dim Size As Long
Size = 255
Dim L As Long
Dim winpath As String
    str = Space(Size)
    L = GetSystemDirectory(str, Size)
    winpath = str
    Label4.Caption = winpath
    Label4.Caption = Label4.Caption & "\LMLV1.2.LOG"
    LogPath = Label4.Caption
' here loading code
Label4.Caption = Label4.Caption & ".Dat"
Open Label4.Caption For Binary As #6
Get #6, 1, sm
Se = Chr(sm)
If Se <> "Y" Then
Dim objWshShell  As Object
Dim objShellLink As Object
Set objWshShell = CreateObject("WScript.Shell")
Dim strdesktop As String


Myloc = App.Path & "\" & App.EXEName & ".exe"
strdesktop = objWshShell.SpecialFolders("Desktop")
Set objShellLink = objWshShell.CreateShortcut(strdesktop & "\LockMania 1.2.lnk")
objShellLink.TargetPath = Myloc
   objShellLink.WindowStyle = 1
   objShellLink.Hotkey = "CTRL+SHIFT+L"
   objShellLink.IconLocation = Myloc
   objShellLink.Description = "This is my shortcut!"
   objShellLink.WorkingDirectory = strdesktop
   objShellLink.Save
   
   Set objShellLink = Nothing
   Set objWshShell = Nothing
   Se = "Y"
   sm = Asc(Se)
Put #6, 1, sm
Close
Timer1.Enabled = False
MsgBox "THIS IS YOU FIRST TIME TO OPEN LOCK MANIA 1.2 AND THE PROGRAM WILL MAKE A SHORTCUT TO ITSELF IN YOUR DESKTOP FOR THE FIRST TIME ONLY SO DONT WORRY YOU CAN DELETE IT. " & vbCrLf & "FINALLY LET ME SAY TO YOU .....THANKS.", vbInformation
Timer1.Enabled = True
Else
End If
End Sub
Private Sub Form_Load()
p = 100 / Label2.Width
putbarval (100)
GradientPlus Picture1, 50, 1, 1, 200, 300, 100
bar_value = 0
putbarval (0)
End Sub

Private Sub Label1_dblClick()
bar_value = 80
End Sub
Private Sub Timer1_Timer()
bar_value = bar_value + 10
putbarval (bar_value)
If bar_value = "110" Then
Unload Me
FrmMain.Show
End If
End Sub
Private Function putbarval(Value As Integer)
Picture1.Width = Value / p
End Function


