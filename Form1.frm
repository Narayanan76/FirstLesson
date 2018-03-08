VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5280
      List            =   "Form1.frx":000A
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   4800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
   Begin VB.DirListBox Dir1 
      Height          =   5040
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   5280
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
    File1.Pattern = Combo1.Text
End Sub

Private Sub Combo1_Click()
    File1.Pattern = Combo1.Text
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo here
    Dim d As String
    d = Dir1.Path
    Dir1.Path = Drive1.Drive
    Exit Sub
here:
    MsgBox "Disk failed!"
    Drive1.Drive = d
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 1
End Sub
