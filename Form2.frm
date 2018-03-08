VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Enter"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
      
Public Function PutHbar(ByRef l As ListBox)
    Dim longest As Long
    longest = 0
    If l.ListCount > 0 Then
        Dim i As Integer
        longest = TextWidth(l.List(0) & "  ")
        For i = 1 To l.ListCount - 1
            If TextWidth(l.List(i) & "  ") > longest Then
                longest = TextWidth(l.List(i) & "  ")
            End If
        Next
    End If
    If ScaleMode = vbTwips Then
        longest = longest / Screen.TwipsPerPixelX  ' if twips change to pixels
        SendMessageByNum List1.hwnd, LB_SETHORIZONTALEXTENT, longest, 0
    End If
End Function

      
Private Sub Command1_Click()
    List1.Clear
    PutHbar List1
End Sub

Private Sub Command2_Click()
    Dim s As String
    s = InputBox("Please enter any text", "List scroll", _
           "this is a simple scrollbar sample for demonstration purposes")
    List1.AddItem s
    PutHbar List1
End Sub
