VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   180
      ScaleWidth      =   4680
      TabIndex        =   3
      Top             =   3015
      Width           =   4680
   End
   Begin VB.PictureBox Picture3 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   4515
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2670
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   345
      Width           =   165
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2670
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   345
      Width           =   165
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   60
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   105
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   4680
         TabIndex        =   4
         Top             =   0
         Width           =   4680
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  End
End Sub

Private Sub Form_Resize()
 Dim Result
 Result = CreateRoundRectRgn(0, 0, Me.Width / 15, (Me.Height / 15), 120, 120)
 Result = SetWindowRgn(Me.hWnd, Result, True)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Resize.ReleaseCapture
Resize.SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTMOVE, 0
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Resize.ReleaseCapture
  Resize.SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Resize.ReleaseCapture
  Resize.SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Resize.ReleaseCapture
  Resize.SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Resize.ReleaseCapture
  Resize.SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTTOP, 0
End Sub
