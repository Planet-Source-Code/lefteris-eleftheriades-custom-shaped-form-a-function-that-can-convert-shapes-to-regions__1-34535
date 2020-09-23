Attribute VB_Name = "Resize"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                        ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
'Move
Public Const HTMOVE = 2
'Resize
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTBOTTOM = 15
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17

Public Const HTMINBUTTON = 8

'The send message function has hadrets of uses
'just like mciSendString, postMessage etc.
'I don't know them all but i do know a couple
'See the Form1.Form_MouseDown() event to see how
'you can use this to move an object

'In an other of my project's relevant to this
'Called sprite editor i use this function for
'resizing a picturebox.
