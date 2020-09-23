Attribute VB_Name = "Mouse"
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As pointapi) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Type pointapi
        X As Long
        Y As Long
End Type

Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Sub moveM(xP As Long, yP As Long)
    Dim nl As Long
    nl = SetCursorPos(xP, yP)
End Sub

Sub LeftClickM(xP As Long, yP As Long)
Dim nl As Long
    nl = SetCursorPos(xP, yP)
    mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
    
End Sub


Sub RightClickM(xP As Long, yP As Long)
Dim nl As Long
    nl = SetCursorPos(xP, yP)
    mouse_event MOUSEEVENTF_RIGHTDOWN, xP, yP, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, xP, yP, 0, 0
End Sub

