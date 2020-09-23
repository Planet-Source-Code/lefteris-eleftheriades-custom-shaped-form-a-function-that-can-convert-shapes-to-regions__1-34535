Attribute VB_Name = "Module1"
Public Declare Function LoadCursorFromFile Lib "user32" _
    Alias "LoadCursorFromFileA" _
    (ByVal lpFileName As String) As Long

Public Declare Function SetClassLong Lib "user32" _
    Alias "SetClassLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Const GCL_HCURSOR = (-12)

Public Function SetObjectCursor&(CursorFile$, ObjectHandle&)
Dim hCursor As Long
Dim lReturn As Long
hCursor = LoadCursorFromFile(CursorFile$)
SetObjectCursor = SetClassLong(ObjectHandle&, GCL_HCURSOR, hCursor)
End Function

Public Sub unSetObjectCursor(ObjectHandle&, OriginalImage&)
 SetClassLong ObjectHandle&, GCL_HCURSOR, OriginalImage
End Sub


