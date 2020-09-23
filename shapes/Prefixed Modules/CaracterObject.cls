VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CaracterObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Caracters"
Attribute VB_Ext_KEY = "Member0" ,"Caracters"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Reads an .ini file with data and returns the data you want
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias _
        "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal nDefault As Long, _
        ByVal lpFileName As String) As Long
'See PaintFunctions module
Private Declare Function BitBlt Lib "gdi32" _
       (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
       ByVal nWidth As Long, ByVal nHeight As Long, _
       ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
       ByVal dwRop As Long) As Long
'BitBlt Consts
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCINVERT = &H660046
'Store X,Y,Width,Height variables
Private Type CaracterDataPoints
      Left As Long
      Top As Long
      Width As Long
      Height As Long
End Type
'The variable that store X,Y,Width,Height
'values read from a file
Private CaracterFile1Data(100) As CaracterDataPoints
Private CaracterFile1Sprites As Long
''''''''''''''''''''''''''''''''''''''''''''''''''
       
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarFrameLeft As Long 'local copy
Private mvarFrameTop As Long 'local copy
Private mvarFrameWidth As Long 'local copy
Private mvarFrameHeight As Long 'local copy
Private mvarSpriteDataFile As String 'local copy
'My Variables
Private GotFileData As Boolean

Private Sub GetCaracterData(SpriteDataFileName As String)
  Dim I%
  Dim FleNme
  FleNme = SpriteDataFileName
  'if the default value specified is returned
  'then the data wasn't found
  If Not GetPrivateProfileInt("Sprite1", "Left", 9999, FleNme) = 9999 Then
   'get how many sprites are contained in the file
   '[General]
   'Sprites = "29" <---
   CaracterFile1Sprites = GetPrivateProfileInt("General", "Sprites", 1, FleNme)
   For I% = 1 To CaracterFile1Sprites
    'Assign data of the file to the array
    CaracterFile1Data(I%).Left = GetPrivateProfileInt("Sprite" & I%, "Left", 0, FleNme)
    CaracterFile1Data(I%).Top = GetPrivateProfileInt("Sprite" & I%, "Top", 0, FleNme)
    CaracterFile1Data(I%).Width = GetPrivateProfileInt("Sprite" & I%, "Width", 0, FleNme)
    CaracterFile1Data(I%).Height = GetPrivateProfileInt("Sprite" & I%, "Height", 0, FleNme)
   Next I%
  Else
   Err.Raise 75
  End If
End Sub
'Return the FrameCount
Friend Property Get FrameCount() As Long
'used when retrieving value of a property,
'on the right side of an assignment.
    FrameCount = CaracterFile1Sprites
End Property


Public Property Let SpriteDataFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PictureFile = 5
    GetCaracterData vData
    mvarSpriteDataFile = vData
End Property


Public Property Get SpriteDataFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PictureFile
     PictureFile = mvarSpriteDataFile
End Property

'''''''''''''''''''''''''''''''''''''''''''''''
'COMMENT THE LINES BELOW FOR THE PROPERTY LET
'TO MARK PROPERTY AS READ ONLY

'Friend Property Let FrameHeight(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FrameHeight = 5
'    mvarFrameHeight = vData
'End Property
'''''''''''''''''''''''''''''''''''''''''''''''

Friend Property Get FrameHeight(FrameID&) As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameHeight(1)
If FrameID& = 0 Then
 Err.Raise 340, , "Control Array element (0) does not exist." & vbCrLf & "The elements start from 1 not 0"
Else
 If FrameID& <= CaracterFile1Sprites Then
    FrameHeight = CaracterFile1Data(FrameID&).Height  'mvarFrameHeight
    'return the heigt
 Else
  If CaracterFile1Sprites <> 0 Then
   Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "The last existing element is " & CaracterFile1Sprites
  Else
   Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "There is no existing element"
  End If
 End If
End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''
'COMMENT THE LINES BELOW FOR THE PROPERTY LET
'TO MARK PROPERTY AS READ ONLY

'Friend Property Let FrameWidth(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.FrameWidth = 5
'    mvarFrameWidth = vData
'End Property
'''''''''''''''''''''''''''''''''''''''''''''''

Friend Property Get FrameWidth(FrameID&) As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameWidth(1)
If FrameID& <= CaracterFile1Sprites Then
   FrameWidth = CaracterFile1Data(FrameID&).Width 'mvarFrameWidth
   'return the frame width
Else
   If CaracterFile1Sprites <> 0 Then
    Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "The last existing element is " & CaracterFile1Sprites
   Else
    Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "There is no existing element"
   End If
End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''
'COMMENT THE LINES BELOW FOR THE PROPERTY LET
'TO MARK PROPERTY AS READ ONLY

'Friend Property Let FrameTop(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.FrameTop = 5
'    mvarFrameTop = vData
'End Property
'''''''''''''''''''''''''''''''''''''''''''''''

Friend Property Get FrameTop(FrameID&) As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameTop
If FrameID& <= CaracterFile1Sprites Then
   FrameTop = CaracterFile1Data(FrameID&).Top 'mvarFrameTop
Else
   If CaracterFile1Sprites <> 0 Then
    Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "The last existing element is " & CaracterFile1Sprites
   Else
    Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "There is no existing element"
   End If
End If
End Property

'''''''''''''''''''''''''''''''''''''''''''''''
'COMMENT THE LINES BELOW FOR THE PROPERTY LET
'TO MARK PROPERTY AS READ ONLY

'Friend Property Let FrameLeft(ByVal vData As Long)
''used when assigning a value to the property, on the left side of an assignment.
''Syntax: X.FrameLeft = 5
'    mvarFrameLeft = vData
'End Property
'''''''''''''''''''''''''''''''''''''''''''''''

Friend Property Get FrameLeft(FrameID&) As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FrameLeft
If FrameID& <= CaracterFile1Sprites Then
   FrameLeft = CaracterFile1Data(FrameID&).Left  'mvarFrameLeft
Else
   If CaracterFile1Sprites <> 0 Then
    Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "The last existing element is " & CaracterFile1Sprites
   Else
    Err.Raise 340, , "Control Array element (" & FrameID& & ") does not exist." & vbCrLf & "There is no existing element"
   End If
End If
End Property



Public Sub Draw(ByVal FrameID&, ByVal X As Long, _
           ByVal Y As Long, BufferHdc As Variant, _
           PictureHdc As Variant, MaskHdc As Variant)
  'gets the region of the original file to be copied
  'and draws the image (Picture Xor mask)
  On Error Resume Next
  BitBlt BufferHdc, X, Y, CaracterFile1Data(FrameID&).Width, CaracterFile1Data(FrameID&).Height, MaskHdc, CaracterFile1Data(FrameID&).Left, CaracterFile1Data(FrameID&).Top, SRCAND
  BitBlt BufferHdc, X, Y, CaracterFile1Data(FrameID&).Width, CaracterFile1Data(FrameID&).Height, PictureHdc, CaracterFile1Data(FrameID&).Left, CaracterFile1Data(FrameID&).Top, SRCINVERT
End Sub


Public Function Add(Key As String, Optional sKey As String) As Caracters
    'create a new object
    Dim objNewMember As Caracters
    Set objNewMember = New Caracters


    'set the properties passed into the method
    objNewMember.Key = Key
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Caracters
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

'CREATED USING VB's 5.00 or greater
'Add-in called:
'Class builder utility
'To use Class builder utility your self
'go to:
'Add-ins -> Add-In Manager -> VB 6 Class Builder Utility
'[Checked] Loaded/Unloaded
'[Checked] Load on start up
