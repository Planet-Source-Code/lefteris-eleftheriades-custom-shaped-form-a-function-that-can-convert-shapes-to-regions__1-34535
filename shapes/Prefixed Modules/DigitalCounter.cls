VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DigitalCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DigCounter"
Attribute VB_Ext_KEY = "Member0" ,"DigCounter"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'BitBlt Constants
Private Const SRCAND = &H8800C6 '8913094
Private Const SRCCOPY = &HCC0020 '13369376
Private Const SRCERASE = &H440328 '4457256
Private Const SRCINVERT = &H660046 '6684742
Private Const SRCPAINT = &HEE0086 '15597702
'Applcation Programming Interface Functions
Private Declare Function BitBlt Lib "gdi32" _
            (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
             ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
             ByVal ySrc As Long, ByVal dwRop As Long) As Long
             
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias _
        "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal nDefault As Long, _
        ByVal lpFileName As String) As Long
'Enums
Enum DigitalCounterColor
 GreenDC = 0
 CyanDC = 1
 BlueDC = 2
 MagentaDC = 3
 RedDC = 4
 YellowDC = 5
End Enum
'local Types
Private Type NumberImageType
      Left As Long
      Top As Long
      Width As Long
      Height As Long
End Type
'local variable to know if the sprite data was posted
Private GotSpriteData As Boolean
'Local Variable to store sprite data posted
Private NumberImage(100) As NumberImageType
'local variable to hold collection
Private mCol As Collection

'local variable(s) to hold property value(s)
Private mvarValue As String 'local copy
Private mvarSpriteDataFile As String 'local copy

Public Property Let Value(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


Public Property Get Value() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property




Public Property Let SpriteDataFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SpriteDataFile = 5
    mvarSpriteDataFile = vData
End Property


Public Property Get SpriteDataFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SpriteDataFile
    SpriteDataFile = mvarSpriteDataFile
End Property


Public Sub Draw(ByVal x As Single, ByVal y As Single, _
                Picturebox As Object, MaskBox As Object, _
                Buffer As Object, Optional Color As _
                DigitalCounterColor = GreenDC)
  Dim I%, NumberImageID As Long, FleNme As String
  'If we have NOT taken the sprite data from the file
  If Not GotSpriteData Then
   FleNme = mvarSpriteDataFile
   If Not GetPrivateProfileInt("Sprite1", "Left", 9999, FleNme) = 9999 Then 'if file returns the default value, then that means that Value was not found. Then file was not found
    CaracterFile1Sprites = GetPrivateProfileInt("General", "Sprites", 1, FleNme) 'Get the number of sprites
    'Assign all the data of the file to the array
    For I% = 0 To CaracterFile1Sprites
      NumberImage(I%).Left = GetPrivateProfileInt("Sprite" & I%, "Left", 0, FleNme)
      NumberImage(I%).Top = GetPrivateProfileInt("Sprite" & I%, "Top", 0, FleNme)
      NumberImage(I%).Width = GetPrivateProfileInt("Sprite" & I%, "Width", 0, FleNme)
      NumberImage(I%).Height = GetPrivateProfileInt("Sprite" & I%, "Height", 0, FleNme)
    Next I%
    GotSpriteData = True
    'So we don't have to do it again
   Else
    Err.Raise 75 'File not found
  End If
 End If
 For I% = 1 To Len(Trim(mvarValue))
     NumberImageID = Mid(Trim(mvarValue), I%, 1) + (Color * 10) + 1
     BitBlt Buffer.hdc, x + ((I% - 1) * 9), y, 9, 16, MaskBox.hdc, NumberImage(NumberImageID).Left, NumberImage(NumberImageID).Top, SRCAND
     BitBlt Buffer.hdc, x + ((I% - 1) * 9), y, 9, 16, Picturebox.hdc, NumberImage(NumberImageID).Left, NumberImage(NumberImageID).Top, SRCINVERT
 Next I%
End Sub


Public Function Add(Key As String, Optional sKey As String) As DigCounter
    'create a new object
    Dim objNewMember As DigCounter
    Set objNewMember = New DigCounter


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

Public Property Get Item(vntIndexKey As Variant) As DigCounter
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

