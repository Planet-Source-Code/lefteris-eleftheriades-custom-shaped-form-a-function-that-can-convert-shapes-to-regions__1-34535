VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00EAEAEA&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5595
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      Height          =   270
      Left            =   8310
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -150
      Top             =   2730
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      Height          =   255
      Index           =   1
      Left            =   8460
      TabIndex        =   4
      Top             =   5010
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5100
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":A1B32
      Height          =   4275
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   4005
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   8580
      X2              =   8280
      Y1              =   4860
      Y2              =   5130
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   360
      X2              =   90
      Y1              =   5190
      Y2              =   5040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":A208C
      Height          =   4275
      Left            =   150
      TabIndex        =   1
      Top             =   570
      Width           =   4035
   End
   Begin VB.Shape Shape4 
      FillStyle       =   7  'Diagonal Cross
      Height          =   225
      Left            =   1080
      Shape           =   2  'Oval
      Top             =   2100
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Shape Shape2 
      Height          =   1395
      Index           =   3
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Shape Shape2 
      Height          =   1365
      Index           =   2
      Left            =   4260
      Shape           =   2  'Oval
      Top             =   -30
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Shape Shape3 
      Height          =   4425
      Left            =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   8745
   End
   Begin VB.Shape Shape2 
      Height          =   1365
      Index           =   1
      Left            =   4290
      Shape           =   2  'Oval
      Top             =   5100
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Shape Shape2 
      Height          =   1395
      Index           =   0
      Left            =   0
      Shape           =   2  'Oval
      Top             =   5100
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Shape Shape1 
      Height          =   675
      Left            =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   8745
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Dim BookRegion&
Enum RgnCombStyle
  RCS_AND = 1 'Shows the part when both regions are touched
  RCS_OR = 2 'Shows the part when one or both regions are touched
  RCS_XOR = 3 'Shows the part when one of both regions are touched
  RCS_diff = 4
  RCS_COPY = 5
End Enum

Enum SScaleMode
 S_Twip = 1
 S_Point = 2
 S_Pixel = 3
 S_Inch = 5
 S_Milimeter = 6
 S_Centimeter = 7
End Enum
Dim BallXDirection As Integer
Dim BallYDirection As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim R1&, R2&, R3&
  BallXDirection = 1
  BallYDirection = 1
  
  'ShapeToRegion, creates Regions and assigns
  'the ID of them to the return variable
  
  'Create the region of the bottom left oval shape
  R1 = ShapeToRegion(Shape2(0), Me)
       'The "ShapeToRegion(Shape2(0))" function does the same as the commented line below
       'CreateEllipticRgn&(Shape2(0).Left / 15, Shape2(0).Top / 15, (Shape2(0).Left / 15) + (Shape2(0).Width / 15), (Shape2(0).Top + Shape2(0).Height) / 15)
  
  'Create the region of the bottom right oval shape
  R2 = ShapeToRegion(Shape2(1), Me)
       'The "ShapeToRegion(Shape2(1))" function does the same as the commented line below
       'CreateEllipticRgn&(Shape2(1).Left / 15, Shape2(1).Top / 15, (Shape2(1).Left / 15) + (Shape2(1).Width / 15), (Shape2(1).Top + Shape2(1).Height) / 15)
       
  'Combine the two regions, delete them and return the
  'Combination to R3
  R3 = CombineRegion(R1&, R2&, RCS_OR, True)
  '————————————————————————————————————'
  '   /|||||||||||\ /|||||||||||\      '
  '  |||||||||||||||||||||||||||||     '
  '   \|||||||||||/ \|||||||||||/      '
  '                                    '
  '————————————————————————————————————'
  
  'Create a shape covering the top half of
  'the two shapes (You will see what for below)
  R1 = ShapeToRegion(Shape1, Me)
       'The "ShapeToRegion(Shape1)" function does the same as the commented line below
       'CreateRectRgn(Shape1.Left / 15, Shape1.Top / 15, (Shape1.Left + Shape1.Width) / 15, (Shape1.Top + Shape1.Height) / 15)
       
  '————————————————————————————————————'
  '  .———————————————————————————.     '
  '  | ___________   ___________ |     '
  '  |/|||||||||||\ /|||||||||||\|     '
  '  '—+++++++++++||++++++++++++—'     '
  '   \|||||||||||/ \|||||||||||/      '
  '                                    '
  '————————————————————————————————————'
  
  'Combine the two regions, do NOT delete them and return the
  'Combination To R2 THE COMBINATION TO USE IS DIFF
  'i.e Cut form an existing region (the box Region (R1))
  'The other two oval regions (R3)
  R2 = CombineRegion(R1&, R3&, RCS_diff, True)
  
  '————————————————————————————————————'
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||.''''''.||||||||.'''''''.|||  '
  '  ||/          \|/\|/           \|  '
  '  '             '  '             '  '
  '————————————————————————————————————'
  'OK! the down part of the form is done
  
  'Now for the body!
  R1 = ShapeToRegion(Shape3, Me)
  R3 = CombineRegion(R1&, R2&, RCS_OR, True)
  'Done with the body
  '————————————————————————————————————'
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||.''''''.||||||||.'''''''.|||  '
  '  ||/          \|/\|/           \|  '
  '  '             '  '             '  '
  '————————————————————————————————————'
  
  'The top right oval
  R1 = ShapeToRegion(Shape2(2), Me)
  R2 = CombineRegion(R1&, R3&, RCS_OR, True)
  
  '————————————————————————————————————'
  '                       ,,,,,,       '
  '                    ,||||||||||.    '
  '   ______________ ,||||||||||||||.  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||.''''''.||||||||.'''''''.|||  '
  '  ||/          \|/\|/           \|  '
  '  '             '  '             '  '
  '————————————————————————————————————'
  
  R1 = ShapeToRegion(Shape2(3), Me)
  BookRegion& = CombineRegion(R1&, R2&, RCS_OR, True)
  '————————————————————————————————————'
  '        ,,,,,          ,,,,,,       '
  '     ,|||||||||.    ,||||||||||.    '
  '   ,|||||||||||||.,||||||||||||||.  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||||||||||||||||||||||||||||||  '
  '  ||||.''''''.||||||||.'''''''.|||  '
  '  ||/          \|/\|/           \|  '
  '  '             '  '             '  '
  '————————————————————————————————————'
  'Book Region finished
  
  'A small oval ball is cut off the book
  R1 = ShapeToRegion(Shape4, Me)
  R3 = CombineRegion(BookRegion&, R1&, RCS_diff, False)
  DeleteObject R1&
  
  SetWindowRgn Me.hwnd, R3, True
  DeleteObject R3&

End Sub

Function CombineRegion(Region1&, Region2&, Style As RgnCombStyle, DeleteSource As Boolean) As Long
'To work it must be on a form, if you want to
'Add this function to a module, See FormShaper
'module, Function ModCombineRegion
Dim ROut&
  ROut& = CreateRectRgn(0, 0, StS(Me.Width, Me.ScaleMode, S_Pixel), StS(Me.Height, Me.ScaleMode, S_Pixel))
  CombineRgn ROut&, Region1&, Region2&, Style
  CombineRegion = ROut&
  If DeleteSource Then
     DeleteObject Region1&
     DeleteObject Region2&
  End If
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTMOVE, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  DeleteObject BookRegion&
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
  Dim R1&, R2&, R3&
  
  'Bounce the small ball shape
  If Shape4.Left + Shape4.Width >= Me.Width Then BallXDirection = -1
  If Shape4.Left <= 0 Then BallXDirection = 1
  If Shape4.Top + Shape4.Height >= Me.Height Then BallYDirection = -1
  If Shape4.Top <= 0 Then BallYDirection = 1
  Shape4.Left = Shape4.Left + (60 * BallXDirection)
  Shape4.Top = Shape4.Top + (60 * BallYDirection)
  'End Bounce the small ball shape
  
  'Make it a region
  R1 = ShapeToRegion(Shape4, Me)
  'Cut it from the book
  R3 = CombineRegion(BookRegion&, R1&, RCS_diff, False)
  'delete the hole
  DeleteObject R1&
  'Apply the region to the form
  SetWindowRgn Me.hwnd, R3, True
  'Delete the book with the hole
  DeleteObject R3&
End Sub
