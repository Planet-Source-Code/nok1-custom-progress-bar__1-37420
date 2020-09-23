VERSION 5.00
Begin VB.UserControl ShapeProgress 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ScaleHeight     =   390
   ScaleWidth      =   5550
   Windowless      =   -1  'True
   Begin VB.Shape ProgressShape 
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   20
      Top             =   20
      Width           =   5535
   End
   Begin VB.Shape ProgressShape 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   325
      Index           =   1
      Left            =   35
      Top             =   35
      Width           =   5295
   End
End
Attribute VB_Name = "ShapeProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'*                                                                 *
'*   This code was written by Nok1 and is rightful property        *
'*   of Nok1.  Code Copyright 2002 Nok1 Inc.  You may use this     *
'*   code as long the name of the original author is included.     *
'*   And, oh ya, if you put this up on a website or something,     *
'*   please tell me - its not like i can stop you from using it.   *
'*                                                                 *
'*******************************************************************



'Color Constants
Const pbcBlue = &HFF0000
Const pbcRed = &HFF&
Const pbcBlack = &H0&
Const pbcGreen = &HFF00&


'Shape Const
Const pbsRectangle = 0
Const pbsCircle = 3
Const pbsRoundRect = 4

'Local Const Variables
Const Initial = 0
Const Second = 1

'Public Declarations
Private Color2 As Long
Private Shape2 As Integer
Private Maximum2 As Long
Private Minimum2 As Long
Private Value2 As Long

Public Event ProgressChanged(ByVal Progress As Integer, ByVal TotalPercent As Long)
Public Event ShapeChanged(ByVal Shape As Integer)
'public Event
Public Property Get Color() As Long
    Color = Color2
End Property
Public Property Let Value(ByVal ProgressCompleted As Long)
    If ProgressCompleted <= Minimum Then
        ProgressCompleted = Minimum
    End If
    'Comment the next 3 lines to allow overloading
    If ProgressCompleted >= Maximum Then
        ProgressCompleted = Maximum
    End If
    
    f = Percent(ProgressCompleted)
    ProgressShape(Second).Width = f
    Value2 = ProgressCompleted
    RaiseEvent ProgressChanged(ProgressCompleted, f)
End Property

Public Property Get Value() As Long
    Value = Value2
End Property

Public Property Let Color(ByVal Color As Long)
    'Sets color from whatever value was passed.  i didnt add a validation checker, so be carefull
    ProgressShape(Second).BackColor = Color
    Color2 = Color
End Property

Public Property Get Shape() As Integer
    Shape = Shape2
End Property
Public Property Let Shape(val As Integer)

Select Case val
    Case 0 'Rectangle
        ProgressShape(Initial).Shape = pbsRectangle
        ProgressShape(Second).Shape = pbsRectangle
    Case 4 'Rounded Rectangle
        ProgressShape(Initial).Shape = pbsRoundRect
        ProgressShape(Second).Shape = pbsRoundRect
End Select
Shape2 = val
RaiseEvent ShapeChanged(val)

End Property
Public Property Get Maximum() As Integer
    Maximum = Maximum2
End Property
Public Property Let Maximum(ByVal iMax As Integer)
    'Displays that the max is smaller than min & will cause
    'error so it just exits.
    If Max < Min Then
        MsgBox "Max<Min"
        Exit Property
    End If
    Maximum2 = iMax
End Property

Public Property Get Minimum() As Integer
    Minimum = Minimum2
End Property
Public Property Let Minimum(ByVal iMin As Integer)
    'Displays that the min is bigger than the max and will cause
    'error so it just exits
    If Min > Max Then
        MsgBox "Min>Max"
        Exit Property
    End If
    Minimum2 = iMin
End Property

Private Sub UserControl_Initialize()
'set the max to 100 (def) on startup
Me.Maximum = 100
Me.Value = 75
End Sub

Private Sub UserControl_Resize()
If UserControl.Height < 100 Then
    UserControl.Height = 100
End If
If UserControl.Width < 500 Then
    UserControl.Width = 500
End If
Call SizeChange

End Sub

Private Sub SizeChange()
If Me.Shape = 0 Or 4 Then
    ProgressShape(Initial).Height = UserControl.Height - 15
    ProgressShape(Second).Height = ProgressShape(Initial).Height - 50
    ProgressShape(Initial).Width = UserControl.Width - 15
End If
End Sub

Private Function Percent(ByVal Value As Long) As Long
    Dim a As Integer, p As Integer
    a = (Maximum - Minimum)
    p = CInt((Value / a) * ProgressShape(Initial).Width)
    Percent = p
End Function
