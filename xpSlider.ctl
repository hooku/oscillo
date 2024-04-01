VERSION 5.00
Begin VB.UserControl xpSlider 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Image imgSlider 
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      Top             =   480
      Width           =   495
   End
   Begin VB.Image imgFocus 
      Height          =   360
      Left            =   360
      Picture         =   "xpSlider.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgNormal 
      Height          =   360
      Left            =   0
      Picture         =   "xpSlider.ctx":076A
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "xpSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SCROLL_WIDTH As Integer = 11
Private Const SCROLL_HEIGHT As Integer = 21

Private Const SCROLL_SLIDER_WIDTH As Integer = 6

Dim slider_backcolor As Long

Dim slider_val As Integer
Dim slider_min As Integer, slider_max As Integer

Dim slider_pressed As Boolean

Public Event Change()

Public Property Get Min() As Integer
Min = slider_min
End Property

Public Property Let Min(ByVal vNewValue As Integer)
slider_min = vNewValue
End Property

Public Property Get Max() As Integer
Max = slider_max
End Property

Public Property Let Max(ByVal vNewValue As Integer)
slider_max = vNewValue
End Property

Public Property Get Value() As Integer
Value = slider_val
End Property

Public Property Let Value(ByVal vNewValue As Integer)
slider_val = vNewValue

' calc slider position
UserControl.imgSlider.left = (UserControl.ScaleWidth - SCROLL_SLIDER_WIDTH * 2) * slider_val / (slider_max - slider_min) + SCROLL_SLIDER_WIDTH
End Property

Private Sub UserControl_Initialize()
With UserControl.imgSlider
    .left = SCROLL_SLIDER_WIDTH
    .Top = 0
    .Picture = UserControl.imgNormal.Picture
End With
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl.imgSlider.Picture = UserControl.imgFocus.Picture
slider_pressed = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If slider_pressed = True Then
    If x > SCROLL_SLIDER_WIDTH And x < (UserControl.ScaleWidth - SCROLL_SLIDER_WIDTH) Then
        UserControl.imgSlider.left = x - SCROLL_SLIDER_WIDTH
    End If
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' TODO: cannot attain 0
' evaluate the real value
slider_val = (slider_max - slider_min) * (UserControl.imgSlider.left + SCROLL_SLIDER_WIDTH) / (UserControl.ScaleWidth - SCROLL_SLIDER_WIDTH * 2)

UserControl.imgSlider.Picture = UserControl.imgNormal.Picture
slider_pressed = False

RaiseEvent Change
End Sub

Public Property Get BackColor() As Long
BackColor = slider_backcolor
End Property

Public Property Let BackColor(ByVal Color As Long)
slider_backcolor = Color
UserControl.BackColor = slider_backcolor
End Property

Public Property Get ScaleHeight() As Variant
ScaleHeight = UserControl.ScaleHeight
End Property
