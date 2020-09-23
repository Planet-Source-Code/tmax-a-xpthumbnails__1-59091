VERSION 5.00
Begin VB.UserControl TMaxButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   205
   ToolboxBitmap   =   "TMaxButton.ctx":0000
End
Attribute VB_Name = "TMaxButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by TMax

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

'Constants used by the CombineRgn() API function.
Const RGN_AND = 1&
Const RGN_OR = 2&
Const RGN_XOR = 3&
Const RGN_DIFF = 4&
Const RGN_COPY = 5&

'Constants used by DrawText function
Const DT_CENTER = &H1
Const DT_SINGLELINE = &H20
Const DT_VCENTER = &H4

Dim m_Font As Font
Dim m_ForeColor As OLE_COLOR
Dim m_Caption As String
Dim m_txtRect As RECT

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseLeave()

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Enabled = True
    m_ForeColor = vbBlack
    Set Font = UserControl.Ambient.Font
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawButton 1
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If UserControl.Enabled = False Then Exit Sub
    If x >= 0 And x <= UserControl.ScaleWidth And y >= 0 And y <= UserControl.ScaleHeight Then
        SetCapture UserControl.hWnd
        DrawButton 3
        RaiseEvent MouseMove(Button, Shift, x, y)
    Else
        DrawButton 0
        ReleaseCapture
        RaiseEvent MouseLeave
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawButton 0
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Paint()
    If UserControl.Enabled = True Then
        DrawButton 0
    Else
        DrawButton 2
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    m_Caption = PropBag.ReadProperty("Caption", "Tmax")
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", UserControl.Ambient.Forecolor)
End Sub


Private Sub UserControl_Resize()
Dim hRgn1 As Long
Dim hRgn2 As Long
    hRgn1 = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 3, 3)
    hRgn2 = CreateRoundRectRgn(1, 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 3, 3)
    CombineRgn hRgn1, hRgn2, hRgn1, RGN_XOR
    hRgn2 = CreateRoundRectRgn(3, 3, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3, 3, 3)
    CombineRgn hRgn1, hRgn2, hRgn1, RGN_OR
    SetWindowRgn UserControl.hWnd, hRgn1, True
    DrawCaption False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", m_Caption, "")
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, UserControl.Ambient.Forecolor)
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    m_Caption = NewCaption
    PropertyChanged "Caption"
    Call UserControl_Paint
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    PropertyChanged "Font"
    Call UserControl_Paint
End Property

Public Property Get Forecolor() As OLE_COLOR
    Forecolor = m_ForeColor
End Property

Public Property Let Forecolor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Call UserControl_Paint
End Property

Sub DrawButton(State As Integer)
    Select Case State
        Case 0 'LostFocus
            UserControl.BackColor = &HE0E0E0
        Case 1 'MouseDown
            UserControl.BackColor = &HE8E8E8 '&H888088
        Case 2 'Disabled
            UserControl.BackColor = &HBAC7C9
        Case 3 'MouseOver
            UserControl.BackColor = &HF5F5F5
    End Select
    If State = 1 Then
        DrawCaption True
    Else
        DrawCaption False
    End If
End Sub

Sub DrawCaption(Clicked As Boolean)
    UserControl.Cls
    UserControl.Forecolor = m_ForeColor
    If Clicked Then
        SetRect m_txtRect, 7, 7, UserControl.ScaleWidth - 6, UserControl.ScaleHeight - 6
    Else
        SetRect m_txtRect, 4, 4, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4
    End If
    lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    DrawText UserControl.hdc, m_Caption, -1, m_txtRect, lwFontAlign
End Sub

