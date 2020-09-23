VERSION 5.00
Begin VB.Form FrmPView 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   15330
      TabIndex        =   6
      Top             =   0
      Width           =   15330
      Begin XPThumbs.TMaxButton TMRotateCW 
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Rotate CW"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMZoomOut 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Zoom Out"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMZoomIn 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Zoom In"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMZoomFit 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Zoom Fit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMRotateCCW 
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Rotate CCW"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMRotate180 
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Rotate 180"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMFlipHorizontal 
         Height          =   375
         Left            =   7920
         TabIndex        =   13
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Flip Horizontal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMFlipVertical 
         Height          =   375
         Left            =   9240
         TabIndex        =   14
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Flip Vertical"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMExit 
         Height          =   375
         Left            =   14640
         TabIndex        =   15
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPThumbs.TMaxButton TMSave 
         Height          =   375
         Left            =   12240
         TabIndex        =   16
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Save BMP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XPThumbs.TMaxButton TBFilename 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Tmax"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdNav 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   255
   End
   Begin VB.HScrollBar HscPic 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.VScrollBar VscPic 
      Height          =   2175
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Pictemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   360
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   240
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "FrmPView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Enum Plg
    Rotate_90 = 0
    Rotate_180 = 1
    Rotate_270 = 2
    Flip_Vertical = 3
    Flip_Horizontal = 4
End Enum

Dim Pts(2)  As POINTAPI

' Adjustment properties for HscPic and VscPic
Sub AdjScroll()
CmdNav.Left = -1000
On Error Resume Next
    HscPic.Max = Pic.Width - Me.ScaleWidth + VscPic.Width
    VscPic.Max = Pic.Height - Me.ScaleHeight + HscPic.Height + PicTop.Height
    If HscPic.Max <= 0 Then
        HscPic.Visible = False
    Else
        HscPic.Visible = True
        HscPic_Change
    End If
    If VscPic.Max <= 0 Then
        VscPic.Visible = False
    Else
        VscPic.Visible = True
        VscPic_Change
    End If
    HscPic.Left = 0
    HscPic.Top = Me.ScaleHeight - HscPic.Height
    HscPic.Width = Me.ScaleWidth - VscPic.Width
    HscPic.LargeChange = HscPic.Max \ 10
    VscPic.Top = PicTop.Height
    VscPic.Left = Me.ScaleWidth - VscPic.Width
    VscPic.Height = Me.ScaleHeight - HscPic.Height - PicTop.Height
    VscPic.LargeChange = VscPic.Max \ 10
    If VscPic.Visible = True Or HscPic.Visible = True Then
        CmdNav.Left = VscPic.Left
        CmdNav.Top = HscPic.Top
    End If
    CenterPic
End Sub

'CmdNav - Navigation button - (Left,Top)(Right,Top),(Right,Bottom),(Left,Buttom)
Private Sub CmdNav_Click()
Static Nav%
Select Case Nav%
    Case 0: HscPic.Value = 0: VscPic.Value = 0
    Case 1: HscPic.Value = HscPic.Max: VscPic.Value = 0
    Case 2: HscPic.Value = HscPic.Max: VscPic.Value = VscPic.Max
    Case 3: HscPic.Value = 0: VscPic.Value = VscPic.Max
End Select
Nav% = Nav% + 1
If Nav% > 3 Then Nav% = 0
End Sub

Private Sub Form_Resize()
ZoomFix
End Sub

Private Sub HscPic_Change()
Pic.Left = -HscPic.Value
End Sub

Private Sub HscPic_Scroll()
HscPic_Change
End Sub

'Save current Picture to BMP format
Sub SaveBMP()
On Error GoTo SaveErr
    Dim SaveFile
    Dim ret&
    SaveFile = Mid(TBFilename.Caption, 1, InStr(1, TBFilename.Caption, "."))
    SaveFile = SaveFile & "bmp"
    If Dir(SaveFile) <> "" Then
        ret& = MsgBox("File :" & SaveFile & " already exist." & vbCrLf & "Overwrite ?", vbYesNo, "Save File")
        If ret& = vbYes Then
            'SavePicture Pic.Image, SaveFile
            SavePicture Pictemp.Picture, SaveFile
        End If
    Else
        'SavePicture Pic.Image, SaveFile
        SavePicture Pictemp.Picture, SaveFile
    End If
    Exit Sub
SaveErr:
    MsgBox "Save File " & SaveFile, vbOKOnly, "Save Error"
    Resume Next
End Sub

Private Sub TMExit_Click()
Unload Me
End Sub

Private Sub TMFlipHorizontal_Click()
Blt Flip_Horizontal
End Sub

Private Sub TMFlipVertical_Click()
Blt Flip_Vertical
End Sub

Private Sub TMRotate180_Click()
Blt Rotate_180
End Sub

Private Sub TMRotateCCW_Click()
Blt Rotate_90
End Sub

Private Sub TMRotateCW_Click()
Blt Rotate_270
End Sub

Private Sub TMSave_Click()
SaveBMP
End Sub

Private Sub TMZoomFix_Click()
ZoomFix
End Sub

Private Sub TMZoomIn_Click()
ZoomIn 0.4
End Sub

Private Sub TMZoomOut_Click()
ZoomOut 0.4
End Sub

Private Sub VscPic_Change()
Pic.Top = -VscPic.Value + PicTop.Height
End Sub

Private Sub VscPic_Scroll()
VscPic_Change
End Sub

'PlgBlt Function
'Rotate 90,180,270
'Flip vertical, horizontal
Public Sub Blt(ByVal Deg As Plg)
On Error GoTo BltError
Me.MousePointer = 11
Select Case Deg
Case 0
        Pts(0).y = Pictemp.ScaleWidth
        Pts(0).x = 0
        Pts(1).x = 0
        Pts(1).y = 0
        Pts(2).y = Pictemp.ScaleWidth
        Pts(2).x = Pictemp.ScaleHeight
Case 1
        Pts(0).x = Pictemp.ScaleWidth
        Pts(0).y = Pictemp.ScaleHeight
        Pts(1).x = 0
        Pts(1).y = Pictemp.ScaleHeight
        Pts(2).x = Pictemp.ScaleWidth
        Pts(2).y = 0
Case 2
        Pts(0).x = Pictemp.ScaleHeight
        Pts(0).y = 0
        Pts(1).x = Pictemp.ScaleHeight
        Pts(1).y = Pictemp.ScaleWidth
        Pts(2).x = 0
        Pts(2).y = 0
Case 3
        Pts(0).x = 0
        Pts(0).y = Pictemp.ScaleHeight
        Pts(1).x = Pictemp.ScaleWidth
        Pts(1).y = Pictemp.ScaleHeight
        Pts(2).x = 0
        Pts(2).y = 0
Case 4
        Pts(0).x = Pictemp.ScaleWidth
        Pts(0).y = 0
        Pts(1).x = 0
        Pts(1).y = 0
        Pts(2).x = Pictemp.ScaleWidth
        Pts(2).y = Pictemp.ScaleHeight
End Select
    Pic.Cls
    If Deg = Rotate_90 Or Deg = Rotate_270 Then
        Pic.Width = Pictemp.Height
        Pic.Height = Pictemp.Width
    End If
    PlgBlt Pic.hdc, Pts(0), Pictemp.hdc, 0, 0, Pictemp.ScaleWidth, Pictemp.ScaleHeight, 0, 0, 0
    Me.MousePointer = 0
    Pic.Picture = Pic.Image
    Pictemp.Picture = Pic.Image
    CenterPic
    If VscPic.Visible Or HscPic.Visible Then
        AdjScroll
    End If
    Exit Sub
BltError:
    MsgBox "Error :" & Err.Description, vbOKOnly, "Blt Error"
    Resume Next
End Sub

'Centering Pic, TBFilename
Public Sub CenterPic()
TBFilename.Visible = False
TBFilename.Width = Len(TBFilename.Caption) * 150
TBFilename.Left = (Me.ScaleWidth - TBFilename.Width) / 2
TBFilename.Top = PicTop.Height + 10
TBFilename.Visible = True
Pic.Left = (Me.ScaleWidth - Pic.Width) \ 2
Pic.Top = (Me.ScaleHeight - Pic.Height) \ 2 + PicTop.Height
HscPic.Value = HscPic.Max \ 2
VscPic.Value = VscPic.Max \ 2
End Sub

Sub ZoomOut(Factor As Single)
    Pic.Height = Pic.Height - Pic.Height * Factor
    Pic.Width = Pic.Width - Pic.Width * Factor
    Pic.PaintPicture Pic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
    Pictemp.Height = Pic.Height
    Pictemp.Width = Pic.Width
    Pictemp.PaintPicture Pic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
    AdjScroll
End Sub

Sub ZoomIn(Factor As Single)
    Pic.Height = Pic.Height + Pic.Height * Factor
    Pic.Width = Pic.Width + Pic.Width * Factor
    Pic.PaintPicture Pic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
    Pictemp.Height = Pic.Height
    Pictemp.Width = Pic.Width
    Pictemp.PaintPicture Pic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
    AdjScroll
End Sub

Sub ZoomFix()
    If Pic.Width > Pic.Height Then
        IWidth = Me.ScaleWidth
        IHeight = IWidth * Pic.Height / Pic.Width
    Else
        IHeight = Me.ScaleHeight
        IWidth = IHeight * Pic.Width / Pic.Height
    End If
    If Pic.ScaleWidth > Me.ScaleWidth Then
        IWidth = IWidth
        IHeight = IHeight
    Else
        IWidth = IWidth / 1.1
        IHeight = IHeight / 1.1
    End If
    Pic.Height = IHeight
    Pic.Width = IWidth
    Pic.PaintPicture Pic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
    Pictemp.Height = Pic.Height
    Pictemp.Width = Pic.Width
    Pictemp.PaintPicture Pic, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
    AdjScroll
End Sub

