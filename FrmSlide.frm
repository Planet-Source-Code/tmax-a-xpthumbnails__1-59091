VERSION 5.00
Begin VB.Form FrmSlide 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   480
   End
   Begin XPThumbs.TMaxButton TBSlide 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "|<"
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
   Begin XPThumbs.TMaxButton TBSlide 
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "<"
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
   Begin XPThumbs.TMaxButton TBSlide 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   ">"
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
   Begin XPThumbs.TMaxButton TBSlide 
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   ">|"
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
   Begin XPThumbs.TMaxButton TBSlide 
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   615
      _ExtentX        =   1085
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
   Begin XPThumbs.TMaxButton TBFilename 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
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
End
Attribute VB_Name = "FrmSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CurrentSlide%

Private Sub Form_DblClick()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Load()
CurrentSlide = -1
Timer1.Enabled = True
Timer1.Interval = 3000
End Sub

Sub MoveSlideButton()
Dim i%
For i% = 0 To 4
    TBSlide(i%).Top = Me.ScaleHeight - TBSlide(i%).Height
Next i%
TBSlide(4).Left = Me.ScaleWidth - TBSlide(4).Width - 10
For i% = 3 To 0 Step -1
    TBSlide(i%).Left = TBSlide(i% + 1).Left - TBSlide(i%).Width
Next i%
End Sub

Private Sub Form_Resize()
MoveSlideButton
End Sub

Private Sub TBSlide_Click(Index As Integer)
On Error Resume Next
Dim IWidth As Long
Dim IHeight As Long
Dim Picshow As StdPicture
Dim i%
For i% = 0 To 4
    TBSlide(i%).Enabled = True
Next i%

Select Case Index
    Case 0
        CurrentSlide = 0
    Case 1
        If CurrentSlide > 0 Then CurrentSlide = CurrentSlide - 1
    Case 2
        If CurrentSlide < FrmThumb.File1.ListCount - 1 Then CurrentSlide = CurrentSlide + 1
    Case 3
        CurrentSlide = FrmThumb.File1.ListCount - 1
    Case 4
        Unload Me
End Select
    If CurrentSlide = 0 Then TBSlide(0).Enabled = False: TBSlide(1).Enabled = False
    If CurrentSlide = FrmThumb.File1.ListCount - 1 Then TBSlide(2).Enabled = False: TBSlide(3).Enabled = False
    Set Picshow = LoadPicture(FrmThumb.FilePath & FrmThumb.File1.List(CurrentSlide))
    TBFilename.Caption = FrmThumb.FilePath & FrmThumb.File1.List(CurrentSlide)
    TBFilename.Width = Len(TBFilename.Caption) * 150
    TBFilename.Left = (Me.ScaleWidth - TBFilename.Width) / 2
    TBFilename.Visible = True
        If Picshow.Width > Picshow.Height Then
            IWidth = Me.ScaleWidth '   Picshow.Width
            IHeight = IWidth * Picshow.Height / Picshow.Width
        Else
            IHeight = Me.ScaleHeight '  Picshow.Height
            IWidth = IHeight * Picshow.Width / Picshow.Height
        End If
    If Picshow.Width > Me.ScaleWidth Then
        IWidth = IWidth \ 1.1
        IHeight = IHeight \ 1.1
    End If
    Me.Cls
    Me.PaintPicture Picshow, (Me.ScaleWidth - IWidth) / 2, (Me.ScaleHeight - IHeight) / 2, IWidth, IHeight
    Set Picshow = Nothing
End Sub

Private Sub Timer1_Timer()
TBSlide_Click 2
If CurrentSlide = FrmThumb.File1.ListCount - 1 Then CurrentSlide = -1
End Sub


