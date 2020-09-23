VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin XPThumbs.TMaxButton TMVote 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Vote"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
   End
   Begin XPThumbs.TMaxButton TMOK 
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   480
   End
   Begin XPThumbs.TMaxButton TMXPCalendar 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "XPCalendar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16744576
   End
   Begin VB.Label LblCreator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mailto:tmax_net@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4440
      MouseIcon       =   "frmAbout.frx":208F5
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "mailto:tmax_net@yahoo.com"
      Top             =   5640
      Width           =   2430
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   360
      TabIndex        =   0
      Top             =   3240
      Width           =   3045
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   5520
      TabIndex        =   1
      Top             =   3960
      Width           =   3045
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim start As Boolean
Dim Apppath$
Const Addr1 = "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId="
Const Addr2 = "&lngWId=1"

Private Sub Form_Load()
Dim hRgn1 As Long
    Me.ScaleMode = 3
    hRgn1 = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 12, 12)
    SetWindowRgn Me.hWnd, hRgn1, True
    Me.ScaleMode = 1
    Apppath = App.Path + IIf(Right(App.Path, 1) <> "\", "\", "")
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.FileDescription
    Me.Left = -Me.Width
    Me.Top = (Screen.Height - Me.Height) / 2
    start = True
    Timer1.Enabled = True
    If Dir(Apppath + "Vote.Adr") = "" Then
        CheckAddr
    Else
        OpenAddr
    End If
End Sub

Private Sub LblCreator_Click()
Dim ret&
ret& = ShellExecute(Me.hWnd, "open", LblCreator.Tag, vbNullString, vbNullString, SW_SHOWNORMAL)
LblCreator.Enabled = False
End Sub


Private Sub Timer1_Timer()
If start Then
    If (Me.Left < (Screen.Width - Me.ScaleWidth) / 2) Then
        Me.Left = Me.Left + 1400
    Else
        Timer1.Enabled = False
    End If
Else
    If (Me.Left < Screen.Width) Then
        Me.Left = Me.Left + 700
    Else
        Timer1.Enabled = False
        Unload Me
    End If
End If
End Sub

Sub CheckAddr()
Dim Filename$, txtfile$, Addr$
Dim f1
    Filename = Dir(Apppath + "@PSC*.txt")
    txtfile = Mid(Filename, InStr(1, Filename, "Me_") + 3, 5)
    Addr$ = Addr1 + txtfile + Addr2
    TMVote.Tag = Addr$
    f1 = FreeFile
    Open Apppath + "Vote.Adr" For Output As #f1
        Print #1, Addr$
    Close #1
End Sub

Sub OpenAddr()
Dim f1, ReadAdr
    f1 = FreeFile
    Open Apppath + "Vote.Adr" For Input As #f1
        Line Input #f1, ReadAdr
    Close #f1
    TMVote.Tag = ReadAdr
End Sub
Private Sub TMOK_Click()
 start = False
 Timer1.Enabled = True
End Sub

Private Sub TMVote_Click()
Dim ret&
    ret& = ShellExecute(Me.hWnd, "open", TMVote.Tag, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub TMXPCalendar_Click()
Dim ret&, Addr$
    Addr$ = Addr1$ + "57595" + Addr2$
    ret& = ShellExecute(Me.hWnd, "open", Addr$, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub
