VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmThumb 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   14730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmThumb.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmThumb.frx":0CCE
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   982
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XPThumbs.TMaxButton TMAbout 
      Height          =   375
      Left            =   13200
      TabIndex        =   14
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "About"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
   End
   Begin XPThumbs.TMaxButton TBSlide 
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Slide Show"
      Top             =   120
      Width           =   1695
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Slide Show"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
   End
   Begin XPThumbs.TMaxButton TBExit 
      Height          =   375
      Left            =   14640
      TabIndex        =   12
      ToolTipText     =   "Exit"
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
      ForeColor       =   8388608
   End
   Begin XPThumbs.TMaxButton TBCreate 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Create Thumbnails"
      Top             =   120
      Width           =   2055
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Create Thumbnail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
   End
   Begin VB.PictureBox PicTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   -15000
      Picture         =   "FrmThumb.frx":45965
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   7
      Top             =   -15000
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox PicImg 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   4920
      Picture         =   "FrmThumb.frx":48E9B
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   6
      Top             =   1275
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Image Img 
         BorderStyle     =   1  'Fixed Single
         Height          =   1800
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1800
      End
   End
   Begin VB.CommandButton CmdReload 
      Caption         =   "LoadBin"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   -1.50000e5
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   11040
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      Pattern         =   "*.jpg;*.bmp;*.gif;*.ico"
      TabIndex        =   1
      Top             =   2460
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ImgThumb 
      Left            =   11280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LVPic 
      Height          =   9015
      Left            =   4560
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   15901
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.Image ImgInvPic 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   6960
      Picture         =   "FrmThumb.frx":4C85E
      Top             =   3360
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image ImgDisp 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   7440
      Picture         =   "FrmThumb.frx":53601
      Top             =   1200
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label LblFileInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   8760
      Width           =   45
   End
   Begin VB.Label LblPath 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      Top             =   11040
      Width           =   720
   End
   Begin VB.Label LblFilename 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1995
      TabIndex        =   8
      Top             =   8280
      Width           =   90
   End
   Begin VB.Image ImgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   360
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'Transparent
      Height          =   4095
      Left            =   210
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Menu MnuFView 
      Caption         =   "fView"
      Visible         =   0   'False
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "FrmThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by TMax (tmax_net@yahoo.com)
'
'XPThumbs - Xpress Thumbnails Viewer (plus Image Viewer / slideshow)
'           Create thumbnails and save to "Imgpb.xpt"
'           when change to directory , it search for "Imgpb.xpt"
'           If found then load "Imgpb.xpt"
'           else auto createthumbnails.
'
'           Use PropertyBag to save and load ImageList and ListView data
'           Use Dictionary to store Filename and CRC
'           Use clsCRC to calculate picture file CRC
'
'           Pros -  Fast thumbnails view
'           Cons -  Increase Disc space, cannot direct write to CD or others protected media
'
'           clsCRC.cls - CRC Checksum Class from Fredrik Qvarfort

Public FilePath As String
Public FileSelect As String
Public Dict As Dictionary
Dim McCrc As clsCRC
Const ThumbSize = 120

Private Sub Dir1_Change()
    FilePath = Dir1.Path + IIf(Right$(Dir1.Path, 1) <> "\", "\", "")
    File1.Path = FilePath
    Dir1.Refresh
    File1.Refresh
    LVPic.ListItems.Clear
    LblOnOff False
    If File1.ListCount > 0 Then StoreCrc
    If Dir$(FilePath + "Imgpb.xpt") <> "" And File1.ListCount > 0 Then
        LoadList
    Else
       'Make it auto create.            -- attach "Imgpb.xpt" to directory
        CreateThumbnails
    End If
End Sub

Private Sub Dir1_Click()
    LblOnOff False
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then Unload Me
    Me.Show
    Set Dict = New Dictionary
    Set McCrc = New clsCRC
    McCrc.Algorithm = CRC32
    SetThumbSize
    Dir1_Change
End Sub

Private Sub LVPic_DblClick()
On Error GoTo LvPicErr
    Dim FPview As New FrmPView
    With FPview
        .Pic = LoadPicture(LVPic.SelectedItem.Key)
        .Pictemp = .Pic.Image
        .AdjScroll
        .CenterPic
        .TBFilename.Caption = LVPic.SelectedItem.Key
        .Show
    End With
    Exit Sub
LvPicErr:
    MsgBox "Load Picture: " & LVPic.SelectedItem.Key & vbCrLf & Err.Description, vbOKOnly, "LvPic LoadPicture"
    
End Sub

Private Sub LVPic_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo LoadPicError
    ImgPreview.Visible = False
    ImgPreview.Picture = LoadPicture(Item.Key)
    LblFilename.Caption = Item.Text
    LblPath.Caption = Item.Key
    If ImgPreview.Picture.Width > ImgPreview.Picture.Height Then
        ImgPreview.Width = 240
        ImgPreview.Height = ImgPreview.Picture.Height / ImgPreview.Picture.Width * ImgPreview.Width
    Else
        ImgPreview.Height = 240
        ImgPreview.Width = ImgPreview.Picture.Width / ImgPreview.Picture.Height * ImgPreview.Height
    End If
    ImgPreview.Left = (Shape1.Width - ImgPreview.Width) \ 2 + Shape1.Left
    ImgPreview.Top = (Shape1.Height - ImgPreview.Height) \ 2 + Shape1.Top
    LblFileInfo.Caption = BmpInfo(ImgPreview) & "  CRC=" & ImgThumb.ListImages(Item.Index).Tag
    LblOnOff True
    ImgPreview.Visible = True
    Exit Sub
LoadPicError:
    LblOnOff False
    ImgPreview.Visible = False
    ImgPreview.Stretch = False
    ImgPreview.Picture = ImgInvPic.Picture
    ImgPreview.Left = (Shape1.Width - ImgPreview.Width) \ 2 + Shape1.Left
    ImgPreview.Top = (Shape1.Height - ImgPreview.Height) \ 2 + Shape1.Top
    ImgPreview.Stretch = True
    ImgPreview.Visible = True
'Resume Next
End Sub

Sub LblOnOff(OnOff As Boolean)
    LblFilename.Visible = OnOff
    LblPath.Visible = OnOff
    LblFileInfo.Visible = OnOff
    ImgPreview.Visible = OnOff
End Sub

Sub LoadList()
On Error Resume Next
    ImgPreview.Top = -10000
    LVPic.ListItems.Clear
    LVPic.Icons = Nothing
    ImgThumb.ListImages.Clear
    LoadImageList FilePath + "Imgpb.xpt"
End Sub

Sub CreateThumbnails()
On Error Resume Next
    LblOnOff False
    Dim CrcStr As String
    ImgPreview.Top = -10000
    PB1.Max = File1.ListCount
    PB1.Visible = True
    LVPic.Icons = Nothing
    ImgThumb.ListImages.Clear
    LVPic.ListItems.Clear
    PicImg.Visible = True
    PicImg.Picture = ImgDisp.Picture
    For i% = 0 To File1.ListCount - 1
        PicImg.Cls
        Img.Picture = LoadPicture(FilePath & File1.List(i%))
        If Img.Picture.Width > Img.Picture.Height Then
            Img.Width = ThumbSize - 34
            Img.Height = Img.Picture.Height * (ThumbSize - 34) / Img.Picture.Width
        Else
            Img.Height = ThumbSize - 34
            Img.Width = Img.Picture.Width * (ThumbSize - 34) / Img.Picture.Height
        End If
        Img.Top = (PicImg.Height - Img.Height) \ 2
        Img.Left = (PicImg.Width - Img.Width) \ 2
        PicImg.Picture = PicImg.Image
        BitBlt Pictemp.hdc, 0, 0, ThumbSize, ThumbSize, PicImg.hdc, 0, 0, vbSrcCopy
        Pictemp.Picture = Pictemp.Image
        ImgThumb.ListImages.Add i% + 1, FilePath & File1.List(i%), Pictemp.Picture
        ImgThumb.ListImages(i% + 1).Tag = Dict.Item(FilePath & File1.List(i%))
        LVPic.Icons = ImgThumb
        LVPic.ListItems.Add i% + 1, FilePath & File1.List(i%), File1.List(i%), i% + 1
        PB1.Value = i% + 1
    Next i%
    SaveImageList FilePath & "Imgpb.xpt"
    PB1.Visible = False
    PicImg.Visible = False
    ImgPreview.Visible = True
    Set PicImg = Nothing
    Set Pictemp = Nothing
End Sub

Private Sub LVPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MnuFView
End Sub

Private Sub MnuDelete_Click()
Dim ret&
    ret& = MsgBox("Delete File :" & LVPic.SelectedItem.Key & vbCrLf & "Are You Sure?", vbYesNo, "Delete File")
    If ret& = vbYes Then
        Kill LVPic.SelectedItem.Key
        ImgPreview.Visible = False
        File1.Refresh
        If File1.ListCount = 0 Then
            LVPic.ListItems.Clear
            Kill FilePath & "Imgpb.xpt"
            Exit Sub
        End If
        LVPic.Icons = Nothing
        ImgThumb.ListImages.Remove (LVPic.SelectedItem.Index)
        LVPic.ListItems.Clear
        LVPic.Icons = ImgThumb
        For i% = 0 To File1.ListCount - 1
            LVPic.ListItems.Add i% + 1, ImgThumb.ListImages(i% + 1).Key, File1.List(i%), i% + 1
        Next i%
        StoreCrc
        SaveImageList FilePath & "Imgpb.xpt"
    End If
End Sub

Private Sub TBCreate_Click()
CreateThumbnails
End Sub

Private Sub TBExit_Click()
Unload Me
End Sub

Private Sub TBSlide_Click()
FrmSlide.Show 1
End Sub

Sub SetThumbSize()
    PicImg.Width = ThumbSize
    PicImg.Height = ThumbSize
    Pictemp.Width = ThumbSize
    Pictemp.Height = ThumbSize
End Sub

' save the property bag to a file
Public Sub SaveImageList(ByVal Filename As String)
On Error GoTo SaveImgErr
    Dim pb As New PropertyBag
    Dim varTemp As Variant
    Dim handle As Long
    pb.WriteProperty "CRC", ImgThumb.Tag
    pb.WriteProperty "ImageList", ImgThumb.object
    pb.WriteProperty "ListView", LVPic.object
    varTemp = pb.Contents
    If Len(Dir$(Filename)) Then Kill Filename
    handle = FreeFile
    Open Filename For Binary As #handle
    Put #handle, , varTemp
    Close #handle
    Set pb = Nothing
    Exit Sub
SaveImgErr:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Save Image Error"
    Resume Next
End Sub

' Load file and read its contents
Public Sub LoadImageList(ByVal Filename As String)
On Error GoTo LoadImgErr
    Dim pb As New PropertyBag
    Dim varTemp As Variant
    Dim handle As Long
    Dim LImg As ListImage
    Dim ImgLocal As Object
    Dim lItem As ListItem
    Dim LvLocal As Object
    If Len(Dir$(Filename)) = 0 Then Err.Raise 53
    handle = FreeFile
    Open Filename For Binary As #handle
    Get #handle, , varTemp
    Close #handle
    ' rebuild the property bag object
    pb.Contents = varTemp
    Set ImgLocal = pb.ReadProperty("ImageList")
    Set LvLocal = pb.ReadProperty("ListView")
    If ImgLocal.ListImages.Count <> File1.ListCount Or pb.ReadProperty("CRC") <> ImgThumb.Tag Then
       MsgBox "CRC unmatch"
      ' CreateThumbnails
      '  Exit Sub
    End If
    For Each LImg In ImgLocal.ListImages
        ImgThumb.ListImages.Add LImg.Index, LImg.Key, LImg.Picture
        ImgThumb.ListImages(LImg.Index).Tag = LImg.Tag
        If Dir(LImg.Key) = "" Then
            MsgBox "File Missing"
           CreateThumbnails
          Exit Sub
        Else
            If LImg.Tag <> Dict.Item(LImg.Key) Then
                MsgBox "Crc Error" & LImg.Index & "  ===  " & Hex(McCrc.CalculateFile(LImg.Key)) & vbCrLf & Dict.Item(LImg.Key)
                CreateThumbnails  'will replace with  RearrangeThumbNails (LImg.Tag)
                ''Create ONLY CRC diff thumbnail
             Exit Sub
            End If
        End If
    Next
    Set LVPic.Icons = ImgThumb
    For Each lItem In LvLocal.ListItems
         LVPic.ListItems.Add lItem.Index, lItem.Key, lItem.Text, lItem.Index
    Next
    Set LvLocal = Nothing
    Set ImgLocal = Nothing
    Set pb = Nothing
    Exit Sub
LoadImgErr:
    MsgBox "Error #" & Err.Number & vbCrLf & Err.Description, vbOKOnly, "Load Image Error"
    Resume Next
End Sub

Function BmpInfo(ByRef TheStdPicture As StdPicture) As String
    Dim tBMP As BITMAP
    GetObjectAPI TheStdPicture.handle, Len(tBMP), tBMP
    BmpInfo = Str$(tBMP.bmWidth) & " x " & Str(tBMP.bmHeight) & " x " & Str(tBMP.bmBitsPixel) & "b"
End Function

Sub StoreCrc()
Dim i%
Dim a, b
Dim CrcVal As String
On Error Resume Next
PB1.Max = File1.ListCount
PB1.Visible = True
Set Dict = Nothing
Set Dict = New Dictionary
For i% = 0 To File1.ListCount - 1
    'If Not Dict.Exists(Dict.Keys(I%)) Then
        Dict.Add FilePath & File1.List(i%), Hex(McCrc.CalculateFile(FilePath & File1.List(i%)))
    'Else
        'If Dict.Keys(i%) <> Hex(McCrc.CalculateFile(FilePath & File1.List(i%))) Then MsgBox "Unmatch"
    'End If
    CrcVal = CrcVal + Dict.Item(FilePath & File1.List(i%))
    PB1.Value = i% + 1
    DoEvents
Next i%
McCrc.Clear
ImgThumb.Tag = Hex(McCrc.CalculateString(CrcVal))
PB1.Visible = False
End Sub

Private Sub TMAbout_Click()
frmAbout.Show 1
End Sub
