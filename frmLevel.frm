VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLevel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "About Level"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLevel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.Label UploadLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12600
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image UploadButton 
      Height          =   615
      Left            =   11880
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label ExportLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image ExportButton 
      Height          =   615
      Left            =   7200
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label RenameLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10080
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label DeleteLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12480
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Image RenameButton 
      Height          =   615
      Left            =   9480
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image DeleteButton 
      Height          =   615
      Left            =   11805
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image lvlImg 
      Height          =   2295
      Left            =   430
      Top             =   1850
      Width           =   3495
   End
   Begin VB.Image DayNightImg 
      Height          =   495
      Left            =   13200
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image LvlStyleImg 
      Height          =   615
      Left            =   12000
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Image GameStyleImg 
      Height          =   615
      Left            =   12000
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label StageLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   7200
      TabIndex        =   4
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Label GameLabel2l 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Label GameLabel1l 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7200
      TabIndex        =   2
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label LevelMakerLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   13330
      Top             =   315
      Width           =   615
   End
   Begin VB.Label LevelName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   12135
   End
End
Attribute VB_Name = "frmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DeleteLabel_Click()
DeleteButton_Click
End Sub

Private Sub ExportLabel_Click()
ExportButton_Click
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
On Error Resume Next
frmLevel.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\frmbg-aboutlevel.png")
Image1.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-exit-aboutlevel.png")
LevelName.Font.Name = "AsepriteFont"
LevelName.Caption = Right(frmMain.ListLocal.Text, Len(frmMain.ListLocal.Text) - 1)

    Dim levelcontent As String
    Open LevelFolder & "\" & LevelName.Caption & ".swe" For Input As #5
    Line Input #5, levelcontent
    Close #5
    levelcontent = Base64Decode(levelcontent)
    levelcontent2 = Split(levelcontent, ",")
    Dim LevelMaker As String
    LevelMaker = Replace(Join(Filter(levelcontent2, Chr(34) & "user" & Chr(34)), ""), Chr(34) & "user" & Chr(34) & ": ", "")
    LevelMaker = Replace(LevelMaker, Chr(34), "")
    If LevelMaker = " 0" Then LevelMaker = ConstStr(3)
    If LevelMaker = " 0.000" Then LevelMaker = ConstStr(3)
    If LevelMaker = " " Then LevelMaker = ConstStr(3)
    If LevelMaker = "" Then LevelMaker = ConstStr(3)
    Dim GameStyle As String
    GameStyle = Replace(Join(Filter(levelcontent2, Chr(34) & "apariencia" & Chr(34)), ""), Chr(34) & "apariencia" & Chr(34) & ": ", "")
    GameStyle = Replace(GameStyle, " } ]", "")
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\gamestyle-unknown.png")
    If GameStyle = " 0" Then
    GameStyle = "SMB1"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\gamestyle-smb1.png")
    ElseIf GameStyle = " 1" Then
    GameStyle = "SMB3"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\gamestyle-smb3.png")
    ElseIf GameStyle = " 2" Then
    GameStyle = "SMW"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\gamestyle-smw.png")
    ElseIf GameStyle = " 3" Then
    GameStyle = "NSMBU"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\gamestyle-nsmbu.png")
    End If
    Dim GameLabel1, GameLabel2, LocaleSuffixTmp As String
    LocaleSuffixTmp = "es-es"
    If Locale = "en-us" Then LocaleSuffixTmp = "en-us"
    lvlImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\tags-" & CStr(Int((8 - 0 + 1) * Rnd + 0)) & "-" & LocaleSuffixTmp & ".png")
    GameLabel1 = Replace(Join(Filter(levelcontent2, Chr(34) & "etiqueta1" & Chr(34)), ""), Chr(34) & "etiqueta1" & Chr(34) & ": ", "")
    GameLabel2 = Replace(Join(Filter(levelcontent2, Chr(34) & "etiqueta2" & Chr(34)), ""), Chr(34) & "etiqueta2" & Chr(34) & ": ", "")
    If CheckFileExists(App.Path & "\Assets\tags-" & Replace(CStr(GameLabel1), " ", "") & "-" & LocaleSuffixTmp & ".png") Then lvlImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\tags-" & Replace(CStr(GameLabel1), " ", "") & "-" & LocaleSuffixTmp & ".png")
    If GameLabel1 = " -1" Then
    GameLabel1 = "---"
    Else
    GameLabel1 = GameLabel(GameLabel1)
    End If
    If GameLabel2 = " -1" Then
    GameLabel2 = "---"
    Else
    GameLabel2 = GameLabel(GameLabel2)
    End If
    Dim AutoScroll As String
    AutoScroll = Replace(Join(Filter(levelcontent2, Chr(34) & "autoavance" & Chr(34)), ""), Chr(34) & "autoavance" & Chr(34) & ": ", "")
    If AutoScroll = " 0" Then AutoScroll = "---"
    If AutoScroll = " 1" Then AutoScroll = GameLabel(29)
    If AutoScroll = " 2" Then AutoScroll = GameLabel(30)
    If AutoScroll = " 3" Then AutoScroll = GameLabel(31)
    Dim StageStyle, LevelTimer, IsDayNight As String
    LevelTimer = Replace(Join(Filter(levelcontent2, Chr(34) & "cronometro" & Chr(34)), ""), Chr(34) & "cronometro" & Chr(34) & ": ", "")
    IsDayNight = Replace(Join(Filter(levelcontent2, Chr(34) & "modo_noche" & Chr(34)), ""), Chr(34) & "modo_noche" & Chr(34) & ": ", "")
    If IsDayNight = " 0" Then
    IsDayNight = GameLabel(27)
    DayNightImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\lvl-day.png")
    ElseIf IsDayNight = " 1" Then
    IsDayNight = GameLabel(28)
    DayNightImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\lvl-night.png")
    End If
    StageStyle = Replace(Join(Filter(levelcontent2, Chr(34) & "entorno" & Chr(34)), ""), Chr(34) & "entorno" & Chr(34) & ": ", "")
    StageStyle = Replace(StageStyle, Chr(34), "")
    LvlStyleImg.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\stage-" & Replace(StageStyle, " ", "") & ".png")
    If StageStyle = " ground" Then
    StageStyle = GameLabel(16)
    ElseIf StageStyle = " underground" Then
    StageStyle = GameLabel(17)
    ElseIf StageStyle = " sky" Then
    StageStyle = GameLabel(18)
    ElseIf StageStyle = " forest" Then
    StageStyle = GameLabel(19)
    ElseIf StageStyle = " desert" Then
    StageStyle = GameLabel(20)
    ElseIf StageStyle = " castle" Then
    StageStyle = GameLabel(21)
    ElseIf StageStyle = " ghost" Then
    StageStyle = GameLabel(22)
    ElseIf StageStyle = " airship" Then
    StageStyle = GameLabel(23)
    ElseIf StageStyle = " underwater" Then
    StageStyle = GameLabel(24)
    ElseIf StageStyle = " snow" Then
    StageStyle = GameLabel(25)
    ElseIf StageStyle = " fall" Then
    StageStyle = GameLabel(26)
    ElseIf StageStyle = " beach" Then
    StageStyle = GameLabel(35)
    End If
    Dim LevelCondition As String
    LevelCondition = Replace(Join(Filter(levelcontent2, Chr(34) & "condiciones_count" & Chr(34)), ""), Chr(34) & "condiciones_count" & Chr(34) & ": ", "")
    If Replace(Join(Filter(levelcontent2, Chr(34) & "condiciones" & Chr(34)), ""), Chr(34) & "condiciones" & Chr(34) & ": ", "") = " 0" Then
    LevelCondition = GameLabel(36)
    Else
    If LevelCondition = " 0" Then LevelCondition = GameLabel(32)
    If LevelCondition = " 1" Then LevelCondition = GameLabel(33)
    If LevelCondition = " 3" Then LevelCondition = GameLabel(34)
    End If
   LevelMakerLabel.Font.Name = "DinkieBitmap 9pxDemo"
   GameLabel1l.Font.Name = "DinkieBitmap 9pxDemo"
   GameLabel2l.Font.Name = "DinkieBitmap 9pxDemo"
   StageLabel.Font.Name = "DinkieBitmap 9pxDemo"
LevelMakerLabel.ForeColor = RGB(110, 119, 126)
GameLabel1l.ForeColor = RGB(110, 119, 126)
GameLabel2l.ForeColor = RGB(110, 119, 126)
StageLabel.ForeColor = RGB(110, 119, 126)
    LevelMakerLabel.Caption = LevelMaker
    GameLabel1l.Caption = GameLabel1
    GameLabel2l.Caption = GameLabel2
    StageLabel.Caption = GameStyle & " " & StageStyle & IsDayNight & " " & LevelTimer & " " & AutoScroll & vbCrLf & LevelCondition
    RenameButton.ToolTipText = ConstStr(10)
    DeleteButton.ToolTipText = ConstStr(4)
    DeleteLabel.Caption = ConstStr(4)
DeleteLabel.ForeColor = RGB(29, 42, 67)
   DeleteLabel.Font.Name = "DinkieBitmap 9pxDemo"
    RenameLabel.Caption = ConstStr(10)
RenameLabel.ForeColor = RGB(29, 42, 67)
   RenameLabel.Font.Name = "DinkieBitmap 9pxDemo"
    ExportButton.ToolTipText = ConstStr(11)
    ExportLabel.Caption = ConstStr(11)
ExportLabel.ForeColor = RGB(29, 42, 67)
   ExportLabel.Font.Name = "DinkieBitmap 9pxDemo"
    RenameButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-rename.png")
    DeleteButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-delete.png")
    ExportButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-export.png")
    LvlStyleImg.ToolTipText = StageStyle & " " & IsDayNight
    DayNightImg.ToolTipText = StageStyle & " " & IsDayNight
    GameStyleImg.ToolTipText = GameStyle
    '5.0b3 upload added
    UploadButton.ToolTipText = ConstStr(37)
    UploadLabel.Caption = ConstStr(37)
UploadLabel.ForeColor = RGB(29, 42, 67)
   UploadLabel.Font.Name = "DinkieBitmap 9pxDemo"
    UploadButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-upload.png")
End Sub


Private Sub Image1_Click()
PlaySFX "snd_close_guardabot.ogg"
frmMain.SetFocus
Unload Me
End Sub
Private Sub DeleteButton_Click()
PlaySFX "snd_aceptar.ogg"
frmDeleteConfirm.Show
SetParent frmDeleteConfirm.hWnd, frmMain.hWnd
frmDeleteConfirm.Move 5000, 3000
End Sub

Private Sub ExportButton_Click()
On Error GoTo Exit2
PlaySFX "snd_aceptar.ogg"
Dim filename_select As String
CommonDialog1.FileName = frmLevel.LevelName.Caption & ".swe"
CommonDialog1.DialogTitle = ConstStr(12)
CommonDialog1.InitDir = DesktopFolder
CommonDialog1.Filter = "SMMWE Level|*.swe"
CommonDialog1.ShowSave
filename_select = CommonDialog1.FileName
FileCopy LevelFolder & "\" & frmLevel.LevelName.Caption & ".swe", CommonDialog1.FileName
Exit2:
End Sub

Private Sub RenameButton_Click()
frmRename.Show
SetParent frmRename.hWnd, frmMain.hWnd
frmRename.Move 5000, 3000
PlaySFX "snd_aceptar.ogg"
End Sub

Private Sub RenameLabel_Click()
RenameButton_Click
End Sub

Private Sub UploadButton_Click()
frmUpload.Show
SetParent frmUpload.hWnd, frmMain.hWnd
frmUpload.Move 3000, 2000
PlaySFX "snd_aceptar.ogg"
End Sub
