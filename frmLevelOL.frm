VERSION 5.00
Begin VB.Form frmLevelOL 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Image CopyLinkButton 
      Height          =   855
      Left            =   12360
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label CommentLabel 
      Alignment       =   2  'Center
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
      Left            =   7320
      TabIndex        =   4
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Label DownloadLabel 
      Alignment       =   2  'Center
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
      Left            =   1320
      TabIndex        =   3
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Image CommentButton 
      Height          =   855
      Left            =   6480
      Top             =   5880
      Width           =   5295
   End
   Begin VB.Image DownloadButton 
      Height          =   855
      Left            =   480
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Image lvlImg 
      Height          =   2295
      Left            =   120
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lvlInfos 
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
      Height          =   1335
      Left            =   7320
      TabIndex        =   2
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Label lvlTag 
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
      Height          =   735
      Left            =   7800
      TabIndex        =   1
      Top             =   3240
      Width           =   6135
   End
   Begin VB.Image DecLabel 
      Height          =   615
      Left            =   7320
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image GameStyleImg 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label LevelNameLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A Level"
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
      Height          =   1695
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   11055
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   13440
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmLevelOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommentButton_Click()
PlaySFX "snd_aceptar.ogg"
Shell "cmd /c start " & Chr(34) & " " & Chr(34) & " " & Chr(34) & OLWebIP & Me.LevelNameLabel.Caption & ".swe?preview" & Chr(34)
End Sub

Private Sub CommentLabel_Click()
CommentButton_Click
End Sub

Private Sub CopyLinkButton_Click()
PlaySFX "snd_aceptar.ogg"
Clipboard.SetText Me.LevelNameLabel.Caption & vbCrLf & "https://smmwe-cloud.vercel.app/main/" & Replace(frmLevelOL.LevelNameLabel.Caption, " ", "%20") & ".swe"
ShowMsgBox ("LINK")
End Sub

Private Sub DownloadButton_Click()
PlaySFX "snd_aceptar.ogg"
If IsPreloadEnable Then
    ShowMsgBox "LOADING"
    DoEvents
    If CheckFileExists(LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & ".swe") = False Then
        If DownloadMethod = 1 Then
        Call URLDownloadToFile(0, OLAPIIP & "smmweroot/" & Replace(frmLevelOL.LevelNameLabel.Caption, " ", "%20") & ".swe" & ProxyDlSuffix, LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & ".swe", 0, 0)
        Else
        Call URLDownloadToFile(0, OLWebIP & "main/" & Replace(frmLevelOL.LevelNameLabel.Caption, " ", "%20") & ".swe", LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & ".swe", 0, 0)
        End If
    Else
        If DownloadMethod = 1 Then
        Call URLDownloadToFile(0, OLAPIIP & "smmweroot/" & Replace(frmLevelOL.LevelNameLabel.Caption, " ", "%20") & ".swe" & ProxyDlSuffix, LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & " (2).swe", 0, 0)
        Else
        Call URLDownloadToFile(0, OLWebIP & "main/" & Replace(frmLevelOL.LevelNameLabel.Caption, " ", "%20") & ".swe", LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & " (2).swe", 0, 0)
        End If
    End If
Else
    If CheckFileExists(LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & ".swe") = False Then
        FileCopy ConfigFolder & "\.ParseTemp.tmp", LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & ".swe"
    Else
        FileCopy ConfigFolder & "\.ParseTemp.tmp", LevelFolder & "\" & frmLevelOL.LevelNameLabel.Caption & " (2).swe"
    End If
End If
frmMsgBox.Hide
Unload frmMsgBox
ShowMsgBox "SUCCESS"
frmMain.LocalLevelsRefresh
End Sub

Private Sub DownloadLabel_Click()
DownloadButton_Click
End Sub

Private Sub Form_Load()
If OperateType = 2 And frmMain.ListOL.Text <> "" Then
Me.BackColor = RGB(254, 252, 238)
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-exit-aboutlevel.png")
GameStyleImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\gamestyle-unknown.png")
LevelNameLabel.ForeColor = RGB(89, 15, 16)
LevelNameLabel.Caption = right(frmMain.ListOL.Text, Len(frmMain.ListOL.Text) - 1)
LevelNameLabel.Font.Name = "AsepriteFont"
DecLabel.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-oltag.png")
CopyLinkButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-copylink.png")
'Load Default cfg
lvlTag.ForeColor = RGB(89, 15, 16)
lvlTag.Font.Name = "DinkieBitmap 9pxDemo"
lvlTag.Caption = "?  ?"
lvlInfos.ForeColor = RGB(89, 15, 16)
lvlInfos.Font.Name = "DinkieBitmap 9pxDemo"
lvlInfos.Caption = ConstStr(19)
Dim LocaleSuffixTmp As String
LocaleSuffixTmp = "es-es"
If Locale = "en-us" Then LocaleSuffixTmp = "en-us"
lvlImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\tags-0-" & LocaleSuffixTmp & ".png")
DownloadButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-jugar.png")
CommentButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-jugar.png")
CommentLabel.Caption = ConstStr(20)
DownloadLabel.Caption = ConstStr(21)
CommentLabel.Font.Name = "DinkieBitmap 9pxDemo"
DownloadLabel.Font.Name = "DinkieBitmap 9pxDemo"
DownloadLabel.ForeColor = RGB(89, 15, 16)
CommentLabel.ForeColor = RGB(89, 15, 16)
DoEvents
DoEvents
If IsPreloadEnable Then
lvlInfos.Caption = ConstStr(26)
If CheckFileExists(ConfigFolder & "\.ParseTemp.tmp") Then Kill ConfigFolder & "\.ParseTemp.tmp"
If DownloadMethod = 1 Then
Call URLDownloadToFile(0, OLAPIIP & "smmweroot/" & frmLevelOL.LevelNameLabel.Caption & ".swe", ConfigFolder & "\.ParseTemp.tmp", 0, 0)
Else
Call URLDownloadToFile(0, OLWebIP & "main/" & frmLevelOL.LevelNameLabel.Caption & ".swe", ConfigFolder & "\.ParseTemp.tmp", 0, 0)
End If
Dim levelcontent As String
Open ConfigFolder & "\.ParseTemp.tmp" For Input As #8
Line Input #8, levelcontent
Close #8
levelcontent = Base64Decode(levelcontent)
    levelcontent2 = Split(levelcontent, ",")
    Dim LevelMaker As String
    LevelMaker = Replace(Join(Filter(levelcontent2, Chr(34) & "user" & Chr(34)), ""), Chr(34) & "user" & Chr(34) & ": ", "")
    LevelMaker = Replace(LevelMaker, Chr(34), "")
    If LevelMaker = " 0" Then LevelMaker = ConstStr(3)
    If LevelMaker = " 0.000" Then LevelMaker = ConstStr(3)
    If LevelMaker = " " Then LevelMaker = ConstStr(3)
    If LevelMaker = "" Then LevelMaker = ConstStr(3)
    lvlInfos.Caption = "By " & LevelMaker
    DoEvents
    Dim GameStyle As String
    GameStyle = Replace(Join(Filter(levelcontent2, Chr(34) & "apariencia" & Chr(34)), ""), Chr(34) & "apariencia" & Chr(34) & ": ", "")
    GameStyle = Replace(GameStyle, " } ]", "")
    If GameStyle = " 0" Then
    GameStyle = "SMB1"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\gamestyle-smb1.png")
    ElseIf GameStyle = " 1" Then
    GameStyle = "SMB3"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\gamestyle-smb3.png")
    ElseIf GameStyle = " 2" Then
    GameStyle = "SMW"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\gamestyle-smw.png")
    ElseIf GameStyle = " 3" Then
    GameStyle = "NSMBU"
    GameStyleImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\gamestyle-nsmbu.png")
    End If
    lvlInfos.Caption = lvlInfos.Caption & vbCrLf & GameStyle
    DoEvents
    
    StageStyle = Replace(Join(Filter(levelcontent2, Chr(34) & "entorno" & Chr(34)), ""), Chr(34) & "entorno" & Chr(34) & ": ", "")
    StageStyle = Replace(StageStyle, Chr(34), "")
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
    lvlInfos.Caption = lvlInfos.Caption & " " & StageStyle
    DoEvents
     LevelTimer = Replace(Join(Filter(levelcontent2, Chr(34) & "cronometro" & Chr(34)), ""), Chr(34) & "cronometro" & Chr(34) & ": ", "")
    IsDayNight = Replace(Join(Filter(levelcontent2, Chr(34) & "modo_noche" & Chr(34)), ""), Chr(34) & "modo_noche" & Chr(34) & ": ", "")
    If IsDayNight = " 0" Then
    IsDayNight = GameLabel(27)
    ElseIf IsDayNight = " 1" Then
    IsDayNight = GameLabel(28)
    End If
    lvlInfos.Caption = lvlInfos.Caption & IsDayNight & LevelTimer & "s"
    DoEvents
    GameLabel1 = Replace(Join(Filter(levelcontent2, Chr(34) & "etiqueta1" & Chr(34)), ""), Chr(34) & "etiqueta1" & Chr(34) & ": ", "")
    GameLabel2 = Replace(Join(Filter(levelcontent2, Chr(34) & "etiqueta2" & Chr(34)), ""), Chr(34) & "etiqueta2" & Chr(34) & ": ", "")
    lvlImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\tags-" & CStr(Int((8 - 0 + 1) * Rnd + 0)) & "-" & LocaleSuffixTmp & ".png")
    If CheckFileExists(App.path & "\Assets\tags-" & Replace(CStr(GameLabel1), " ", "") & "-" & LocaleSuffixTmp & ".png") Then lvlImg.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\tags-" & Replace(CStr(GameLabel1), " ", "") & "-" & LocaleSuffixTmp & ".png")
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
    DoEvents
    If Locale = "zh-cn" Then
    lvlTag.Caption = GameLabel1 & " " & GameLabel2
    Else
    lvlTag.Caption = GameLabel1 & ", " & GameLabel2
    End If
End If
End If
End Sub

Private Sub Image1_Click()
PlaySFX "snd_close_guardabot.ogg"
frmMain.SetFocus
Unload Me
End Sub
