VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17400
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   17400
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.TextBox SearchText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      IMEMode         =   3  'DISABLE
      Left            =   4200
      TabIndex        =   11
      Top             =   4680
      Width           =   9255
   End
   Begin VB.TextBox PageNumTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   2  'OFF
      Left            =   13320
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox ListOL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7230
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   13575
   End
   Begin VB.ListBox ListLocal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7230
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   13575
   End
   Begin VB.Image SettingsButton 
      Height          =   1095
      Left            =   10800
      Top             =   8520
      Width           =   975
   End
   Begin VB.Image AboutButton 
      Height          =   1095
      Left            =   9720
      Top             =   8520
      Width           =   975
   End
   Begin VB.Image OLTag3 
      Height          =   975
      Left            =   15090
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label SearchLabel 
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
      Height          =   855
      Left            =   4200
      TabIndex        =   12
      Top             =   3600
      Width           =   9135
   End
   Begin VB.Label PageNumLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   12960
      TabIndex        =   9
      Top             =   260
      Width           =   1455
   End
   Begin VB.Image PageBtnR 
      Height          =   615
      Left            =   14400
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image PageBtn 
      Height          =   615
      Left            =   13080
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image PageBtnL 
      Height          =   615
      Left            =   11760
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label OLTagLabel2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label OLTagLabel1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Image OLTag2 
      Height          =   855
      Left            =   7560
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Image OLTag1 
      Height          =   855
      Left            =   0
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label OLCastleLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ExploreLevels"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label MundialesLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   4
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Image MundialesButton 
      Height          =   855
      Left            =   9720
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Image ImportButton 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   16320
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label LocalLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11640
      TabIndex        =   3
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Image LocalButton 
      Height          =   855
      Left            =   9720
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Image MainMenuButton 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   16320
      Top             =   120
      Width           =   975
   End
   Begin VB.Label TitleLabelLocal 
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
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.Image TitleLocal 
      Height          =   1095
      Left            =   4800
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label LevelCounterLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Counter"
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
      Left            =   13800
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image OLCastle 
      Height          =   975
      Left            =   240
      Top             =   240
      Width           =   975
   End
   Begin VB.Image MundialesBG 
      Height          =   8415
      Left            =   0
      Top             =   1560
      Width           =   17415
   End
   Begin VB.Image MainMenuBG 
      Height          =   9735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
InternalVersion = "5.1"
AppVersion = "5.1"
'Load Config
Open ConfigFolder & "\MainConfig.txt" For Input As #3
Line Input #3, Locale
Line Input #3, ConfigTmp
If ConfigTmp = 1 Then
IsSFXEnable = True
Else
IsSFXEnable = False
End If
Line Input #3, ConfigTmp
If ConfigTmp = 1 Then
IsBGMEnable = True
Else
IsBGMEnable = False
End If
Line Input #3, ConfigTmp
If ConfigTmp = 1 Then
IsPreloadEnable = True
Else
IsPreloadEnable = False
End If
Line Input #3, ConfigTmp
If ConfigTmp = 1 Then
ProxyDlSuffix = "?proxied"
Else
ProxyDlSuffix = ""
End If
Line Input #3, UseMirror
Line Input #3, ConfigTmp
If ConfigTmp = 1 Then
DownloadMethod = 1
Else
DownloadMethod = 0
End If
Close #3
'Load Mirrorlist
Open App.path & "\Assets\mirrorlist.txt" For Input As #8
    MirrorTmp = ""
    MirrorTmp2 = ""
    Do While Not EOF(8)
    Line Input #8, MirrorTmp2
    If left(MirrorTmp2, 1) <> "'" Then
    MirrorTmp = MirrorTmp & MirrorTmp2 & vbCrLf
    End If
    Loop
    MirrorList = Split(MirrorTmp, vbCrLf)
    ReDim Preserve MirrorList(UBound(MirrorList) + 1)
    I = 0
    For I = 0 To UBound(MirrorList)
    If MirrorList(I) = "[" & UseMirror & "]" Then
    OLWebIP = Replace(MirrorList(I + 2), "Base=", "")
    OLAPIIP = Replace(MirrorList(I + 3), "API=", "")
    Exit For
    End If
    Next I
Close #8

'Load Locale
Open App.path & "\Locale\" & Locale & ".lang" For Input As #1
    LocaleTmp = ""
    LocaleTmp2 = ""
    Do While Not EOF(1)
    Line Input #1, LocaleTmp2
    LocaleTmp = LocaleTmp & LocaleTmp2 & vbCrLf
    Loop
    ConstStr = Split(LocaleTmp, vbCrLf)
    ReDim Preserve ConstStr(UBound(ConstStr) + 1)
Close #1
'Load Locale
Open App.path & "\Locale\label-" & Locale & ".lang" For Input As #6
    LocaleTmp = ""
    LocaleTmp2 = ""
    Do While Not EOF(6)
    Line Input #6, LocaleTmp2
    LocaleTmp = LocaleTmp & LocaleTmp2 & vbCrLf
    Loop
    GameLabel = Split(LocaleTmp, vbCrLf)
    ReDim Preserve ConstStr(UBound(ConstStr) + 1)
    Set LocaleTmp = Nothing
    Set LocaleTmp2 = Nothing
Close #6
'Load DinkieBitmap Font
frmMain.Font.Name = "DinkieBitmap 9pxDemo"
'Show Text
    LocaleSuffix = "es-es"
    If Locale = "en-us" Then LocaleSuffix = "en-us"
    If Locale = "zh-cn" Then LocaleSuffix = "zh-cn"
frmMain.Caption = ConstStr(0) & " " & AppVersion
frmMain.ForeColor = RGB(89, 15, 16)
frmMain.BackColor = RGB(250, 228, 192)
'Init SDLMixelVB
If IsSFXEnable Or IsBGMEnable Then
    If SDL_InitAudio < 0 Then         'Initialize SDL Library first
        MsgBox "SDL Error: " & SDL_GetError, vbOKOnly + vbExclamation
        Exit Sub
    End If
    Mix_Init '0                                   'Init SDL Mixer itself
    Mix_OpenAudio 44100, AUDIO_S16LSB, 2, 2048   'Init Open audio stream
End If
MainMenuButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-mainmenu.png")
ImportButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-import.png")
MainMenuBG.Visible = False
PageNumTxt.Visible = False
SearchText.Visible = False
SearchLabel.Visible = False
LocalLevels
End Sub

Private Sub Form_Terminate()
'Stop SDLMixerX and the process
If IsSFXEnable Or IsBGMEnable Then
    Mix_CloseAudio
    Mix_Quit
    SDL_Quit
    End If
    Call RemoveFontResourceEx(ConfigFolder & "\Assets\DinkieBitmap-9pxDemoMod.ttf", FR_PRIVATE, 0)
    Call RemoveFontResourceEx(App.path & "\Assets\AsepriteFont.ttf", FR_PRIVATE, 1)
    Shell "taskkill /f /im SMMWECloudLevelMgr.exe", vbMinimizedNoFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Stop SDLMixerX and the process
If IsSFXEnable Or IsBGMEnable Then
    Mix_CloseAudio
    Mix_Quit
    SDL_Quit
    End If
    Call RemoveFontResourceEx(ConfigFolder & "\Assets\DinkieBitmap-9pxDemoMod.ttf", FR_PRIVATE, 0)
    Call RemoveFontResourceEx(App.path & "\Assets\AsepriteFont.ttf", FR_PRIVATE, 1)
    Shell "taskkill /f /im SMMWECloudLevelMgr.exe", vbMinimizedNoFocus
End Sub
Public Sub LocalLevels()
'Load Local levels
On Error Resume Next
OperateType = 1
PageBtnL.Visible = False
PageBtnR.Visible = False
PageBtn.Visible = False
PageNumLabel.Visible = False
SearchText.Visible = False
SearchLabel.Visible = False
OLTag1.Visible = False
OLTag2.Visible = False
OLTag3.Visible = False
PageNumLabel.Visible = False
OLTagLabel1.Visible = False
OLTagLabel2.Visible = False
OLCastle.Visible = False
MundialesBG.Visible = False
LevelCounterLocal.Visible = True
ListLocal.Visible = True
OLCastleLabel.Visible = False
TitleLocal.Visible = True
TitleLabelLocal.Visible = True
LocalButton.Visible = False
ListOL.Visible = False
ImportButton.Visible = True
DoEvents
frmMain.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\frmbg-locallevels.png")
frmMain.Caption = ConstStr(0) & " " & AppVersion & " - " & ConstStr(1)
PlayMusic ("snd_guardabot.ogg")
Dim fname As String
fname = Dir(LevelFolder & "\*.swe", 7)
ListLocal.Clear
Do
If fname = "" Then Exit Do
ListLocal.AddItem " " & Replace(fname, ".swe", "")
fname = Dir()
Loop
ReDim locallevel(0 To ListLocal.ListCount - 1) As String
For I = 0 To ListLocal.ListCount - 1
locallevel(I) = ListLocal.List(I)
Next
LevelCounterLocal.ForeColor = RGB(250, 228, 192)
LevelCounterLocal.Font.Name = "AsepriteFont"
LevelCounterLocal.Caption = CStr(ListLocal.ListCount) & "/60"
ListLocal.Font.Name = "AsepriteFont"
ListLocal.ForeColor = RGB(89, 15, 16)
TitleLocal.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-coursebot.png")
TitleLabelLocal.Caption = ConstStr(1)
TitleLabelLocal.ForeColor = RGB(250, 228, 192)
TitleLabelLocal.Font.Name = "DinkieBitmap 9pxDemo"
End Sub
Public Sub LocalLevelsRefresh()
'Refresh
On Error Resume Next
Dim fname As String
fname = Dir(LevelFolder & "\*.swe", 7)
ListLocal.Clear
Do
If fname = "" Then Exit Do
ListLocal.AddItem " " & Replace(fname, ".swe", "")
fname = Dir()
Loop
ReDim locallevel(0 To ListLocal.ListCount - 1) As String
For I = 0 To ListLocal.ListCount - 1
locallevel(I) = ListLocal.List(I)
Next
LevelCounterLocal.Caption = CStr(ListLocal.ListCount) & "/60"
End Sub


Public Sub SMMWECloud()
PageNum = 1
PageNumLabel.Caption = 1
IsLoading = True
'Load online levels
On Error Resume Next
OperateType = 2
ListOL.Font.Name = "DinkieBitmap 9pxDemo"
ListOL.AddItem ""
ListOL.AddItem ConstStr(15)
DoEvents
LevelCounterLocal.Visible = False
OLCastle.Visible = True
MundialesBG.Visible = True
OLCastleLabel.Visible = True
OLTagLabel1.Visible = True
OLTagLabel2.Visible = True
OLTag1.Visible = True
OLTag2.Visible = True
OLTag3.Visible = True
PageNumLabel.Visible = True
ListLocal.Visible = False
TitleLocal.Visible = False
TitleLabelLocal.Visible = False
LocalButton.Visible = False
ImportButton.Visible = False
PageBtnL.Visible = True
PageBtnR.Visible = True
PageBtn.Visible = True
ListOL.Font.Name = "AsepriteFont"
ListOL.ForeColor = RGB(89, 15, 16)
ListOL.BackColor = RGB(254, 252, 238)
OLCastle.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-smmwecloud-castle.png")
MundialesBG.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\frmbg-smmwecloud-fg.png")
frmMain.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\frmbg-smmwecloud.png")
OLTag1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag2.png")
OLTag2.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag1.png")
OLTag3.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tagstar.png")
PageBtnL.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-pagesl.png")
PageBtnR.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-pagesr.png")
PageBtn.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-pages.png")
OLTagLabel1.Font.Name = "DinkieBitmap 9pxDemo"
OLTagLabel2.ForeColor = RGB(0, 135, 134)
OLTagLabel1.ForeColor = RGB(137, 229, 232)
OLTagLabel2.Font.Name = "DinkieBitmap 9pxDemo"
PageNumLabel.Font.Name = "AsepriteFont"
OLTagLabel1.Caption = ConstStr(16)
OLTagLabel2.Caption = ConstStr(17)
frmMain.Caption = ConstStr(0) & " " & AppVersion & " - " & ConstStr(2)
OLCastleLabel.Caption = ConstStr(14)
OLCastleLabel.Font.Name = "AsepriteFont"
DoEvents
PlayMusic ("snd_niveles_mundiales.ogg")
'Get Level list
PageNum = 1
Dim OLLevelList, OLLevelList2() As String
OLLevelList = PostDataSWE(OLWebIP & "main?filename", "pagenum=" & CStr(PageNum))
OLLevelList2 = Split(OLLevelList, vbLf)
ReDim Preserve OLLevelList2(UBound(OLLevelList2))
ListOL.Clear
For I = 0 To UBound(OLLevelList2) - 1
ListOL.AddItem Replace((" " & OLLevelList2(l)), ".swe", "")
l = l + 1
Next
PageNumLabel.Caption = CStr(PageNum)
IsLoading = False
End Sub
Public Sub SMMWECloudRefresh()
On Error GoTo ErrHandler1
IsLoading = True
ListOL.Clear
ListOL.Font.Name = "DinkieBitmap 9pxDemo"
ListOL.AddItem ""
ListOL.AddItem ConstStr(15)
DoEvents
Dim OLLevelList, OLLevelList2() As String
OLLevelList = PostDataSWE(OLWebIP & "main?filename", "pagenum=" & CStr(PageNum))
OLLevelList2 = Split(OLLevelList, vbLf)
ReDim Preserve OLLevelList2(UBound(OLLevelList2))
ListOL.Clear
ListOL.Font.Name = "AsepriteFont"
For I = 0 To UBound(OLLevelList2) - 1
ListOL.AddItem Replace((" " & OLLevelList2(l)), ".swe", "")
l = l + 1
Next
PageNumLabel.Caption = CStr(PageNum)
IsLoading = False
Exit Sub
ErrHandler1:
PageNum = PageNum - 1
SMMWECloudRefresh
Exit Sub
End Sub

Private Sub Image1_Click()

End Sub

Private Sub ListOL_Click()
If IsLoading = False Then
PlaySFX "snd_open_niveles_mundiales.ogg"
frmLevelOL.Show
SetParent frmLevelOL.hWnd, frmMain.hWnd
frmLevelOL.Move 1650, 2000
End If
End Sub

Private Sub OLTag1_Click()
PlaySFX "snd_aceptar.ogg"
OperateType = 2
PageNumTxt.Visible = False
PageNumLabel.Visible = True
PageBtnL.Visible = True
PageBtnR.Visible = True
PageBtn.Visible = True
ListOL.Visible = True
SearchText.Visible = False
SearchLabel.Visible = False
OLTag1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag2.png")
OLTag2.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag1.png")
OLTagLabel2.ForeColor = RGB(0, 135, 134)
OLTagLabel1.ForeColor = RGB(137, 229, 232)
If IsSearching Then
SMMWECloudRefresh
End If
End Sub
Private Sub OLTag2_Click()
OperateType = 3
PlaySFX "snd_aceptar.ogg"
PageNumTxt.Visible = False
PageNumLabel.Visible = False
PageBtnL.Visible = False
PageBtnR.Visible = False
PageBtn.Visible = False
ListOL.Visible = False
SearchText.Visible = True
SearchLabel.Visible = True
SearchLabel.Caption = ConstStr(18)
SearchLabel.ForeColor = RGB(255, 255, 255)
SearchText.ForeColor = RGB(89, 15, 16)
SearchLabel.Font.Name = "DinkieBitmap 9pxDemo"
SearchText.Font.Name = "AsepriteFont"
OLTag1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag1.png")
OLTag2.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag2.png")
OLTagLabel1.ForeColor = RGB(0, 135, 134)
OLTagLabel2.ForeColor = RGB(137, 229, 232)
End Sub



Private Sub ImportButton_Click()
'Import Level
On Error GoTo Exit1
PlaySFX "snd_aceptar.ogg"
Dim filename_select As String
CommonDialog1.DialogTitle = ConstStr(13)
CommonDialog1.InitDir = DesktopFolder
CommonDialog1.Filter = "SMMWE Level|*.swe"
CommonDialog1.ShowOpen
filename_select = CommonDialog1.filename
filename_select = right(filename_select, Len(filename_select) - InStrRev(filename_select, "\"))
FileCopy CommonDialog1.filename, LevelFolder & "\" & filename_select
Exit1:
LocalLevelsRefresh
End Sub

Private Sub ListLocal_Click()
PlaySFX "snd_open_guardabot.ogg"
frmLevel.Show
SetParent frmLevel.hWnd, frmMain.hWnd
frmLevel.Move 1650, 1560
End Sub

Private Sub LocalButton_Click()
PlaySFX "snd_aceptar.ogg"
ListLocal.Visible = True
ImportButton.Visible = True
LocalLevels
MainMenuBG.Visible = False
LocalButton.Visible = False
LocalLabel.Visible = False
AboutButton.Visible = False
SettingsButton.Visible = False
ListOL.Visible = False
MundialesButton.Visible = False
MundialesLabel.Visible = False
End Sub

Private Sub LocalLabel_Click()
LocalButton_Click
End Sub

Private Sub MainMenuBG_Click()
LocalButton.Visible = False
LocalLabel.Visible = False
MundialesButton.Visible = False
MundialesLabel.Visible = False
AboutButton.Visible = False
SettingsButton.Visible = False
PlaySFX "snd_cerrar_menu.ogg"
MainMenuBG.Visible = False
If OperateType = 1 Then
ListLocal.Visible = True
ImportButton.Visible = True
End If
If OperateType = 2 Then
ListOL.Visible = True
End If
If OperateType = 3 Then
SearchText.Visible = True
End If
End Sub

Private Sub MainMenuButton_Click()
'Popup main menu
MainMenuBG.ZOrder
LocalButton.ZOrder
MundialesButton.ZOrder
LocalLabel.ZOrder
MundialesLabel.ZOrder
PageNumTxt = False
If frmLevel.Visible = False And frmLevelOL.Visible = False Then
PlaySFX "snd_abrir_menu.ogg"
ListLocal.Visible = False
ImportButton.Visible = False
ListOL.Visible = False
SearchText.Visible = False
MainMenuBG.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\menu-" & LocaleSuffix & ".png")
MainMenuBG.Visible = True
LocalButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-coursebot.png")
MundialesButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-smmwecloud.png")
LocalButton.Visible = True
LocalLabel.Visible = True
MundialesButton.Visible = True
MundialesLabel.Visible = True
LocalLabel.Font.Name = "DinkieBitmap 9pxDemo"
LocalLabel.Caption = ConstStr(1)
MundialesLabel.Font.Name = "DinkieBitmap 9pxDemo"
MundialesLabel.Caption = ConstStr(2)
LocalLabel.ForeColor = RGB(89, 15, 16)
MundialesLabel.ForeColor = RGB(89, 15, 16)
Unload frmLevel
Unload frmLevelOL
AboutButton.Visible = True
SettingsButton.Visible = True
AboutButton.ZOrder
SettingsButton.ZOrder
AboutButton.ToolTipText = ConstStr(28)
SettingsButton.ToolTipText = ConstStr(27)
AboutButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-about.png")
SettingsButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-settings.png")
End If
End Sub

Private Sub MundialesButton_Click()
'Click SMMWE Cloud Label
PlaySFX "snd_aceptar.ogg"
ListLocal.Visible = False
OLCastle.Visible = True
OLCastleLabel.Visible = True
ImportButton.Visible = False
MainMenuBG.Visible = False
AboutButton.Visible = False
SettingsButton.Visible = False
LocalButton.Visible = False
LocalLabel.Visible = False
ListOL.Visible = True
MundialesButton.Visible = False
MundialesLabel.Visible = False
DoEvents
SMMWECloud
End Sub

Private Sub MundialesLabel_Click()
MundialesButton_Click
End Sub


Private Sub OLTag3_Click()
PlaySFX "snd_wrong.ogg"
End Sub

Private Sub OLTagLabel1_Click()
OLTag1_Click
End Sub

Private Sub OLTagLabel2_Click()
OLTag2_Click
End Sub

Private Sub PageBtnL_Click()
If PageNumTxt.Visible = True Then
    If CInt(PageNumTxt.Text) > 0 Then
    PageNum = CInt(PageNumTxt.Text)
    PageNumTxt.Visible = False
    PlaySFX "snd_aceptar.ogg"
PageNumLabel.Caption = CStr(PageNum)
SMMWECloudRefresh
    End If
    Else
If PageNum <> 1 And PageNum <> 0 Then
PlaySFX "snd_aceptar.ogg"
PageNum = PageNum - 1
PageNumLabel.Caption = CStr(PageNum)
SMMWECloudRefresh
End If
End If
End Sub

Private Sub PageBtnR_Click()
PlaySFX "snd_aceptar.ogg"
If PageNumTxt.Visible = True Then
If CInt(PageNumTxt.Text) > 0 Then
PageNum = CInt(PageNumTxt.Text)
PageNumTxt.Visible = False
PageNumLabel.Caption = CStr(PageNum)
SMMWECloudRefresh
End If
Else
PageNum = PageNum + 1
PageNumLabel.Caption = CStr(PageNum)
SMMWECloudRefresh
End If
End Sub

Private Sub PageNumLabel_Click()
PageNumTxt.Visible = 1
PageNumTxt.Font.Name = "AsepriteFont"
PageNumTxt.Text = CStr(PageNum)
End Sub

Private Sub PageNumTxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
PlaySFX "snd_aceptar.ogg"
If CInt(PageNumTxt.Text) > 0 Then
PageNum = CInt(PageNumTxt.Text)
PageNumLabel.Caption = CStr(PageNum)
PageNumTxt.Visible = False
SMMWECloudRefresh
End If
End If
End Sub
Private Sub SearchText_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler2
If KeyAscii = 13 Then
IsLoading = True
'Get full level list
OperateType = 2
ListOL.Visible = True
IsSearching = True
SearchText.Visible = False
SearchLabel.Visible = False
ListOL.Clear
ListOL.Font.Name = "DinkieBitmap 9pxDemo"
ListOL.AddItem ""
ListOL.AddItem ConstStr(15)
DoEvents
Dim OLLevelList, OLLevelList2() As String
OLLevelList = GetDataSWE(OLAPIIP & "smmweroot")
OLLevelList2 = Split(OLLevelList, vbLf)
OLLevelList2 = Filter(OLLevelList2, SearchText.Text)
ReDim Preserve OLLevelList2(UBound(OLLevelList2))
ListOL.Clear
ListOL.Font.Name = "AsepriteFont"
For I = 0 To UBound(OLLevelList2) - 1
ListOL.AddItem Replace((" " & OLLevelList2(l)), ".swe", "")
l = l + 1
Next
IsLoading = False
End If
Exit Sub
ErrHandler2:
PlaySFX "snd_wrong.ogg"
PageNumTxt.Visible = False
PageNumLabel.Visible = False
PageBtnL.Visible = False
PageBtnR.Visible = False
PageBtn.Visible = False
ListOL.Visible = False
SearchText.Visible = True
SearchLabel.Visible = True
SearchLabel.Caption = ConstStr(18)
SearchLabel.ForeColor = RGB(255, 255, 255)
SearchText.ForeColor = RGB(89, 15, 16)
SearchLabel.Font.Name = "DinkieBitmap 9pxDemo"
SearchText.Font.Name = "AsepriteFont"
OLTag1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag1.png")
OLTag2.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-tag2.png")
OLTagLabel1.ForeColor = RGB(0, 135, 134)
OLTagLabel2.ForeColor = RGB(137, 229, 232)
ShowMsgBox "SEARCHERR"
End Sub

Private Sub SettingsButton_Click()
frmSettings.Show
PlaySFX "snd_aceptar.ogg"
If OperateType = 1 Then
ListLocal.Visible = True
ImportButton.Visible = True
End If
If OperateType = 2 Then
ListOL.Visible = True
End If
If OperateType = 3 Then
SearchText.Visible = True
End If
MainMenuBG.Visible = False
LocalButton.Visible = False
LocalLabel.Visible = False
AboutButton.Visible = False
SettingsButton.Visible = False
MundialesButton.Visible = False
MundialesLabel.Visible = False
End Sub
Private Sub AboutButton_Click()
frmAbout.Show
PlaySFX "snd_aceptar.ogg"
If OperateType = 1 Then
ListLocal.Visible = True
ImportButton.Visible = True
End If
If OperateType = 2 Then
ListOL.Visible = True
End If
If OperateType = 3 Then
SearchText.Visible = True
End If
MainMenuBG.Visible = False
LocalButton.Visible = False
LocalLabel.Visible = False
AboutButton.Visible = False
SettingsButton.Visible = False
MundialesButton.Visible = False
MundialesLabel.Visible = False
End Sub
