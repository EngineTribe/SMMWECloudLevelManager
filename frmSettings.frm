VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   9270
   ClientTop       =   5850
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8940
   Begin VB.CheckBox EnableSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   1440
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      Height          =   495
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmSettings.frx":2AFA
      Left            =   4080
      List            =   "frmSettings.frx":2AFC
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   3480
      Width           =   3255
   End
   Begin VB.CheckBox EnableCDN 
      Caption         =   "CDN"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   960
      Width           =   4575
   End
   Begin VB.CheckBox EnablePreload 
      Caption         =   "Preload"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4215
      Left            =   3840
      TabIndex        =   9
      Top             =   0
      Width           =   4935
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   495
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmSettings.frx":2AFE
         Left            =   240
         List            =   "frmSettings.frx":2B00
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   2415
      End
   End
   Begin VB.CheckBox EnableMusic 
      Caption         =   "Music"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CheckBox EnableSFX 
      Caption         =   "SFX"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
   End
   Begin VB.OptionButton esES 
      Caption         =   "Espanol"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.OptionButton enUS 
      Caption         =   "English"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.OptionButton zhCN 
         Caption         =   "¼òÌåÖÐÎÄ"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Label UpdateLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3500
      Width           =   1695
   End
   Begin VB.Image UpdateButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label NoLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4250
      Width           =   1695
   End
   Begin VB.Label YesLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4250
      Width           =   1695
   End
   Begin VB.Image NoButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2040
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Image YesButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EnableSort_Click()
If EnableSort.Value = 1 Then
MundialesSort = True
EnableSort.Caption = ConstStr(55)
Else
MundialesSort = False
EnableSort.Caption = ConstStr(56)
End If
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
'load GUI
Me.Caption = ConstStr(0) & " " & ConstStr(27)
Me.BackColor = RGB(250, 228, 192)
Frame1.BackColor = RGB(250, 228, 192)
Frame1.Caption = ConstStr(29)
Frame1.ForeColor = RGB(89, 15, 16)
Frame2.BackColor = RGB(250, 228, 192)
Frame2.Caption = ConstStr(14)
Frame2.ForeColor = RGB(89, 15, 16)
Frame1.Font.Name = "DinkieBitmap 9pxDemo"
Frame2.Font.Name = "DinkieBitmap 9pxDemo"
zhCN.Font.Name = "DinkieBitmap 9pxDemo"
enUS.Font.Name = "DinkieBitmap 9pxDemo"
esES.Font.Name = "DinkieBitmap 9pxDemo"
zhCN.ForeColor = RGB(89, 15, 16)
enUS.ForeColor = RGB(89, 15, 16)
esES.ForeColor = RGB(89, 15, 16)
YesLabel.Font.Name = "DinkieBitmap 9pxDemo"
NoLabel.Font.Name = "DinkieBitmap 9pxDemo"
zhCN.BackColor = RGB(250, 228, 192)
enUS.BackColor = RGB(250, 228, 192)
esES.BackColor = RGB(250, 228, 192)
YesLabel.Caption = ConstStr(30)
NoLabel.Caption = ConstStr(8)
YesButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
NoButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
UpdateButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
EnableSFX.Font.Name = "DinkieBitmap 9pxDemo"
EnableMusic.Font.Name = "DinkieBitmap 9pxDemo"
EnablePreload.Font.Name = "DinkieBitmap 9pxDemo"
EnableCDN.Font.Name = "DinkieBitmap 9pxDemo"
EnableSFX.Caption = ConstStr(32)
EnableMusic.Caption = ConstStr(33)
EnablePreload.Caption = ConstStr(34)
EnableCDN.Caption = ConstStr(35)
EnableSFX.BackColor = RGB(250, 228, 192)
EnableMusic.BackColor = RGB(250, 228, 192)
EnablePreload.BackColor = RGB(250, 228, 192)
EnableCDN.BackColor = RGB(250, 228, 192)
EnableSFX.ForeColor = RGB(89, 15, 16)
EnableMusic.ForeColor = RGB(89, 15, 16)
EnablePreload.ForeColor = RGB(89, 15, 16)
EnableCDN.ForeColor = RGB(89, 15, 16)
YesLabel.ForeColor = RGB(89, 15, 16)
NoLabel.ForeColor = RGB(89, 15, 16)
Label1.Caption = ConstStr(45)
Label1.ForeColor = RGB(89, 15, 16)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Combo1.ForeColor = RGB(89, 15, 16)
Combo1.Font.Name = "DinkieBitmap 9pxDemo"
Combo1.AddItem "API"
Combo1.AddItem "WEB"
Label2.Caption = ConstStr(46)
Label2.ForeColor = RGB(89, 15, 16)
Label2.Font.Name = "DinkieBitmap 9pxDemo"
Combo2.ForeColor = RGB(89, 15, 16)
Combo2.Font.Name = "DinkieBitmap 9pxDemo"
UpdateLabel.ForeColor = RGB(89, 15, 16)
UpdateLabel.Font.Name = "DinkieBitmap 9pxDemo"
UpdateLabel.Caption = ConstStr(48)
EnableSort.Font.Name = "DinkieBitmap 9pxDemo"
EnableSort.BackColor = RGB(250, 228, 192)
EnableSort.ForeColor = RGB(89, 15, 16)
If MundialesSort Then
EnableSort.Value = 1
EnableSort.Caption = ConstStr(55)
Else
EnableSort.Value = 0
EnableSort.Caption = ConstStr(56)
End If
'load Locale
If Locale = "zh-cn" Then
zhCN.Value = True
enUS.Value = False
esES.Value = False
ElseIf Locale = "en-us" Then
zhCN.Value = False
enUS.Value = True
esES.Value = False
ElseIf Locale = "es-es" Then
zhCN.Value = False
enUS.Value = False
esES.Value = True
End If
'load settings from global variables
If IsSFXEnable Then EnableSFX.Value = 1
If IsBGMEnable Then EnableMusic.Value = 1
If IsPreloadEnable Then EnablePreload.Value = 1
If ProxyDlSuffix = "?proxied" Then EnableCDN.Value = 1
If DownloadMethod = 1 Then
Combo1.Text = "API"
Else
Combo1.Text = "WEB"
End If
Debug.Print Locale
    MirrorlistTmp = Split(Replace(Join(MirrorList, vbCrLf), "]", ""), "[")
    I = 1
    For I = 1 To UBound(MirrorlistTmp)
    If Locale = "zh-cn" Then
    Combo2.AddItem Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(0)
    ElseIf Locale = "en-us" Then
    Combo2.AddItem Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(1)
    ElseIf Locale = "es-es" Then
    Combo2.AddItem Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(2)
    End If
    Next I
    I = 1
    For I = 1 To UBound(MirrorlistTmp)
    If Split(MirrorlistTmp(I), vbCrLf)(0) = UseMirror Then
    If Locale = "zh-cn" Then
    Combo2.Text = Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(0)
    ElseIf Locale = "en-us" Then
    Combo2.Text = Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(1)
    ElseIf Locale = "es-es" Then
    Combo2.Text = Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(2)
    End If
    Exit For
    End If
    Next I
End Sub
Private Sub UpdateButton_Click()
PlaySFX "snd_aceptar.ogg"
Me.Hide
ShowMsgBox "UPDATING"
Dim GitHubUrl(1 To 2) As String, WebTmp As String
If Locale = "zh-cn" Then
GitHubUrl(1) = "https://hub.fastgit.org/YidaozhanYa/SMMWECloudLevelManager/"
GitHubUrl(2) = "https://gh-rp.sydzy2.workers.dev/https://api.github.com/repos/YidaozhanYa/SMMWECloudLevelManager/releases/latest"
Else
GitHubUrl(1) = "https://github.com/YidaozhanYa/SMMWECloudLevelManager/"
GitHubUrl(2) = "https://api.github.com/repos/YidaozhanYa/SMMWECloudLevelManager/releases/latest"
End If
WebTmp = GetDataSWE(GitHubUrl(2))
frmMsgBox.Hide
Unload frmMsgBox
Sleep 20
If CSng(Replace(Replace(Join(Filter(Split(WebTmp, ","), "tag_name"), ""), Chr(34), ""), "tag_name:v", "")) > CSng(InternalVersion) Then
    If MsgBox(ConstStr(50) & vbCrLf & ConstStr(51) & CStr(AppVersion) & vbCrLf & ConstStr(52) & Replace(Replace(Join(Filter(Split(WebTmp, ","), Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "v"), ""), Chr(34), ""), "name:v", "") & vbCrLf & ConstStr(53), vbOKCancel + vbExclamation, "") = vbCancel Then
    PlaySFX "snd_close_guardabot.ogg"
    frmMain.SetFocus
    Exit Sub
    Else
    PlaySFX "snd_aceptar.ogg"
    Shell "cmd /c start " & Chr(34) & " " & Chr(34) & " " & Chr(34) & GitHubUrl(1) & "releases/tag/v" & Replace(Replace(Join(Filter(Split(WebTmp, ","), "tag_name"), ""), Chr(34), ""), "tag_name:v", "") & Chr(34)
    'Update
    End If
ElseIf CSng(Replace(Replace(Join(Filter(Split(WebTmp, ","), "tag_name"), ""), Chr(34), ""), "tag_name:v", "")) < CSng(InternalVersion) Then
PlaySFX "snd_wrong.ogg"
MsgBox "Error: Newer than GitHub", vbOKOnly, ""
Else
ShowMsgBox "LATEST"
End If
DeleteUrlCacheEntry GitHubUrl(2)
Unload frmSettings
End Sub

Private Sub UpdateLabel_Click()
UpdateButton_Click
End Sub

Private Sub zhCN_Click()
PlaySFX "snd_aceptar.ogg"
zhCN.Value = True
enUS.Value = False
esES.Value = False
End Sub
Private Sub enUS_Click()
PlaySFX "snd_aceptar.ogg"
zhCN.Value = False
enUS.Value = True
esES.Value = False
End Sub
Private Sub esES_Click()
PlaySFX "snd_aceptar.ogg"
zhCN.Value = False
enUS.Value = False
esES.Value = True
End Sub
'events
Private Sub YesLabel_Click()
YesButton_Click
End Sub
Private Sub NoLabel_Click()
NoButton_Click
End Sub
Private Sub NoButton_Click()
PlaySFX "snd_close_guardabot.ogg"
Unload Me
End Sub

Private Sub YesButton_Click()
'Write checkboxes and optionbuttons to main configuration file
PlaySFX "snd_aceptar.ogg"
Kill ConfigFolder & "\MainConfig.txt"
Open ConfigFolder & "\MainConfig.txt" For Output As #2
If zhCN.Value Then LocaleNew = "zh-cn"
If enUS.Value Then LocaleNew = "en-us"
If esES.Value Then LocaleNew = "es-es"
Print #2, LocaleNew
If EnableSFX.Value = 1 Then
IsSFXEnable = True
Print #2, "1"
Else
IsSFXEnable = False
Print #2, "0"
End If
If EnableMusic.Value = 1 Then
IsBGMEnable = True
Print #2, "1"
Else
IsBGMEnable = False
Print #2, "0"
End If
If EnablePreload.Value = 1 Then
IsPreloadEnable = True
Print #2, "1"
Else
IsPreloadEnable = False
Print #2, "0"
End If
If EnableCDN.Value = 1 Then
ProxyDlSuffix = "?proxied"
Print #2, "1"
Else
ProxyDlSuffix = ""
Print #2, "0"
End If
'Write Mirror
    If Locale = "zh-cn" Then
        For I = 1 To UBound(MirrorlistTmp)
        If Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(0) = Combo2.Text Then
        Print #2, Split(MirrorlistTmp(I), vbCrLf)(0)
UseMirror = Split(MirrorlistTmp(I), vbCrLf)(0)
        Exit For
        End If
        Next I
    ElseIf Locale = "en-us" Then
        For I = 1 To UBound(MirrorlistTmp)
        If Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(1) = Combo2.Text Then
        Print #2, Split(MirrorlistTmp(I), vbCrLf)(0)
UseMirror = Split(MirrorlistTmp(I), vbCrLf)(0)
        Exit For
        End If
        Next I
    ElseIf Locale = "es-es" Then
        For I = 1 To UBound(MirrorlistTmp)
        If Split(Replace(Join(Filter(Split(MirrorlistTmp(I), vbCrLf), "Name="), ""), "Name=", ""), ",")(2) = Combo2.Text Then
        Print #2, Split(MirrorlistTmp(I), vbCrLf)(0)
UseMirror = Split(MirrorlistTmp(I), vbCrLf)(0)
        Exit For
        End If
        Next I
    End If
If Combo1.Text = "API" Then
DownloadMethod = 1
Print #2, "1"
Else
DownloadMethod = 0
Print #2, "0"
End If
If EnableSort.Value = True Then
MundialesSort = True
Print #2, "1"
Else
MundialesSort = False
Print #2, "0"
End If
Close #2
Locale = LocaleNew
Unload Me
ShowMsgBox "SAVED"
End Sub
