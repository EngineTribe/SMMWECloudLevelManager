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
      Height          =   3735
      Left            =   3840
      TabIndex        =   9
      Top             =   0
      Width           =   4935
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
   Begin VB.Label NoLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   4200
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
      Left            =   4440
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image NoButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6720
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Image YesButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4440
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
YesButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-blank.png")
NoButton.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-blank.png")
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
If zhCN.Value Then Locale = "zh-cn"
If enUS.Value Then Locale = "en-us"
If esES.Value Then Locale = "es-es"
Print #2, Locale
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
Close #2
Unload Me
ShowMsgBox "SAVED"
End Sub
