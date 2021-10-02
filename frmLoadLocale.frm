VERSION 5.00
Begin VB.Form frmLoadLocale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SMMWE Cloud Level Manager"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoadLocale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Espanol"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "English"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "简体中文"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image esES 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4200
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image enUS 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2160
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image zhCN 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmLoadLocale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
'Load DomainList
Open App.path & "\Assets\CloudDomain.txt" For Input As #8
Line Input #8, OLWebIP
Line Input #8, OLAPIIP
Close #8
'Load Variable
On Error Resume Next
Dim oShell
Dim strHomeFolder As String
Set oShell = CreateObject("WScript.Shell")
strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
ConfigFolder = strHomeFolder & "\AppData\Local\SMM_WE\LevelManager"
LevelFolder = strHomeFolder & "\AppData\Local\SMM_WE\Niveles"
DesktopFolder = strHomeFolder & "\AppData\Local\SMM_WE\Desktop"
If CheckFileExists("C:\Windows\System32\winecfg.exe") = True Then GoTo ERR
GoTo EndLoadVar
ERR:
Debug.Print "Error! Entering Wine compatible mode."
LevelFolder = "C:\Users\" & Environ("UserName") & "\AppData\Local\SMM_WE\Niveles"
ConfigFolder = "C:\Users\" & Environ("UserName") & "\AppData\Local\SMM_WE\LevelManager"
DesktopFolder = "C:\Users\" & Environ("UserName") & "\Desktop"
EndLoadVar:
'Make Folders
MkDir strHomeFolder & "\AppData\Local\SMM_WE"
MkDir LevelFolder
MkDir ConfigFolder
'Load DinkieBitmap Font
Call AddFontResourceEx(App.path & "\Assets\DinkieBitmap-9pxDemoMod.ttf", FR_PRIVATE, 0)
Call AddFontResourceEx(App.path & "\Assets\AsepriteFont.ttf", FR_PRIVATE, 1)
frmLoadLocale.Font.Name = "DinkieBitmap 9pxDemo"
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Label2.Font.Name = "DinkieBitmap 9pxDemo"
Label3.Font.Name = "DinkieBitmap 9pxDemo"
Label4.Font.Name = "DinkieBitmap 9pxDemo"
'Show Text
Label1.Caption = "请选择你的语言。" & vbCrLf & "Please select your language." & vbCrLf & "Seleccione su idioma."
Label1.ForeColor = RGB(89, 15, 16)
Label2.ForeColor = RGB(89, 15, 16)
Label3.ForeColor = RGB(89, 15, 16)
Label4.ForeColor = RGB(89, 15, 16)
frmLoadLocale.BackColor = RGB(250, 228, 192)
zhCN.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-blank.png")
enUS.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-blank.png")
esES.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\btn-blank.png")
    If CheckFileExists(ConfigFolder & "\MainConfig.txt") = True Then
    Load frmMain
    Unload Me
   frmMain.Show
    End If
End Sub

'Select Locale
Private Sub zhCN_Click()
Locale = "zh-cn"
Open ConfigFolder & "\MainConfig.txt" For Output As #2
Print #2, Locale
DefaultConfig
Close #2
    Load frmMain
    Unload Me
   frmMain.Show
End Sub
Private Sub enUS_Click()
Locale = "en-us"
Open ConfigFolder & "\MainConfig.txt" For Output As #2
Print #2, Locale
DefaultConfig
Close #2
    Load frmMain
    Unload Me
   frmMain.Show
End Sub
Private Sub esES_Click()
Locale = "es-es"
Open ConfigFolder & "\MainConfig.txt" For Output As #2
Print #2, Locale
DefaultConfig
Close #2
    Load frmMain
    Unload Me
   frmMain.Show
End Sub
'Aliases
Private Sub Label2_Click()
zhCN_Click
End Sub
Private Sub Label3_Click()
enUS_Click
End Sub
Private Sub Label4_Click()
esES_Click
End Sub
Private Sub DefaultConfig()
Print #2, "1"
Print #2, "0"
Print #2, "1"
Print #2, "1"
End Sub
