VERSION 5.00
Begin VB.Form frmLoadFont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading Font"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoadFont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmLoadFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
If TESTVAL <> 1 Then
TESTVAL = 1
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
MkDir strHomeFolder
MkDir ConfigFolder
MkDir ConfigFolder & "\Assets"
If CheckFileExists(ConfigFolder & "\Assets\DinkieBitmap-9pxDemoMod.ttf") = False Then

'Download required files
Label1.Caption = "ÕýÔÚÏÂÔØ ¶¡Ã®µãÕóÌå×ÖÌå ..." & vbCrLf & "Downloading DinkieBitmap Font..." & vbCrLf & "Descargando DinkieBitmap Fuente..."
DoEvents
Call URLDownloadToFile(0, "https://3type.cn/fonts/dinkie_bitmap/downloads/DinkieBitmap_Demo_v0.010.zip", ConfigFolder & "\Assets\DinkieBitmap_Demo_v0.010.zip", 0, 0)
Call URLDownloadToFile(0, OLAPIIP & "static/flips.exe?proxied", ConfigFolder & "\Assets\flips.exe", 0, 0)
Call URLDownloadToFile(0, OLAPIIP & "static/7za.exe?proxied", ConfigFolder & "\Assets\7za.exe", 0, 0)
Call URLDownloadToFile(0, OLAPIIP & "static/Patcher.bat?proxied", ConfigFolder & "\Assets\Patcher.bat", 0, 0)
Call URLDownloadToFile(0, OLAPIIP & "static/DinkieBitmap-9pxDemoMod.bps?proxied", ConfigFolder & "\Assets\DinkieBitmap-9pxDemoMod.bps", 0, 0)
'Patch the font
DoEvents
Shell "cmd /c " & Chr(34) & ConfigFolder & "\Assets\Patcher.bat" & Chr(34), vbMinimizedNoFocus
Sleep 5000
FileCopy ConfigFolder & "\Assets\Fonts\DinkieBitmap-9pxDemo.ttf", ConfigFolder & "\Assets\DinkieBitmap-9pxDemoMod.ttf"
Kill ConfigFolder & "\Assets\DinkieBitmap_Demo_v0.010.zip"
Kill ConfigFolder & "\Assets\Patcher.bat"
Kill ConfigFolder & "\Assets\7za.exe"
Kill ConfigFolder & "\Assets\flips.exe"
Shell "cmd /c rd /s /q " & Chr(34) & ConfigFolder & "\Assets\Fonts" & Chr(34)
frmLoadLocale.Show
Unload Me
DoEvents
Else
frmLoadLocale.Show
Unload Me
End If
End If
End Sub
