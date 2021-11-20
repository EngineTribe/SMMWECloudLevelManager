VERSION 5.00
Begin VB.Form frmRename 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2440
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Image NoButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2400
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image YesButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c
frmRename.BackColor = RGB(250, 228, 192)
frmRename.Caption = ConstStr(10)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Label2.Font.Name = "DinkieBitmap 9pxDemo"
Label3.Font.Name = "DinkieBitmap 9pxDemo"
Label2.Caption = ConstStr(7)
Label3.Caption = ConstStr(8)
Label1.Caption = ConstStr(9)
Label1.ForeColor = RGB(89, 15, 16)
Label2.ForeColor = RGB(89, 15, 16)
Label3.ForeColor = RGB(89, 15, 16)
YesButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
NoButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")

Text1.ForeColor = RGB(89, 15, 16)
Text1.Font.Name = "DinkieBitmap 9pxDemo"
Text1.Text = frmLevel.LevelName.Caption
End Sub

Private Sub NoButton_Click()
PlaySFX "snd_close_guardabot.ogg"
Unload Me
End Sub

Private Sub YesButton_Click()
PlaySFX "snd_aceptar.ogg"
Name LevelFolder & "\" & frmLevel.LevelName.Caption & ".swe" As LevelFolder & "\" & Text1.Text & ".swe"
frmMain.LocalLevelsRefresh
ShowMsgBox "RENAME"
frmMain.SetFocus
Unload frmLevel
Unload Me
End Sub

Private Sub Label2_Click()
YesButton_Click
End Sub

Private Sub Label3_Click()
NoButton_Click
End Sub
