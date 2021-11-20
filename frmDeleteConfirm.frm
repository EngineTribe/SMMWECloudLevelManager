VERSION 5.00
Begin VB.Form frmDeleteConfirm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4530
   Icon            =   "frmDeleteConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2440
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image NoButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2400
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image YesButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   240
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "frmDeleteConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const a& = -1
Private Const b& = &H1
Private Const c& = &H2
Private Sub Form_Load()
SetWindowPos Me.hWnd, a, 0, 0, 0, 0, b Or c
frmDeleteConfirm.Caption = ConstStr(4)
frmDeleteConfirm.BackColor = RGB(250, 228, 192)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Label2.Font.Name = "DinkieBitmap 9pxDemo"
Label3.Font.Name = "DinkieBitmap 9pxDemo"
Label1.ForeColor = RGB(89, 15, 16)
Label2.ForeColor = RGB(89, 15, 16)
Label3.ForeColor = RGB(89, 15, 16)
Label1.Caption = ConstStr(5) & frmLevel.LevelName.Caption & ConstStr(6)
Label2.Caption = ConstStr(7)
Label3.Caption = ConstStr(8)
YesButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
NoButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Label2_Click()
YesButton_Click
End Sub

Private Sub Label3_Click()
NoButton_Click
End Sub

Private Sub NoButton_Click()
PlaySFX "snd_close_guardabot.ogg"
Unload Me
End Sub

Private Sub YesButton_Click()
PlaySFX "snd_delete_level.ogg"
Kill LevelFolder & "\" & frmLevel.LevelName.Caption & ".swe"
frmMain.LocalLevelsRefresh
ShowMsgBox "DELETE"
frmMain.SetFocus
Unload frmLevel
Unload Me
End Sub
