VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5280
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "frmAbout.frx":2AFA
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   43.776
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CheckUpdate"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OpenSMMWECloudWebsite"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AboutText"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmAbout.Caption = Form1.AboutButton.Caption
If Locale = "zh-cn" Then
Label1.Caption = Replace(Title, Version, "") & vbCrLf & "V" & Version & vbCrLf & "By ÊÇÒ»µ¶Õ¶ßÕ"
Else
Label1.Caption = Replace(Title, Version, "") & vbCrLf & "V" & Version & vbCrLf & "By YidaozhanYa"
End If
Label2.Caption = ErrorText(28)
Label3.Caption = ErrorText(29)
End Sub

Private Sub Label2_Click()
Shell "cmd /c start https://cloud.smmwe.ml/", vbMinimizedNoFocus
End Sub

Private Sub Label3_Click()
If Locale = "zh-cn" Then
Shell "cmd /c start https://hub.fastgit.org/YidaozhanYa/SMMWECloudLevelManager/releases/latest", vbMinimizedNoFocus
Else
Shell "cmd /c start https://github.com/YidaozhanYa/SMMWECloudLevelManager/releases/latest", vbMinimizedNoFocus
End If
End Sub

Private Sub Picture2_Click()

End Sub
