VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
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
      Caption         =   "Label1"
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
      Caption         =   "Label1"
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
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Label1.Caption = Title & vbCrLf & "V" & Version & vbCrLf & "By ÊÇÒ»µ¶Õ¶ßÕ"
Label2.Caption = ErrorText(28)
Label3.Caption = ErrorText(29)
End Sub

Private Sub Label2_Click()
Shell "cmd /c start https://cloud.smmwe.ml/", vbMinimizedNoFocus
End Sub

Private Sub Label3_Click()

Shell "cmd /c start https://github.com/YidaozhanYa/SMMWECloudLevelManager/releases/latest", vbMinimizedNoFocus
End Sub
