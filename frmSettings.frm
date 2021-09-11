VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.OptionButton Language3 
      Caption         =   "Spanish"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton Language2 
      Caption         =   "English"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton Language1 
         Caption         =   "S.Chinese"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox EnableCaching 
      Caption         =   "EnableCaching"
      Height          =   300
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmSettings.Caption = Title & " - " & ErrorText(52)
EnableCaching.Caption = ErrorText(53)
SaveButton.Caption = ErrorText(54)
CancelButton.Caption = ErrorText(34)
Frame1.Caption = ErrorText(55)
Language1.Caption = ErrorText(56)
Language2.Caption = ErrorText(57)
Language3.Caption = ErrorText(58)

EnableCaching.Value = IsEnableCache
If Locale = "zh-cn" Then
Language1.Value = True
ElseIf Locale = "en-us" Then
Language2.Value = True
ElseIf Locale = "es-es" Then
Language3.Value = True
End If
End Sub
Private Sub Language1_Click()
Language2.Value = False
Language3.Value = False
End Sub
Private Sub Language2_Click()
Language1.Value = False
Language3.Value = False
End Sub
Private Sub Language3_Click()
Language2.Value = False
Language1.Value = False
End Sub

Private Sub SaveButton_Click()
If Language1.Value = True Then Locale = "zh-cn"
If Language2.Value = True Then Locale = "en-us"
If Language3.Value = True Then Locale = "es-es"
IsEnableCache = CInt(EnableCaching.Value)
Kill ConfigFolder & "\SMMWECloudLevelManager.cfg"
Open ConfigFolder & "\SMMWECloudLevelManager.cfg" For Output As #2
Print #2, Locale
Print #2, CStr(IsEnableCache)
Close #2
Unload Me
End Sub
Private Sub CancelButton_Click()
Unload Me
End Sub
