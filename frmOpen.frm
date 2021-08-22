VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton OpenButton 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3510
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   3120
      Pattern         =   "*.swe"
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Open file form
'from Super Mario Bros. X 1.3

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub
Private Sub OpenButton_Click()
    OpenLevel
End Sub
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_DblClick()
    OpenLevel
End Sub
Private Sub OpenLevel()
        Dim fname As String, I As Integer
    If File1.ListIndex > 0 Then
            FileCopy File1.Path & "\" & File1.List(File1.ListIndex), LevelFolder & "\" & File1.List(File1.ListIndex)
            Form1.List1.AddItem Replace(File1.List(File1.ListIndex), ".swe", "")
            Form1.LevelCounter.Caption = CStr(Form1.List1.ListCount) & ErrorText(25)
        Unload Me
    ElseIf File1.ListCount > 0 Then
            FileCopy File1.Path & "\" & File1.List(0), LevelFolder & "\" & File1.List(0)
            Form1.List1.AddItem Replace(File1.List(File1.ListIndex), ".swe", "")
            Form1.LevelCounter.Caption = CStr(Form1.List1.ListCount) & ErrorText(25)
        Unload Me
    Else
        MsgBox ErrorText(35), vbOKOnly, "Sorry.", 0, 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Err
    frmOpen.Caption = ErrorText(31)
    OpenButton.Caption = ErrorText(31)
    OpenButton.Caption = ErrorText(34)
    File1.Pattern = "*.swe"
    Dim oShell
    Set oShell = CreateObject("WScript.Shell")
    Dir1.Path = oShell.ExpandEnvironmentStrings("%UserProfile%") & "\Desktop"
    Exit Sub
Err:
End Sub
