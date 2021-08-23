VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "SMMWE Cloud Tool Config"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Œ¢»Ì—≈∫⁄"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2325
   ScaleWidth      =   5595
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command3 
      Caption         =   "Espanol"
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "English"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ºÚÃÂ÷–Œƒ"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Err2
Open ConfigFolder & "\SMMWECloudLocale.cfg" For Output As #2
Print #2, "zh-cn"
Close #2
Unload Me
    Load Form1
    Form1.Show
    Exit Sub
Err2:
MkDir ConfigFolder
MkDir ConfigFolder & "\Niveles"
Open ConfigFolder & "\SMMWECloudLocale.cfg" For Output As #2
Print #2, "zh-cn"
Close #2
Unload Me
    Load Form1
    Form1.Show
End Sub
Private Sub Command2_Click()
On Error GoTo Err3
Open ConfigFolder & "\SMMWECloudLocale.cfg" For Output As #2
Print #2, "en-us"
Close #2
Unload Me
    Load Form1
    Form1.Show
    Exit Sub
Err3:
MkDir ConfigFolder
MkDir ConfigFolder & "\Niveles"
Open ConfigFolder & "\SMMWECloudLocale.cfg" For Output As #2
Print #2, "en-us"
Close #2
Unload Me
    Load Form1
    Form1.Show
End Sub
Private Sub Command3_Click()
On Error GoTo Err4
Open ConfigFolder & "\SMMWECloudLocale.cfg" For Output As #2
Print #2, "es-es"
Close #2
Unload Me
    Load Form1
    Form1.Show
    Exit Sub
Err4:
MkDir ConfigFolder
MkDir ConfigFolder & "\Niveles"
Open ConfigFolder & "\SMMWECloudLocale.cfg" For Output As #2
Print #2, "en-us"
Close #2
Unload Me
    Load Form1
    Form1.Show
End Sub

Private Sub Form_Load()
On Error GoTo Err
LevelFolder = "C:\Users\" & Environ("UserName") & "\AppData\Local\SMM_WE\Niveles"
ConfigFolder = "C:\Users\" & Environ("UserName") & "\AppData\Local\SMM_WE"
DesktopFolder = "C:\Users\" & Environ("UserName") & "\Desktop"
Label1.Caption = "«Î—°‘Òƒ„µƒ”Ô—‘°£" & vbCrLf & "Please select your language." & vbCrLf & "Seleccione su idioma."
    If CheckFileExists(ConfigFolder & "\SMMWECloudLocale.cfg") = True Then
    Load Form1
    Unload Me
    Form1.Show
    End If
    Exit Sub
Err:
MsgBox "Error! Use alternative way to get environment variables."
Dim oShell
Dim strHomeFolder As String
Set oShell = CreateObject("WScript.Shell")
strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
ConfigFolder = strHomeFolder & "\AppData\Local\SMM_WE"
LevelFolder = strHomeFolder & "\AppData\Local\SMM_WE\Niveles"
DesktopFolder = strHomeFolder & "\AppData\Local\SMM_WE\Desktop"
Label1.Caption = "«Î—°‘Òƒ„µƒ”Ô—‘°£" & vbCrLf & "Please select your language." & vbCrLf & "Seleccione su idioma."
    If CheckFileExists(ConfigFolder & "\SMMWECloudLocale.cfg") = True Then
    Load Form1
    Unload Me
    Form1.Show
    End If
End Sub
