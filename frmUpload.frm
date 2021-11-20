VERSION 5.00
Begin VB.Form frmUpload 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8460
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox CustomNameText 
      Appearance      =   0  'Flat
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3360
      Width           =   7455
   End
   Begin VB.OptionButton OpCustomName 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   375
   End
   Begin VB.OptionButton OpWithMaker 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   375
   End
   Begin VB.OptionButton OpOriginalName 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.Label NoLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6200
      TabIndex        =   8
      Top             =   4440
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
      Left            =   3800
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Image NoButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6120
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Image YesButton 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3720
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label WithMakerLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   7815
   End
   Begin VB.Label OrigNameLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   7815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CustomNameText_Click()
If CustomNameText.Text = ConstStr(40) Then CustomNameText.Text = ""
OpCustomName.Value = True
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
Me.BackColor = RGB(250, 228, 192)
Me.Caption = ConstStr(38)
'Init
Label1.Caption = ConstStr(39)
Label1.ForeColor = RGB(89, 15, 16)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
OrigNameLabel.ForeColor = RGB(89, 15, 16)
OrigNameLabel.Font.Name = "DinkieBitmap 9pxDemo"
WithMakerLabel.ForeColor = RGB(89, 15, 16)
WithMakerLabel.Font.Name = "DinkieBitmap 9pxDemo"
OpOriginalName.BackColor = RGB(250, 228, 192)
OpOriginalName.Value = True
OrigNameLabel.Caption = NoCHS(frmLevel.LevelName.Caption)
OpWithMaker.BackColor = RGB(250, 228, 192)
OpWithMaker.Value = False
If frmLevel.LevelMakerLabel.Caption = ConstStr(3) Then
WithMakerLabel.Caption = NoCHS(frmLevel.LevelName.Caption & " (Offline)")
Else
WithMakerLabel.Caption = NoCHS(frmLevel.LevelName.Caption & " By " & Right(frmLevel.LevelMakerLabel.Caption, Len(frmLevel.LevelMakerLabel.Caption) - 1))
End If
OpCustomName.BackColor = RGB(250, 228, 192)
OpCustomName.Value = False
CustomNameText.ForeColor = RGB(89, 15, 16)
CustomNameText.Font.Name = "DinkieBitmap 9pxDemo"
CustomNameText.Text = ConstStr(40)
YesLabel.Caption = ConstStr(37)
NoLabel.Caption = ConstStr(8)
YesLabel.ForeColor = RGB(89, 15, 16)
YesLabel.Font.Name = "DinkieBitmap 9pxDemo"
NoLabel.ForeColor = RGB(89, 15, 16)
NoLabel.Font.Name = "DinkieBitmap 9pxDemo"
NoButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
YesButton.Picture = StdPictureEx.LoadPicture(App.Path & "\Assets\btn-blank.png")
End Sub

Private Function NoCHS(ByVal FilterString As String) As String
'NoCHS function
'https://zhidao.baidu.com/question/215459751.html
Dim s As String, Tmp As String
Dim I As Long
For I = 1 To Len(FilterString)
Tmp = Mid(FilterString, I, 1)
If Asc(Tmp) < 255 And Asc(Tmp) > 0 Then
s = s & Tmp
End If
Next
NoCHS = s
End Function

Private Sub NoButton_Click()
PlaySFX "snd_close_guardabot.ogg"
Unload Me
End Sub

Private Sub OpCustomName_Click()
CustomNameText.SetFocus
If CustomNameText.Text = ConstStr(40) Then CustomNameText.Text = ""
End Sub

Private Sub OrigNameLabel_Click()
OpOriginalName.Value = True
End Sub

Private Sub WithMakerLabel_Click()
OpWithMaker.Value = True
End Sub

Private Sub YesButton_Click()
Dim UploadFileName As String
If OpOriginalName Then
UploadFileName = OrigNameLabel.Caption
End If
If OpWithMaker Then
UploadFileName = WithMakerLabel.Caption
End If
If OpCustomName Then
If CustomNameText.Text = ConstStr(40) Or CustomNameText.Text = "" Then
PlaySFX "snd_wrong.ogg"
Exit Sub
Else
UploadFileName = CustomNameText.Text
End If
End If
If MsgBox(ConstStr(41), vbOKCancel + vbExclamation, ConstStr(42)) = vbCancel Then
PlaySFX "snd_close_guardabot.ogg"
frmUpload.SetFocus
Exit Sub
End If
'Test if exists
frmUpload.Hide
DoEvents
ShowMsgBox "UPLOADING"
DoEvents
PlaySFX "snd_aceptar.ogg"
Dim CanUpload As Boolean, LevelContentTmp As String
CanUpload = False
If CheckFileExists(ConfigFolder & "\" & ".Upload.tmp") = True Then Kill ConfigFolder & "\" & ".Upload.tmp"
Call URLDownloadToFile(0, OLAPIIP & "smmweroot/" & Replace(UploadFileName, " ", "%20") & ".swe", ConfigFolder & "\" & ".Upload.tmp", 0, 0)
 If CheckFileExists(ConfigFolder & "\" & ".Upload.tmp") = False Then
        CanUpload = True
    Else
        Open ConfigFolder & "\" & ".Upload.tmp" For Input As #7
        Input #7, LevelContentTmp
        Close #7
        LevelContentTmp = Join(Filter(Split(LevelContentTmp, vbLf), "itemNotFound"), "")
        If LevelContentTmp <> "" Then CanUpload = True
    End If
    If CanUpload = False Then
    If OpOriginalName Then
    UploadFileName = WithMakerLabel.Caption
    Else
    UploadFileName = UploadFileName & " (2)"
    End If
    End If
    LevelContentTmp = ""
        Open LevelFolder & "\" & frmLevel.LevelName.Caption & ".swe" For Input As #13
        Line Input #13, LevelContentTmp
        Close #13
        Debug.Print PostDataSWE(OLAPIIP & "smmweroot/?upload=" & Replace(UploadFileName, " ", "%20") & ".swe&key=yidaozhan-gq-franyer-farias-apiv2", LevelContentTmp)
        PlaySFX "snd_open_niveles_mundiales.ogg"
DeleteUrlCacheEntry ("https://apiv2.smmwe.ml/smmweroot/")
        frmMsgBox.Hide
        Unload frmMsgBox
        Sleep 10
ShowMsgBox "SUCCESS2"
Unload Me
End Sub

Private Sub YesLabel_Click()
YesButton_Click
End Sub

Private Sub NoLabel_Click()
NoButton_Click
End Sub

