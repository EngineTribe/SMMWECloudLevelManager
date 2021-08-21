VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EngineTool"
   ClientHeight    =   6105
   ClientLeft      =   9375
   ClientTop       =   3435
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9960
   Begin VB.CommandButton AboutButton 
      Caption         =   "About"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton RenameButton 
      Caption         =   "Rename"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton UploadButton 
      Caption         =   "Upload"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton DownloadButton 
      Caption         =   "Download"
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "Delete"
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton LocalLevelsButton 
      Caption         =   "LocalLevels"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton OnlineLevelsButton 
      Caption         =   "OnlineLevels"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   2070
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   6135
      Left            =   0
      Picture         =   "Form1.frx":2AFA
      ScaleHeight     =   6075
      ScaleWidth      =   9915
      TabIndex        =   8
      Top             =   0
      Width           =   9975
      Begin VB.CommandButton InfoButton 
         Caption         =   "LevelInfo"
         Height          =   495
         Left            =   8130
         TabIndex        =   11
         Top             =   90
         Width           =   1695
      End
      Begin VB.CommandButton ExtractButton 
         Caption         =   "Extract"
         Height          =   495
         Left            =   8130
         TabIndex        =   10
         Top             =   1890
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   375
         ItemData        =   "Form1.frx":A2F4
         Left            =   2040
         List            =   "Form1.frx":A2F6
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label LevelCounter 
         BackStyle       =   0  'Transparent
         Caption         =   "LevelCounter"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   5520
         Width           =   5895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Function DownloadFile(ByVal strURL As String, ByVal strFile As String) As Boolean
   Dim lngReturn As Long
   lngReturn = URLDownloadToFile(0, strURL, strFile, 0, 0)
   If lngReturn = 0 Then DownloadFile = True
End Function





Private Sub Form_Load()
Version = "1.0"
'设列表背景
List1.BackColor = RGB(240, 252, 250)
Combo1.BackColor = RGB(240, 252, 250)
Dim oShell
Set oShell = CreateObject("WScript.Shell")
LevelFolder = oShell.ExpandEnvironmentStrings("%UserProfile%")
ConfigFolder = LevelFolder & "\AppData\Local\SMM_WE"
LevelFolder = LevelFolder & "\AppData\Local\SMM_WE\Niveles"
    If CheckFileExists(ConfigFolder & "\SMMWECloudLocale.cfg") = True Then
    Dim locale As String
    Open ConfigFolder & "\SMMWECloudLocale.cfg" For Input As #3
    Line Input #3, locale
'加载语言
    If locale = "zh-cn" Then
    LocalLevelsButton.Caption = "本地关卡"
    OnlineLevelsButton.Caption = "在线关卡"
    UploadButton.Caption = "上传"
    AboutButton.Caption = "关于"
    InfoButton.Caption = "关卡信息"
    Title = "SMMWE Cloud 关卡管理器 " & Version
    RenameError = "重命名失败，关卡名字不能留空。"
    ErrorText(0) = "确定要删除 “"
    ErrorText(1) = "” 吗？"
    ErrorText(2) = "请输入 “"
    ErrorText(3) = "” 的新名称。"
    ErrorText(4) = "下载"
    ErrorText(5) = "正在下载中..."
    ErrorText(6) = "下载完成！"
    ErrorText(7) = "重命名"
    ErrorText(8) = "重命名完毕！"
    ErrorText(9) = "删除"
    ErrorText(10) = "删除完毕！"
    ErrorText(11) = "(关卡为3.0.0离线版制作)"
    ErrorText(12) = "关卡作者："
    ErrorText(13) = "关卡场景："
    ErrorText(14) = "游戏风格："
    ErrorText(15) = "标签1："
    ErrorText(16) = "标签2："
    ErrorText(17) = "时间："
    ErrorText(18) = "自动卷轴："
    ErrorText(19) = "慢速"
    ErrorText(20) = "常速"
    ErrorText(21) = "快速"
    ErrorText(22) = "提取"
    ErrorText(23) = "提取完成！"
    ErrorText(24) = "SMMWE Cloud 玩家上传"
    ErrorText(25) = "个关卡"
    ErrorText(26) = "加载中"
    ErrorText(27) = "页数"
    ErrorText(28) = "打开 SMMWE Cloud 网页版"
    ErrorText(29) = "检查更新"
    
    GameLabel(0) = "自动"
    GameLabel(1) = "短小精悍"
    GameLabel(2) = "多人对战"
    GameLabel(3) = "主题"
    GameLabel(4) = "BOSS战"
    GameLabel(5) = "单人"
    GameLabel(6) = "计时挑战"
    GameLabel(7) = "自动卷轴"
    GameLabel(8) = "技巧"
    GameLabel(9) = "射击"
    GameLabel(10) = "音乐"
    GameLabel(11) = "美术"
    GameLabel(12) = "传统"
    GameLabel(13) = "解谜"
    GameLabel(14) = "林克"
    GameLabel(15) = "无"
     JSONError = "(关卡解析失败)"
ElseIf locale = "en-us" Then
     LocalLevelsButton.Caption = "Local Level"
     OnlineLevelsButton.Caption = "Online Level"
     UploadButton.Caption = "Upload"
     AboutButton.Caption = "About"
    InfoButton.Caption = "Level Info"
    Title = "SMMWE Cloud Level Manager & Version"
     RenameError = "Failed to rename, level name cannot be left blank."
     ErrorText(0) = "Are you sure you want to delete " & Chr(34)
     ErrorText(1) = Chr(34) & "? "
     ErrorText(2) = "Please enter the new name for " & Chr(34)
     ErrorText(3) = Chr(34) & ". "
     ErrorText(4) = "Download"
     ErrorText(5) = "Downloading..."
     ErrorText(6) = "Download complete!"
     ErrorText(7) = "Rename"
     ErrorText(8) = "Rename complete!"
     ErrorText(9) = "Delete"
     ErrorText(10) = "Deletion complete!"
    ErrorText(11) = "(Made with 3.0.0 Offline patch)"
    ErrorText(12) = "Maker: "
    ErrorText(13) = "Stage: "
    ErrorText(14) = "Game style: "
    GameLabel(0) = "Automatic"
    GameLabel(1) = "Short but intense"
    GameLabel(2) = "Multiplayer versus"
    GameLabel(3) = "Theme"
    GameLabel(4) = "BOSS fight"
    GameLabel(5) = "Singleplayer"
    GameLabel(6) = "Time challenge"
    GameLabel(7) = "Autoscroll"
    GameLabel(8) = "Skills"
    GameLabel(9) = "Shooting"
    GameLabel(10) = "Music"
    GameLabel(11) = "Art"
    GameLabel(12) = "Traditional"
    GameLabel(13) = "Puzzles"
    GameLabel(14) = "Link"
    GameLabel(15) = "None"
    ErrorText(15) = "Label 1: "
    ErrorText(16) = "Label 2: "
    ErrorText(17) = "Timer: "
    ErrorText(18) = "Autoscroll: "
    ErrorText(19) = "Slow"
    ErrorText(20) = "Normal"
    ErrorText(21) = "Fast"
    ErrorText(22) = "Extract"
    ErrorText(23) = "Completed!"
    ErrorText(24) = "SMMWE Cloud Users Uploaded"
    ErrorText(25) = " Levels"
    ErrorText(26) = "Loading"
    ErrorText(27) = "Page"
    ErrorText(28) = "Open SMMWE Cloud Website"
    ErrorText(29) = "Check Update"
     JSONError = "(Level parsing failed)"
ElseIf locale = "es-es" Then
      LocalLevelsButton.Caption = "Nivel local"
      OnlineLevelsButton.Caption = "Nivel en linea"
      UploadButton.Caption = "Subir Nivel"
    InfoButton.Caption = "info de nivel"
      AboutButton.Caption = "Sobre"
    Title = "SMMWE Cloud Level Manager " & Version
      RenameError = "No se pudo cambiar el nombre, el nombre del nivel no se puede dejar en blanco."
      ErrorText(0) = "Esta seguro de que desea borrar" & Chr(34)
      ErrorText(1) = Chr(34) & "?"
      ErrorText(2) = "Ingrese el nuevo nombre para" & Chr(34)
      ErrorText(3) = Chr(34) & "."
      ErrorText(4) = "Descargar"
      ErrorText(5) = "Descargando ..."
      ErrorText(6) = "Descarga completa!"
      ErrorText(7) = "Cambiar nombre"
      ErrorText(8) = "Cambio de nombre completo!"
      ErrorText(9) = "Borrar"
      ErrorText(10) = "Eliminacion completa!"
      ErrorText(11) = "(Hecho con el parche 3.0.0 sin conexion)"
     ErrorText(12) = "Creador :"
     ErrorText(13) = "Escenario: "
    ErrorText(14) = "Estilo de juego: "
    GameLabel(0) = "Automatismos"
    GameLabel(1) = "Corto pero intenso"
    GameLabel(2) = "Conpetitivo"
    GameLabel(3) = "Tematico"
    GameLabel(4) = "Contra jefes"
    GameLabel(5) = "En solitario"
    GameLabel(6) = "Contrareloj"
    GameLabel(7) = "Autoavance"
    GameLabel(8) = "Habilidad"
    GameLabel(9) = "Disparos"
    GameLabel(10) = "Musica"
    GameLabel(11) = "Artistico"
    GameLabel(12) = "Tradicional"
    GameLabel(13) = "Puzles"
    GameLabel(14) = "Link"
    GameLabel(15) = "Ninguno"
    ErrorText(15) = "Etiqueta 1: "
     ErrorText(16) = "Etiqueta 2: "
     ErrorText(17) = "Cronometro: "
     ErrorText(18) = "Autoavance: "
    ErrorText(19) = "Lento"
    ErrorText(20) = "Normal"
    ErrorText(21) = "Rapido"
    ErrorText(22) = "Extraer"
     ErrorText(23) = "Completado!"
     JSONError = "(Error de analisis de nivel)"
     ErrorText(24) = "SMMWE Cloud subidos"
     ErrorText(25) = " Niveles"
     ErrorText(26) = "Cargando"
    ErrorText(27) = "Pagina"
    ErrorText(28) = "Sitio web abierto SMMWE Cloud"
    ErrorText(29) = "Buscar actualizacion"
    End If
    Close #3
    End If
    DownloadButton.Caption = ErrorText(4)
    RenameButton.Caption = ErrorText(7)
    DeleteButton.Caption = ErrorText(9)
    ExtractButton.Caption = ErrorText(22)
'删除在线关卡列表缓存
    If CheckFileExists(ConfigFolder & "\SMMWECloudLevelList.json") = True Then Kill ConfigFolder & "\SMMWECloudLevelList.json"
'处理界面
    Form1.Caption = Title & " - " & LocalLevelsButton.Caption
DeleteButton.Visible = True
InfoButton.Visible = True
RenameButton.Visible = True
DownloadButton.Visible = False
UploadButton.Visible = False
ExtractButton.Visible = True
Combo1.Visible = False
Combo1.AddItem (ErrorText(24))
Combo1.Text = (ErrorText(24))
List1.Top = 120
List1.Height = 5340
    '加载本地关卡
Dim fname As String
fname = Dir(LevelFolder & "\*.swe", 7)
List1.Clear
Do
If fname = "" Then Exit Do
List1.AddItem Replace(fname, ".swe", "")
fname = Dir()
Loop
ReDim locallevel(0 To List1.ListCount - 1) As String
For i = 0 To List1.ListCount - 1
locallevel(i) = List1.List(i)
Next
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub




Private Sub LocalLevelsButton_click()
'处理界面
DeleteButton.Visible = True
RenameButton.Visible = True
InfoButton.Visible = True
DownloadButton.Visible = False
UploadButton.Visible = False
ExtractButton.Visible = True
    Form1.Caption = Title & " - " & LocalLevelsButton.Caption
Combo1.Visible = False
List1.Top = 120
List1.Height = 5340
'加载本地关卡
List1.Clear
fname = Dir(LevelFolder & "\*.swe", 7)
Do
If fname = "" Then Exit Do
List1.AddItem Replace(fname, ".swe", "")
fname = Dir()
Loop
ReDim locallevel(0 To List1.ListCount - 1)
For i = 0 To List1.ListCount - 1
locallevel(i) = List1.List(i)
Next
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub
Private Sub Combo1_Click()
'在线关卡列表 Combo
    If CheckFileExists(ConfigFolder & "\SMMWECloudLevelList.json") = True Then Kill ConfigFolder & "\SMMWECloudLevelList.json"
List1.Clear
List1.AddItem ErrorText(26)
DoEvents
LevelSourceUrl = "https://cloud.smmwe.ml/main/"

    Debug.Print DownloadFile(LevelSourceUrl & "?filename", ConfigFolder & "\SMMWECloudLevelList.json")
    Dim filelist As String
    Open ConfigFolder & "\SMMWECloudLevelList.json" For Input As #1
    Line Input #1, filelist
    onlinelevel = Split(filelist, vbLf)
    onlinelevel = Filter(onlinelevel, ".swe")
    onlineleveltmp = Join(onlinelevel, vbCrLf)
    onlineleveltmp = Replace(onlineleveltmp, ".swe", "")
    onlinelevel = Split(onlineleveltmp, vbCrLf)
    Dim tmp2 As Integer
    tmp2 = UBound(onlinelevel) - LBound(onlinelevel)
    Dim s As Long, i As Long
List1.Clear
    For i = 0 To tmp2
        List1.AddItem onlinelevel(i)
    Next i
    Close #1
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub
Private Sub OnlineLevelsButton_Click()
'在线关卡按钮
    If CheckFileExists(ConfigFolder & "\SMMWECloudLevelList.json") = True Then Kill ConfigFolder & "\SMMWECloudLevelList.json"
If Combo1.Text = ErrorText(24) Then LevelSourceUrl = "https://cloud.smmwe.ml/main/"

    Form1.Caption = Title & " - " & OnlineLevelsButton.Caption
List1.Clear
List1.AddItem ErrorText(26)
DeleteButton.Visible = False
RenameButton.Visible = False
ExtractButton.Visible = False
DownloadButton.Visible = True
InfoButton.Visible = False
UploadButton.Visible = True
Combo1.Visible = True
List1.Top = 600
List1.Height = 4860
    Debug.Print DownloadFile(LevelSourceUrl & "?filename", ConfigFolder & "\SMMWECloudLevelList.json")
    Dim filelist As String
    Open ConfigFolder & "\SMMWECloudLevelList.json" For Input As #1
    Line Input #1, filelist
    
    onlinelevel = Split(filelist, vbLf)
    onlinelevel = Filter(onlinelevel, ".swe")
    onlineleveltmp = Join(onlinelevel, vbCrLf)
    onlineleveltmp = Replace(onlineleveltmp, ".swe", "")
    onlinelevel = Split(onlineleveltmp, vbCrLf)
    Dim tmp2 As Integer
    tmp2 = UBound(onlinelevel) - LBound(onlinelevel)
    Dim s As Long, i As Long
    
List1.Clear
    For i = 0 To tmp2
        List1.AddItem onlinelevel(i)
    Next i
    Close #1
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub

'删除
Private Sub DeleteButton_Click()
If List1.Text <> "" Then
    IfDelete = MsgBox(ErrorText(0) & List1.Text & ErrorText(1), 1, "")
    If IfDelete = 1 Then
    Kill LevelFolder & "\" & List1.Text & ".swe"
    DeleteButton.Caption = ErrorText(10)
    Sleep (1000)
    DeleteButton.Caption = ErrorText(9)
List1.Clear
fname = Dir(LevelFolder & "\*.swe", 7)
Do
If fname = "" Then Exit Do
List1.AddItem Replace(fname, ".swe", "")
fname = Dir()
Loop
ReDim locallevel(0 To List1.ListCount - 1)
For i = 0 To List1.ListCount - 1
locallevel(i) = List1.List(i)
Next
    End If
End If
End Sub
'下载关卡
Private Sub DownloadButton_Click()
If List1.Text <> "" Then
    DownloadButton.Caption = ErrorText(5)
    Debug.Print DownloadFile(LevelSourceUrl & List1.Text & ".swe", LevelFolder & "\" & List1.Text & ".swe")
    DownloadButton.Caption = ErrorText(6)
    DoEvents
    Sleep (1000)
    DownloadButton.Caption = ErrorText(4)
End If
End Sub

'重命名
Private Sub RenameButton_Click()

If List1.Text <> "" Then
    NewName = InputBox(ErrorText(2) & List1.Text & ErrorText(3), "")
    If NewName <> "" Then
     If NewName <> " " Then
       Name LevelFolder & "\" & List1.Text & ".swe" As LevelFolder & "\" & NewName & ".swe"
    RenameButton.Caption = ErrorText(8)
    DoEvents
    Sleep (1000)
    RenameButton.Caption = ErrorText(7)
        List1.Clear
        fname = Dir(LevelFolder & "\*.swe", 7)
        Do
        If fname = "" Then Exit Do
        List1.AddItem Replace(fname, ".swe", "")
        fname = Dir()
        Loop
        ReDim locallevel(0 To List1.ListCount - 1)
        For i = 0 To List1.ListCount - 1
        locallevel(i) = List1.List(i)
        Next
    Else
        MsgBox RenameError
    End If
    Else
        MsgBox RenameError
    End If
End If
End Sub

Private Sub UploadButton_Click()
Shell "cmd /c start https://cloud.smmwe.ml/upload", vbMinimizedNoFocus
End Sub


Private Sub AboutButton_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub InfoButton_Click()
If List1.Text <> "" Then
    Dim LevelContent As String
    Open LevelFolder & "\" & List1.Text & ".swe" For Input As #4
    Line Input #4, LevelContent
    Close #4
    LevelContent = Base64Decode(LevelContent)
    Debug.Print LevelContent
    Dim LevelMaker As String
    LevelMaker = JSONParse("MAIN.AJUSTES[0].user", LevelContent)
    If LevelMaker = "0" Then LevelMaker = ErrorText(11)
    If LevelMaker = "0.000" Then LevelMaker = ErrorText(11)
    Dim GameStyle As String
    GameStyle = JSONParse("MAIN.AJUSTES[0].apariencia", LevelContent)
    If GameStyle = "0" Then GameStyle = "SMB1"
    If GameStyle = "1" Then GameStyle = "SMB3"
    If GameStyle = "2" Then GameStyle = "SMW"
    If GameStyle = "3" Then GameStyle = "NSMBU"
    gamelabel1 = JSONParse("MAIN.AJUSTES[0].etiqueta1", LevelContent)
    gamelabel2 = JSONParse("MAIN.AJUSTES[0].etiqueta2", LevelContent)
    If gamelabel1 = "0" Then gamelabel1 = GameLabel(12)
    If gamelabel1 = "1" Then gamelabel1 = GameLabel(13)
    If gamelabel1 = "2" Then gamelabel1 = GameLabel(6)
    If gamelabel1 = "3" Then gamelabel1 = GameLabel(7)
    If gamelabel1 = "4" Then gamelabel1 = GameLabel(0)
    If gamelabel1 = "5" Then gamelabel1 = GameLabel(1)
    If gamelabel1 = "6" Then gamelabel1 = GameLabel(3)
    If gamelabel1 = "7" Then gamelabel1 = GameLabel(2)
    If gamelabel1 = "8" Then gamelabel1 = GameLabel(10)
    If gamelabel1 = "9" Then gamelabel1 = GameLabel(11)
    If gamelabel1 = "10" Then gamelabel1 = GameLabel(8)
    If gamelabel1 = "11" Then gamelabel1 = GameLabel(9)
    If gamelabel1 = "12" Then gamelabel1 = GameLabel(4)
    If gamelabel1 = "13" Then gamelabel1 = GameLabel(5)
    If gamelabel1 = "14" Then gamelabel1 = GameLabel(14)
    If gamelabel2 = "0" Then gamelabel2 = GameLabel(12)
    If gamelabel2 = "1" Then gamelabel2 = GameLabel(13)
    If gamelabel2 = "2" Then gamelabel2 = GameLabel(6)
    If gamelabel2 = "3" Then gamelabel2 = GameLabel(7)
    If gamelabel2 = "4" Then gamelabel2 = GameLabel(0)
    If gamelabel2 = "5" Then gamelabel2 = GameLabel(1)
    If gamelabel2 = "6" Then gamelabel2 = GameLabel(3)
    If gamelabel2 = "7" Then gamelabel2 = GameLabel(2)
    If gamelabel2 = "8" Then gamelabel2 = GameLabel(10)
    If gamelabel2 = "9" Then gamelabel2 = GameLabel(11)
    If gamelabel2 = "10" Then gamelabel2 = GameLabel(8)
    If gamelabel2 = "11" Then gamelabel2 = GameLabel(9)
    If gamelabel2 = "12" Then gamelabel2 = GameLabel(4)
    If gamelabel2 = "13" Then gamelabel2 = GameLabel(5)
    If gamelabel2 = "14" Then gamelabel2 = GameLabel(14)
    If gamelabel1 = "-1" Then gamelabel1 = GameLabel(15)
    If gamelabel2 = "-1" Then gamelabel2 = GameLabel(15)
    Dim AutoScroll As String
    AutoScroll = JSONParse("MAIN.AJUSTES[0].autoavance", LevelContent)
    If AutoScroll = "0" Then AutoScroll = GameLabel(15)
    If AutoScroll = "1" Then AutoScroll = ErrorText(19)
    If AutoScroll = "2" Then AutoScroll = ErrorText(20)
    If AutoScroll = "3" Then AutoScroll = ErrorText(21)
    MsgBox ErrorText(12) & JSONParse("MAIN.AJUSTES[0].user", LevelContent) & vbCrLf & ErrorText(13) & JSONParse("MAIN.AJUSTES[0].entorno", LevelContent) _
    & vbCrLf & ErrorText(13) & GameStyle & vbCrLf & ErrorText(15) & gamelabel1 & "  " & ErrorText(16) & gamelabel2 & vbCrLf & ErrorText(17) & JSONParse("MAIN.AJUSTES[0].cronometro", LevelContent) _
    & vbCrLf & ErrorText(18) & AutoScroll _
    , vbOKOnly, InfoButton.Caption
End If
End Sub
Private Sub ExtractButton_Click()

If List1.Text <> "" Then
Dim oShell
Set oShell = CreateObject("WScript.Shell")
DesktopFolder = oShell.ExpandEnvironmentStrings("%UserProfile%")
DesktopFolder = DesktopFolder & "\Desktop"
FileCopy LevelFolder & "\" & List1.Text & ".swe", DesktopFolder & "\" & List1.Text & ".swe"
    ExtractButton.Caption = ErrorText(23)
    DoEvents
    Sleep (1000)
    ExtractButton.Caption = ErrorText(22)
End If
End Sub

