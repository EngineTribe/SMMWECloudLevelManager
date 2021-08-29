VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMMWECloudLevelManager"
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
   Begin VB.CommandButton Download2Button 
      Caption         =   "OK"
      Height          =   495
      Left            =   4530
      TabIndex        =   24
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton IDConfirm 
      Caption         =   "OK"
      Height          =   495
      Left            =   4560
      TabIndex        =   22
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox ID4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   20
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox ID3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   19
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox ID2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   18
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox ID1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   17
      Top             =   2280
      Width           =   615
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
      Height          =   4260
      Left            =   2040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton PageButton 
      Caption         =   "Page"
      Height          =   495
      Left            =   8160
      TabIndex        =   12
      Top             =   720
      Width           =   1695
   End
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
      Top             =   3120
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
      Begin VB.CommandButton OfficialButton 
         Caption         =   "Official"
         Height          =   495
         Left            =   90
         TabIndex        =   16
         Top             =   1290
         Width           =   1695
      End
      Begin VB.CommandButton ImportButton 
         Caption         =   "Import"
         Height          =   495
         Left            =   8130
         TabIndex        =   15
         Top             =   1890
         Width           =   1695
      End
      Begin VB.PictureBox SearchButton 
         Height          =   375
         Left            =   7560
         Picture         =   "Form1.frx":A2F4
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Search 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton InfoButton 
         Caption         =   "LevelInfo"
         Height          =   495
         Left            =   8130
         TabIndex        =   10
         Top             =   90
         Width           =   1695
      End
      Begin VB.CommandButton ExtractButton 
         Caption         =   "Extract"
         Height          =   495
         Left            =   8130
         TabIndex        =   9
         Top             =   2490
         Width           =   1695
      End
      Begin VB.Label LevelInfos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   3975
         Left            =   2040
         TabIndex        =   23
         Top             =   120
         Width           =   7575
      End
      Begin VB.Label IDToolTip 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   1680
         Width           =   3255
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
         TabIndex        =   11
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
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Function DownloadFile(ByVal StrUrl As String, ByVal strFile As String) As Boolean
   Dim lngReturn As Long
   DoEvents
   lngReturn = URLDownloadToFile(0, StrUrl, strFile, 0, 0)
   DoEvents
   If lngReturn = 0 Then DownloadFile = True
End Function
Public Function PostData(ByVal StrUrl As String, ByVal StrData As String) As Variant
  On Error GoTo Err:
  Dim XMLHTTP As Object
  Dim DataS As String
  Dim DataB() As Byte
  Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
  XMLHTTP.Open "POST", StrUrl, True
  XMLHTTP.send (StrData)
  Do Until XMLHTTP.ReadyState = 4
    DoEvents
    Sleep (5)
  Loop
PostData = "Completed"
  Set XMLHTTP = Nothing
  Exit Function
Err:
  PostData = ""
End Function





Private Sub Form_Load()
'如果没关卡就跳过rt9
On Error Resume Next
Version = "4.0"
'设列表背景
List1.BackColor = RGB(240, 252, 250)
Search.BackColor = RGB(240, 252, 250)
LevelInfos.BackColor = RGB(240, 252, 250)
   MkDir ConfigFolder
MkDir LevelFolder
    
    If CheckFileExists(ConfigFolder & "\SMMWECloudLocale.cfg") = True Then
    Open ConfigFolder & "\SMMWECloudLocale.cfg" For Input As #3
    Line Input #3, Locale
'加载语言
    If Locale = "zh-cn" Then
    LocalLevelsButton.Caption = "本地关卡"
    OnlineLevelsButton.Caption = "在线关卡"
    OfficialButton.Caption = "官方服务器"
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
    ErrorText(26) = "加载中..."
    ErrorText(27) = "页数"
    ErrorText(28) = "打开 SMMWE Cloud 网页版"
    ErrorText(29) = "检查更新"
    ErrorText(30) = "在本页中搜索..."
    ErrorText(31) = "导入"
    ErrorText(32) = "导入完成！"
    ErrorText(33) = "关卡"
    ErrorText(34) = "取消"
    ErrorText(35) = "这个文件夹中没有关卡文件。"
    ErrorText(36) = "将会上传 "
    ErrorText(37) = " 到 SMMWE Cloud。确定？"
    ErrorText(38) = "上传成功！"
    ErrorText(39) = "上传"
    ErrorText(40) = "上传中..."
    ErrorText(41) = "请输入关卡ID。"
    ErrorText(42) = " (重试第"
    ErrorText(43) = "次)"
    ErrorText(44) = "关卡名字："
    ErrorText(45) = "赞数："
    ErrorText(46) = "  踩数："
    ErrorText(47) = "通关数："
    ErrorText(48) = "  失败数："
    ErrorText(49) = "  通关率："
    
    GameLabel(0) = "自动马力欧"
    GameLabel(1) = "一次通过"
    GameLabel(2) = "多人对战"
    GameLabel(3) = "机关设计"
    GameLabel(4) = "BOSS战"
    GameLabel(5) = "单打"
    GameLabel(6) = "计时挑战"
    GameLabel(7) = "自动卷轴"
    GameLabel(8) = "技巧"
    GameLabel(9) = "射击"
    GameLabel(10) = "音乐"
    GameLabel(11) = "美术"
    GameLabel(12) = "标准"
    GameLabel(13) = "解谜"
    GameLabel(14) = "林克"
    GameLabel(15) = "无"
    GameLabel(16) = "地面"
    GameLabel(17) = "地下"
    GameLabel(18) = "天空"
    GameLabel(19) = "丛林"
    GameLabel(20) = "沙漠"
    GameLabel(21) = "城堡"
    GameLabel(22) = "鬼屋"
    GameLabel(23) = "飞行船"
    GameLabel(24) = "水中"
    GameLabel(25) = "雪原"
    GameLabel(26) = "秋天"
    GameLabel(27) = "白天"
    GameLabel(28) = "夜晚"
ElseIf Locale = "en-us" Then
     LocalLevelsButton.Caption = "Local Level"
     OnlineLevelsButton.Caption = "SMMWE Cloud"
    OfficialButton.Caption = "Course World"
     UploadButton.Caption = "Upload Level"
     AboutButton.Caption = "About"
    InfoButton.Caption = "Level Info"
    Title = "SMMWE Cloud Level Manager " & Version
     RenameError = "Failed to rename, level name cannot be left blank."
     ErrorText(0) = "Are you sure you want to delete " & Chr(34)
     ErrorText(1) = Chr(34) & "? "
     ErrorText(2) = "Please enter the new name for " & Chr(34)
     ErrorText(3) = Chr(34) & ". "
     ErrorText(4) = "Download"
     ErrorText(5) = "Downloading..."
     ErrorText(6) = "Completed!"
     ErrorText(7) = "Rename"
     ErrorText(8) = "Completed!"
     ErrorText(9) = "Delete"
     ErrorText(10) = "Completed!"
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
    GameLabel(16) = "Ground"
    GameLabel(17) = "Underground"
    GameLabel(18) = "Athletic"
    GameLabel(19) = "Jungle"
    GameLabel(20) = "Desert"
    GameLabel(21) = "Castle"
    GameLabel(22) = "Ghost House"
    GameLabel(23) = "Airship"
    GameLabel(24) = "Underwater"
    GameLabel(25) = "Snow"
    GameLabel(26) = "Autumn"
    GameLabel(27) = "Day"
    GameLabel(28) = "Night"
    ErrorText(15) = "Label 1: "
    ErrorText(16) = "Label 2: "
    ErrorText(17) = "Timer: "
    ErrorText(18) = "Autoscroll: "
    ErrorText(19) = "Slow"
    ErrorText(20) = "Normal"
    ErrorText(21) = "Fast"
    ErrorText(22) = "Export"
    ErrorText(23) = "Completed!"
    ErrorText(24) = "SMMWE Cloud Users Uploaded"
    ErrorText(25) = " Levels"
    ErrorText(26) = "Loading..."
    ErrorText(27) = "Page"
    ErrorText(28) = "Open SMMWE Cloud Website"
    ErrorText(29) = "Check Update"
    ErrorText(30) = "Search in this page..."
    ErrorText(31) = "Import"
    ErrorText(32) = "Completed!"
    ErrorText(33) = "Level"
    ErrorText(34) = "Cancel"
    ErrorText(35) = "No level file was found in that directory."
    ErrorText(36) = "Will upload "
    ErrorText(37) = " to SMMWE Cloud, sure?"
    ErrorText(38) = "Upload Completed!"
    ErrorText(39) = "Upload Level"
    ErrorText(40) = "Uploading..."
    ErrorText(41) = "Please enter the level ID."
    ErrorText(42) = " (Retry "
    ErrorText(43) = " times)"
    ErrorText(44) = "Name: "
    ErrorText(45) = "Likes: "
    ErrorText(46) = "  Dislikes: "
    ErrorText(47) = "Clears:"
    ErrorText(48) = "  Fails:"
    ErrorText(49) = "  Clear Rate:"
ElseIf Locale = "es-es" Then
      LocalLevelsButton.Caption = "Niveles local"
      OnlineLevelsButton.Caption = "SMMWE Cloud"
    OfficialButton.Caption = "Niveles Mundial"
      UploadButton.Caption = "Subir Nivel"
    InfoButton.Caption = "Info de Nivel"
      AboutButton.Caption = "Sobre"
    Title = "SMMWE Cloud Level Manager " & Version
      RenameError = "No se pudo cambiar el nombre, el nombre del nivel no se puede dejar en blanco."
      ErrorText(0) = "Esta seguro de que desea borrar" & Chr(34)
      ErrorText(1) = Chr(34) & "?"
      ErrorText(2) = "Ingrese el nuevo nombre para" & Chr(34)
      ErrorText(3) = Chr(34) & "."
      ErrorText(4) = "Descargar"
      ErrorText(5) = "Descargando ..."
      ErrorText(6) = "Completado!"
      ErrorText(7) = "Cambiar nombre"
      ErrorText(8) = "Completado!"
      ErrorText(9) = "Borrar"
      ErrorText(10) = "Completado!"
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
    GameLabel(16) = "Ground"
    GameLabel(17) = "Underground"
    GameLabel(18) = "Athletic"
    GameLabel(19) = "Jungle"
    GameLabel(20) = "Desert"
    GameLabel(21) = "Castle"
    GameLabel(22) = "Ghost House"
    GameLabel(23) = "Airship"
    GameLabel(24) = "Underwater"
    GameLabel(25) = "Snow"
    GameLabel(26) = "Autumn"
    GameLabel(27) = "Dia"
    GameLabel(28) = "Noche"
    ErrorText(15) = "Etiqueta 1: "
     ErrorText(16) = "Etiqueta 2: "
     ErrorText(17) = "Cronometro: "
     ErrorText(18) = "Autoavance: "
    ErrorText(19) = "Lento"
    ErrorText(20) = "Normal"
    ErrorText(21) = "Rapido"
    ErrorText(22) = "Exportar"
     ErrorText(23) = "Completado!"
     ErrorText(24) = "SMMWE Cloud subidos"
     ErrorText(25) = " Niveles"
     ErrorText(26) = "Cargando..."
    ErrorText(27) = "Pagina"
    ErrorText(28) = "Sitio web abierto SMMWE Cloud"
    ErrorText(29) = "Buscar actualizacion"
    ErrorText(30) = "Buscar en esta pagina..."
    ErrorText(31) = "Importar"
    ErrorText(32) = "Completado!"
    ErrorText(33) = "Nivel"
    ErrorText(34) = "Cancelar"
    ErrorText(35) = "No se encontro ningun archivo de nivel en ese directorio."
    ErrorText(36) = "Subira "
    ErrorText(37) = " a SMMWE Cloud, seguro?"
    ErrorText(38) = "Subir completado!"
    ErrorText(39) = "Subir Nivel"
    ErrorText(40) = "Subiendo..."
    ErrorText(41) = "Ingrese el ID de nivel."
    ErrorText(42) = " (Reintentar "
    ErrorText(43) = " veces)"
    ErrorText(44) = "Nombre: "
    ErrorText(45) = "Gustos: "
    ErrorText(46) = "  Disgustos: "
    ErrorText(47) = "Victorias:"
    ErrorText(48) = "  Muertes:"
    ErrorText(49) = "  Tasa clara:"
    End If
    Close #3
    End If
    DownloadButton.Caption = ErrorText(4)
    RenameButton.Caption = ErrorText(7)
    DeleteButton.Caption = ErrorText(9)
    ExtractButton.Caption = ErrorText(22)
    ImportButton.Caption = ErrorText(31)
'删除在线关卡列表缓存
    If CheckFileExists(ConfigFolder & "\SMMWECloudLevelList.txt") = True Then Kill ConfigFolder & "\SMMWECloudLevelList.txt"
    If CheckFileExists(ConfigFolder & "\SMMWEParseTemp.json") = True Then Kill ConfigFolder & "\SMMWEParseTemp.json"
    If CheckFileExists(ConfigFolder & "\SMMWEAuthorImg.png") = True Then Kill ConfigFolder & "\SMMWEAuthorImg.png"
'处理界面
    Form1.Caption = Title & " - " & LocalLevelsButton.Caption
DeleteButton.Visible = True
InfoButton.Visible = True
RenameButton.Visible = True
DownloadButton.Visible = False
UploadButton.Visible = True
ExtractButton.Visible = True
List1.Top = 120
List1.Height = 5340
PageButton.Visible = False
SearchButton.Visible = False
ImportButton.Visible = True
LevelCounter.Visible = True
List1.Visible = True
Search.Visible = False
ID1.Visible = False
ID2.Visible = False
ID3.Visible = False
ID4.Visible = False
IDToolTip.Visible = False
IDConfirm.Visible = False
LevelInfos.Visible = False
Download2Button.Visible = False

Search.Text = ErrorText(30)
Search.ForeColor = RGB(130, 130, 130)

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
For I = 0 To List1.ListCount - 1
locallevel(I) = List1.List(I)
Next
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub



Private Sub LocalLevelsButton_click()
'如果没关卡就跳过rt9
On Error Resume Next
'处理界面
DeleteButton.Visible = True
RenameButton.Visible = True
ImportButton.Visible = True
InfoButton.Visible = True
DownloadButton.Visible = False
UploadButton.Visible = True
ExtractButton.Visible = True
SearchButton.Visible = False
PageButton.Visible = False
Search.Visible = False
List1.Visible = True
LevelCounter.Visible = True
ID1.Visible = False
ID2.Visible = False
ID3.Visible = False
IDToolTip.Visible = False
ID4.Visible = False
IDConfirm.Visible = False
Download2Button.Visible = False
LevelInfos.Visible = False
    Form1.Caption = Title & " - " & LocalLevelsButton.Caption
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
For I = 0 To List1.ListCount - 1
locallevel(I) = List1.List(I)
Next
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub


Private Sub OnlineLevelsButton_Click()
'在线关卡按钮
    If CheckFileExists(ConfigFolder & "\SMMWECloudLevelList.txt") = True Then Kill ConfigFolder & "\SMMWECloudLevelList.txt"
LevelSourceUrl = "https://cloud.smmwe.ml/main/"

    Form1.Caption = Title & " - " & OnlineLevelsButton.Caption
List1.Clear
List1.AddItem ErrorText(26)
DeleteButton.Visible = False
RenameButton.Visible = False
ExtractButton.Visible = False
DownloadButton.Visible = True
InfoButton.Visible = False
ImportButton.Visible = False
UploadButton.Visible = False
PageButton.Visible = True
Search.Visible = True
SearchButton.Visible = True
List1.Visible = True
LevelCounter.Visible = True
ID1.Visible = False
ID2.Visible = False
ID3.Visible = False
ID4.Visible = False
IDToolTip.Visible = False
IDConfirm.Visible = False
LevelInfos.Visible = False
Download2Button.Visible = False
List1.Top = 600
Search.Text = ErrorText(30)
Search.ForeColor = RGB(130, 130, 130)
List1.Height = 4860
'拉取页数
PageNumber = 1
DoEvents
    Debug.Print DownloadFile("https://cloud.smmwe.ml/main/?filename", ConfigFolder & "\SMMWECloudLevelList.txt")
    Dim pagelist As String
    Open ConfigFolder & "\SMMWECloudLevelList.txt" For Input As #6
    Line Input #6, pagelist
    onlinepage = Split(pagelist, vbLf)
    onlinepage = Filter(onlinepage, "Levels Page")
    PageNumberMax = UBound(onlinepage) - LBound(onlinepage) + 2
    Close #6
    '拉取关卡
    Dim filelist As String
    Open ConfigFolder & "\SMMWECloudLevelList.txt" For Input As #1
    Line Input #1, filelist
    OnlineLevel = Split(filelist, vbLf)
    OnlineLevel = Filter(OnlineLevel, ".swe")
    onlineleveltmp = Join(OnlineLevel, vbCrLf)
    onlineleveltmp = Replace(onlineleveltmp, ".swe", "")
    OnlineLevel = Split(onlineleveltmp, vbCrLf)
    Dim tmp2 As Integer
    tmp2 = UBound(OnlineLevel) - LBound(OnlineLevel)
    Dim s As Long, I As Long
    
List1.Clear
    For I = 0 To tmp2
        List1.AddItem OnlineLevel(I)
    Next I
    Close #1
    PageButton.Caption = ErrorText(27) & " " & CStr(PageNumber) & "/" & CStr(PageNumberMax)
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub
'导入 调用资源管理器
Private Sub ImportButton_Click()
frmOpen.Show
End Sub
'删除
Private Sub DeleteButton_Click()
On Error Resume Next
If List1.Text <> "" Then
    IfDelete = MsgBox(ErrorText(0) & List1.Text & ErrorText(1), 1, "")
    If IfDelete = 1 Then
    Kill LevelFolder & "\" & List1.Text & ".swe"
    DeleteButton.Caption = ErrorText(10)
    DoEvents
    Sleep (700)
    DeleteButton.Caption = ErrorText(9)
List1.Clear
fname = Dir(LevelFolder & "\*.swe", 7)
Do
If fname = "" Then Exit Do
List1.AddItem Replace(fname, ".swe", "")
fname = Dir()
Loop
ReDim locallevel(0 To List1.ListCount - 1)
For I = 0 To List1.ListCount - 1
locallevel(I) = List1.List(I)
Next
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
    End If
End If
End Sub
'下载关卡
Private Sub DownloadButton_Click()
If List1.Text <> "" Then
    DownloadButton.Caption = ErrorText(5)
    Dim DownloadFileName As String
    DownloadFileName = List1.Text & ".swe"
    If CheckFileExists(LevelFolder & "\" & List1.Text & ".swe") = True Then DownloadFileName = List1.Text & " (1).swe"
    Debug.Print DownloadFile(LevelSourceUrl & Replace(List1.Text, " ", "%20") & ".swe", LevelFolder & "\" & DownloadFileName)
    DownloadButton.Caption = ErrorText(6)
    DoEvents
    Sleep (500)
    DownloadButton.Caption = ErrorText(4)
End If
End Sub


'重命名
Private Sub RenameButton_Click()
On Error Resume Next
If List1.Text <> "" Then
    NewName = InputBox(ErrorText(2) & List1.Text & ErrorText(3), "")
    If NewName <> "" Then
     If NewName <> " " Then
        If CheckFileExists(LevelFolder & "\" & NewName & ".swe") = True Then NewName = NewName & "(1)"
        Name LevelFolder & "\" & List1.Text & ".swe" As LevelFolder & "\" & NewName & ".swe"
        RenameButton.Caption = ErrorText(8)
        DoEvents
        Sleep (700)
        RenameButton.Caption = ErrorText(7)
        List1.Clear
        fname = Dir(LevelFolder & "\*.swe", 7)
        Do
        If fname = "" Then Exit Do
        List1.AddItem Replace(fname, ".swe", "")
        fname = Dir()
        Loop
        ReDim locallevel(0 To List1.ListCount - 1)
        For I = 0 To List1.ListCount - 1
        locallevel(I) = List1.List(I)
        Next
    Else
        MsgBox RenameError
    End If
    Else
        MsgBox RenameError
    End If
End If
End Sub
Private Sub Search_Click()
If Search.Text = ErrorText(30) Then
Search.Text = ""
Search.ForeColor = RGB(0, 0, 0)
End If
End Sub

Private Sub SearchButton_Click()
'执行搜索
If Search.Text <> (ErrorText(30)) Then
ReDim OnlineLevelSearched(0 To List1.ListCount - 1)
For I = 0 To List1.ListCount - 1
OnlineLevelSearched(I) = List1.List(I)
Next
OnlineLevelSearched2 = Filter(OnlineLevelSearched, Search.Text)
    Dim tmp3 As Integer
    tmp3 = UBound(OnlineLevelSearched2) - LBound(OnlineLevelSearched2)
List1.Clear
    For I = 0 To tmp3
        List1.AddItem OnlineLevelSearched2(I)
    Next I
End If

If Search.Text = "" Then
    Open ConfigFolder & "\SMMWECloudLevelList.txt" For Input As #1
    Line Input #1, filelist
    OnlineLevel = Split(filelist, vbLf)
    OnlineLevel = Filter(OnlineLevel, ".swe")
    onlineleveltmp = Join(OnlineLevel, vbCrLf)
    onlineleveltmp = Replace(onlineleveltmp, ".swe", "")
    OnlineLevel = Split(onlineleveltmp, vbCrLf)
    Dim tmp2 As Integer
    tmp2 = UBound(OnlineLevel) - LBound(OnlineLevel)
    
List1.Clear
    For I = 0 To tmp2
        List1.AddItem OnlineLevel(I)
    Next I
    Close #1
End If

LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub

Private Sub UploadButton_Click()
On Error Resume Next
'上传
Dim CanUpload As Boolean, LevelContentTmp As String
If MsgBox(ErrorText(36) & List1.Text & ErrorText(37), vbOKCancel + vbExclamation, ErrorText(39)) = vbCancel Then GoTo Err6
UploadButton.Caption = ErrorText(40)
CanUpload = False
If List1.Text <> "" Then
    Debug.Print DownloadFile("https://cloud.smmwe.ml/main/" & Replace(List1.Text, " ", "%20") & ".swe", ConfigFolder & "\" & List1.Text & ".tmp")
    If CheckFileExists(ConfigFolder & "\" & List1.Text & ".tmp") = False Then
        CanUpload = True
    Else
        Open ConfigFolder & "\" & List1.Text & ".tmp" For Input As #7
        Line Input #7, LevelContentTmp
        Close #7
        LevelContentTmp = Join(Filter(Split(LevelContentTmp, vbLf), "itemNotFound"), "")
        If LevelContentTmp = "<h1>itemNotFound</h1>" Then CanUpload = True
    End If
    Kill ConfigFolder & "\" & List1.Text & ".tmp"
        Open LevelFolder & "\" & List1.Text & ".swe" For Input As #8
        Line Input #8, LevelContentTmp
        Close #8
    If CanUpload = True Then
        Debug.Print PostData("https://api.smmwe.ml/?upload=" & Replace(List1.Text, " ", "%20") & ".swe&key=yidaozhan-gq-franyer-farias", LevelContentTmp)
    Else
        Dim LevelMaker As String
        LevelMaker = Replace(Join(Filter(Split(LevelContentTmp, ","), Chr(34) & "user" & Chr(34)), ""), Chr(34) & "user" & Chr(34) & ": ", "")
        LevelMaker = Replace(LevelMaker, Chr(34), "")
        Debug.Print PostData("https://api.smmwe.ml/?upload=" & Replace(List1.Text, " ", "%20") & " By" & LevelMaker & ".swe&key=yidaozhan-gq-franyer-farias", LevelContentTmp)
    End If
    UploadButton.Caption = ErrorText(39)
    MsgBox ErrorText(38)
End If
Exit Sub
Err6:
End Sub


Private Sub AboutButton_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub InfoButton_Click()
'关卡信息
If List1.Text <> "" Then
    Dim LevelContent As String
    Open LevelFolder & "\" & List1.Text & ".swe" For Input As #4
    Line Input #4, LevelContent
    Close #4
    LevelContent = Base64Decode(LevelContent)
    levelcontent2 = Split(LevelContent, ",")
    Dim LevelMaker As String
    LevelMaker = Replace(Join(Filter(levelcontent2, Chr(34) & "user" & Chr(34)), ""), Chr(34) & "user" & Chr(34) & ": ", "")
    LevelMaker = Replace(LevelMaker, Chr(34), "")
    If LevelMaker = " 0" Then LevelMaker = ErrorText(11)
    If LevelMaker = " 0.000" Then LevelMaker = ErrorText(11)
    If LevelMaker = " " Then LevelMaker = ErrorText(11)
    If LevelMaker = "" Then LevelMaker = ErrorText(11)
    Dim GameStyle As String
    GameStyle = Replace(Join(Filter(levelcontent2, Chr(34) & "apariencia" & Chr(34)), ""), Chr(34) & "apariencia" & Chr(34) & ": ", "")
    GameStyle = Replace(GameStyle, " } ]", "")
    If GameStyle = " 0" Then GameStyle = "SMB1"
    If GameStyle = " 1" Then GameStyle = "SMB3"
    If GameStyle = " 2" Then GameStyle = "SMW"
    If GameStyle = " 3" Then GameStyle = "NSMBU"
    Dim GameLabel1, GameLabel2 As String
    GameLabel1 = Replace(Join(Filter(levelcontent2, Chr(34) & "etiqueta1" & Chr(34)), ""), Chr(34) & "etiqueta1" & Chr(34) & ": ", "")
    GameLabel2 = Replace(Join(Filter(levelcontent2, Chr(34) & "etiqueta2" & Chr(34)), ""), Chr(34) & "etiqueta2" & Chr(34) & ": ", "")
    If GameLabel1 = " 0" Then GameLabel1 = GameLabel(12)
    If GameLabel1 = " 1" Then GameLabel1 = GameLabel(13)
    If GameLabel1 = " 2" Then GameLabel1 = GameLabel(6)
    If GameLabel1 = " 3" Then GameLabel1 = GameLabel(7)
    If GameLabel1 = " 4" Then GameLabel1 = GameLabel(0)
    If GameLabel1 = " 5" Then GameLabel1 = GameLabel(1)
    If GameLabel1 = " 6" Then GameLabel1 = GameLabel(3)
    If GameLabel1 = " 7" Then GameLabel1 = GameLabel(2)
    If GameLabel1 = " 8" Then GameLabel1 = GameLabel(10)
    If GameLabel1 = " 9" Then GameLabel1 = GameLabel(11)
    If GameLabel1 = " 10" Then GameLabel1 = GameLabel(8)
    If GameLabel1 = " 11" Then GameLabel1 = GameLabel(9)
    If GameLabel1 = " 12" Then GameLabel1 = GameLabel(4)
    If GameLabel1 = " 13" Then GameLabel1 = GameLabel(5)
    If GameLabel1 = " 14" Then GameLabel1 = GameLabel(14)
    If GameLabel2 = " 0" Then GameLabel2 = GameLabel(12)
    If GameLabel2 = " 1" Then GameLabel2 = GameLabel(13)
    If GameLabel2 = " 2" Then GameLabel2 = GameLabel(6)
    If GameLabel2 = " 3" Then GameLabel2 = GameLabel(7)
    If GameLabel2 = " 4" Then GameLabel2 = GameLabel(0)
    If GameLabel2 = " 5" Then GameLabel2 = GameLabel(1)
    If GameLabel2 = " 6" Then GameLabel2 = GameLabel(3)
    If GameLabel2 = " 7" Then GameLabel2 = GameLabel(2)
    If GameLabel2 = " 8" Then GameLabel2 = GameLabel(10)
    If GameLabel2 = " 9" Then GameLabel2 = GameLabel(11)
    If GameLabel2 = " 10" Then GameLabel2 = GameLabel(8)
    If GameLabel2 = " 11" Then GameLabel2 = GameLabel(9)
    If GameLabel2 = " 12" Then GameLabel2 = GameLabel(4)
    If GameLabel2 = " 13" Then GameLabel2 = GameLabel(5)
    If GameLabel2 = " 14" Then GameLabel2 = GameLabel(14)
    If GameLabel1 = " -1" Then GameLabel1 = GameLabel(15)
    If GameLabel2 = " -1" Then GameLabel2 = GameLabel(15)
    Dim AutoScroll As String
    AutoScroll = Replace(Join(Filter(levelcontent2, Chr(34) & "autoavance" & Chr(34)), ""), Chr(34) & "autoavance" & Chr(34) & ": ", "")
    If AutoScroll = " 0" Then AutoScroll = GameLabel(15)
    If AutoScroll = " 1" Then AutoScroll = ErrorText(19)
    If AutoScroll = " 2" Then AutoScroll = ErrorText(20)
    If AutoScroll = " 3" Then AutoScroll = ErrorText(21)
    Dim StageStyle, LevelTimer, IsDayNight As String
    LevelTimer = Replace(Join(Filter(levelcontent2, Chr(34) & "cronometro" & Chr(34)), ""), Chr(34) & "cronometro" & Chr(34) & ": ", "")
    IsDayNight = Replace(Join(Filter(levelcontent2, Chr(34) & "modo_noche" & Chr(34)), ""), Chr(34) & "modo_noche" & Chr(34) & ": ", "")
    If IsDayNight = " 0" Then IsDayNight = GameLabel(27)
    If IsDayNight = " 1" Then IsDayNight = GameLabel(28)
    StageStyle = Replace(Join(Filter(levelcontent2, Chr(34) & "entorno" & Chr(34)), ""), Chr(34) & "entorno" & Chr(34) & ": ", "")
    StageStyle = Replace(StageStyle, Chr(34), "")
    If StageStyle = " ground" Then StageStyle = GameLabel(16)
    If StageStyle = " underground" Then StageStyle = GameLabel(17)
    If StageStyle = " sky" Then StageStyle = GameLabel(18)
    If StageStyle = " forest" Then StageStyle = GameLabel(19)
    If StageStyle = " desert" Then StageStyle = GameLabel(20)
    If StageStyle = " castle" Then StageStyle = GameLabel(21)
    If StageStyle = " ghost" Then StageStyle = GameLabel(22)
    If StageStyle = " airship" Then StageStyle = GameLabel(23)
    If StageStyle = " underwater" Then StageStyle = GameLabel(24)
    If StageStyle = " snow" Then StageStyle = GameLabel(25)
    If StageStyle = " fall" Then StageStyle = GameLabel(26)
    MsgBox ErrorText(12) & LevelMaker & vbCrLf & ErrorText(14) & GameStyle _
    & vbCrLf & ErrorText(13) & StageStyle & " " & IsDayNight & vbCrLf & ErrorText(15) & GameLabel1 & "  " & ErrorText(16) & GameLabel2 & vbCrLf & ErrorText(17) & LevelTimer _
    & vbCrLf & ErrorText(18) & AutoScroll _
    , vbOKOnly, InfoButton.Caption
End If
End Sub
Private Sub ExtractButton_Click()
'导出关卡按钮
If List1.Text <> "" Then
FileCopy LevelFolder & "\" & List1.Text & ".swe", DesktopFolder & "\" & List1.Text & ".swe"
    ExtractButton.Caption = ErrorText(23)
    DoEvents
    Sleep (700)
    ExtractButton.Caption = ErrorText(22)
End If
End Sub
Private Sub PageButton_Click()
'加载页数
PageNumber = PageNumber + 1
If PageNumber = PageNumberMax + 1 Then PageNumber = 1
If PageNumber = 1 Then
LevelSourceUrl = "https://cloud.smmwe.ml/main/"
Else
LevelSourceUrl = "https://cloud.smmwe.ml/main/Levels%20Page%20" & CStr(PageNumber - 1) & "/"
End If
DoEvents
    PageButton.Caption = ErrorText(27) & " " & CStr(PageNumber) & "/" & CStr(PageNumberMax)
    '拉取关卡
    Debug.Print DownloadFile(LevelSourceUrl & "?filename", ConfigFolder & "\SMMWECloudLevelList.txt")
    Dim filelist As String
    Open ConfigFolder & "\SMMWECloudLevelList.txt" For Input As #1
    Line Input #1, filelist
    OnlineLevel = Split(filelist, vbLf)
    OnlineLevel = Filter(OnlineLevel, ".swe")
    onlineleveltmp = Join(OnlineLevel, vbCrLf)
    onlineleveltmp = Replace(onlineleveltmp, ".swe", "")
    OnlineLevel = Split(onlineleveltmp, vbCrLf)
    Dim tmp2 As Integer
    tmp2 = UBound(OnlineLevel) - LBound(OnlineLevel)
    Dim s As Long, I As Long
    
List1.Clear
    For I = 0 To tmp2
        List1.AddItem OnlineLevel(I)
    Next I
    Close #1
LevelCounter.Caption = CStr(List1.ListCount) & ErrorText(25)
End Sub

Private Sub OfficialButton_Click()
'官方API
    Form1.Caption = Title & " - " & OfficialButton.Caption
DeleteButton.Visible = False
InfoButton.Visible = False
RenameButton.Visible = False
DownloadButton.Visible = False
UploadButton.Visible = False
ExtractButton.Visible = False
PageButton.Visible = False
SearchButton.Visible = False
ImportButton.Visible = False
List1.Visible = False
Search.Visible = False
LevelCounter.Visible = False
ID1.Visible = True
ID2.Visible = True
ID3.Visible = True
ID4.Visible = True
IDToolTip.Visible = True
IDConfirm.Visible = True
LevelInfos.Visible = False
IDToolTip.Caption = ErrorText(41)
ID1.SetFocus
End Sub
Private Sub IDConfirm_Click()
'官方API
If ID4.Text <> "" Then
ID1.Visible = False
ID2.Visible = False
ID3.Visible = False
ID4.Visible = False
IDConfirm.Visible = False
Dim LevelParseTmp As String, RetryTimer As Integer, ExitLoop As Boolean, TmpDownloadURL As String
ExitLoop = False
LevelID = ID1.Text & "-" & ID2.Text & "-" & ID3.Text & "-" & ID4.Text
RetryTimer = 0
Dim APIKey, APIOwner, DiscordCDNURL As String
If Rnd() > 0.7 Then
APIKey = APIKey1
APIOwner = APIOwner1
Else
APIKey = APIKey2
APIOwner = APIOwner2
End If
If Locale = "zh-cn" Then
DiscordCDNURL = "https://discordcdnapi.smmwe.ml/"
Else
DiscordCDNURL = "https://cdn.discordapp.com/"
End If
Do
IDToolTip = ErrorText(26) & ErrorText(42) & CStr(RetryTimer) & ErrorText(43)
TmpDownloadURL = "https://smmwe.online/api/request?type=stage&by=id&owner=" & APIOwner & "&api_key=" & APIKey & "&id=" & LevelID
Debug.Print DownloadFile(TmpDownloadURL, ConfigFolder & "\SMMWEParseTemp.json")
Open ConfigFolder & "\SMMWEParseTemp.json" For Input As #12
Line Input #12, LevelParseTmp
Close #12
Debug.Print TmpDownloadURL
If Join(Filter(Split(LevelParseTmp, ","), "error_type"), "") = "" Then ExitLoop = True
Kill ConfigFolder & "\SMMWEParseTemp.json"
DeleteUrlCacheEntry (TmpDownloadURL)
Sleep (500)
DoEvents
RetryTimer = RetryTimer + 1
Loop Until ExitLoop = True

Dim Clears As String, Fails As String
Clears = Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "victorias" & Chr(34)), ""), Chr(34), ""), "victorias:", "")
Fails = Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "muertes" & Chr(34)), ""), Chr(34), ""), "muertes:", "")
LevelInfos.Visible = True
LevelTempName = Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "name" & Chr(34)), ""), Chr(34), ""), "result:{name:", "")
LevelInfos.Caption = ErrorText(44) & LevelTempName
LevelInfos.Caption = LevelInfos.Caption & vbCrLf & ErrorText(12) & Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "author" & Chr(34)), ""), Chr(34), ""), "author:", "")
LevelInfos.Caption = LevelInfos.Caption & vbCrLf & ErrorText(47) & Clears
LevelInfos.Caption = LevelInfos.Caption & ErrorText(48) & Fails
ClearRate = CInt(Clears) + CInt(Fails)
ClearRate = CInt(Clears) / ClearRate
ClearRate = ClearRate * 100
LevelInfos.Caption = LevelInfos.Caption & ErrorText(49) & CStr(Round(ClearRate, 2)) & "%"
LevelInfos.Caption = LevelInfos.Caption & vbCrLf & vbCrLf & ErrorText(45) & Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "likes" & Chr(34)), ""), Chr(34), ""), "likes:", "")
LevelInfos.Caption = LevelInfos.Caption & ErrorText(46) & Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "dislikes" & Chr(34)), ""), Chr(34), ""), "dislikes:", "")
Dim GameStyle2, GameLabel3 As String
GameStyle2 = Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "apariencia" & Chr(34)), ""), Chr(34), ""), "apariencia:", "")
GameLabel3 = Replace(Replace(Join(Filter(Split(LevelParseTmp, ","), Chr(34) & "etiquetas" & Chr(34)), ""), Chr(34), ""), "etiquetas:", "")
    If GameStyle2 = " 0" Then GameStyle2 = "SMB1"
    If GameStyle2 = " 1" Then GameStyle2 = "SMB3"
    If GameStyle2 = " 2" Then GameStyle2 = "SMW"
    If GameStyle2 = " 3" Then GameStyle2 = "NSMBU"
LevelInfos.Caption = LevelInfos.Caption & vbCrLf & ErrorText(14) & GameStyle2 & "  " & ErrorText(15) & GameLabel3
Download2Button.Visible = True
    Download2Button.Caption = ErrorText(4)
End If
End Sub
Private Sub Download2Button_Click()
If CheckFileExists(ConfigFolder & "\SMMWEParseTemp.json") = True Then Kill ConfigFolder & "\SMMWEParseTemp.json"
    Download2Button.Caption = ErrorText(5)
    Dim ExitLoop2 As Boolean, RetryTimer2 As Integer, TmpDownloadUrl2 As String
    RetryTimer2 = 0
    ExitLoop2 = False
    TmpDownloadUrl2 = "https://smmwe.online/api/request?type=stage&by=file&owner=530177024614989824&api_key=9311a9ef9130d88cb6d1620107e5932561f881ae&id=" & LevelID
    DoEvents
Do
Download2Button.Caption = ErrorText(42) & CStr(RetryTimer2) & ErrorText(43)
Debug.Print DownloadFile(TmpDownloadUrl2, ConfigFolder & "\SMMWEParseTemp.json")
Dim LevelParseTmp2 As String
Open ConfigFolder & "\SMMWEParseTemp.json" For Input As #13
Line Input #13, LevelParseTmp2
Close #13
If Join(Filter(Split(LevelParseTmp, ","), "error_type"), "") = "" Then ExitLoop2 = True
Kill ConfigFolder & "\SMMWEParseTemp.json"
DeleteUrlCacheEntry (TmpDownloadUrl2)
Sleep (500)
DoEvents
RetryTimer2 = RetryTimer2 + 1
Loop Until ExitLoop2 = True

LevelParseTmp2 = Replace(Replace(Replace(LevelParseTmp2, Chr(34), ""), "{data:", ""), "}", "")
Open LevelFolder & "\" & LevelTempName & ".swe" For Output As #14
Print #14, LevelParseTmp2;
Close #14
    Download2Button.Caption = ErrorText(6)
End Sub
Private Sub ID1_Change()
ID1.Text = UCase(ID1.Text)
ID1.SelStart = Len(ID1.Text)
If Len(ID1.Text) = 4 Then ID2.SetFocus
End Sub
Private Sub ID2_Change()
ID2.Text = UCase(ID2.Text)
ID2.SelStart = Len(ID2.Text)
If Len(ID2.Text) = 4 Then ID3.SetFocus
End Sub
Private Sub ID3_Change()
ID3.Text = UCase(ID3.Text)
ID3.SelStart = Len(ID3.Text)
If Len(ID3.Text) = 4 Then ID4.SetFocus
End Sub
Private Sub ID4_Change()
ID4.Text = UCase(ID4.Text)
ID4.SelStart = Len(ID4.Text)
If Len(ID4.Text) = 4 Then IDConfirm.SetFocus
End Sub
