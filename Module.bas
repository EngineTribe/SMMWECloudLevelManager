Attribute VB_Name = "Module"
Public AppVersion As String, InternalVersion As String
Public Locale As String
Public ConfigFolder As String
Public DesktopFolder As String
Public LevelFolder As String
Public LocaleSuffix As String
Public OperateType As Integer
Public PageNum As Integer
Public OLWebIP, OLAPIIP As String
Public IsSFXEnable As Boolean
Public IsBGMEnable As Boolean
Public IsPreloadEnable As Boolean
Public ProxyDlSuffix As String
Public ConstStr() As String
Public MirrorList() As String
Public LocalizedVals() As String
Public GameLabel() As String
Public MsgBoxType As String
Public IsLoading As Boolean
Public IsSearching As Boolean
Public DownloadMethod As Integer
Public UseMirror As String
Public MirrorlistTmp() As String
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function dcWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const SYNCHRONIZE = &H100000
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function AddFontResourceEx Lib "gdi32" Alias "AddFontResourceExA" (ByVal sFileName As String, ByVal lFlags As Long, ByVal lReserved As Long) As Long
Public Declare Function RemoveFontResourceEx Lib "gdi32" Alias "RemoveFontResourceExA" (ByVal sFileName As String, ByVal lFlags As Long, ByVal lReserved As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Const FR_PRIVATE As Long = &H10
Public Sub ShowMsgBox(MsgType As String)
MsgBoxType = MsgType
frmMsgBox.Show
frmMsgBox.Top = frmMain.Top + 8500
frmMsgBox.left = frmMain.left + 12000
End Sub
Public Sub PlayMusic(MusicFileName As String)
If IsBGMEnable Then
    'Load Music
    Mix_HaltMusic                       'Stop previous
    Mix_FreeMusic music
    music = Mix_LoadMUS(App.path & "\Assets\" & MusicFileName) 'Open new music
    'Starting of music playback
    Mix_PlayMusic music, -1
    End If
End Sub
Public Sub PlaySFX(MusicFileName As String)
If IsSFXEnable Then
    'Load Music
    sfx = Mix_LoadWAV(App.path & "\Assets\" & MusicFileName)
    'Starting of music playback
    Mix_PlayChannel -1, sfx, 0
    End If
End Sub
Public Function CheckFileExists(FilePath As String) As Boolean
    On Error GoTo ERR
    If Len(FilePath) < 2 Then CheckFileExists = False: Exit Function
            If Dir$(FilePath, vbAllFileAttrib) <> vbNullString Then CheckFileExists = True
    Exit Function
ERR:
    CheckFileExists = False
End Function
