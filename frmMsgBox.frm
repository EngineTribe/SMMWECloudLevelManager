VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
 Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
 
Private Sub Form_Activate()
Timer1.Enabled = False
    Me.BackColor = vbCyan
   SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hWnd, vbCyan, 0&, LWA_COLORKEY
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 2 Or 1
DoEvents
If MsgBoxType = "SUCCESS" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgok.png")
Label1.Caption = ConstStr(22)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Interval = 1500
Timer1.Enabled = True
End If
If MsgBoxType = "SUCCESS2" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgok.png")
Label1.Caption = ConstStr(44)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Interval = 1500
Timer1.Enabled = True
End If
If MsgBoxType = "LOADING" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgdownloading.png")
Label1.Caption = ConstStr(23)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Enabled = False
End If
If MsgBoxType = "UPLOADING" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgdownloading.png")
Label1.Caption = ConstStr(43)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Enabled = False
End If
If MsgBoxType = "LOADINGINFO" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgdownloading.png")
Label1.Caption = ConstStr(27)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Enabled = False
End If
If MsgBoxType = "DELETE" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgok.png")
Label1.Caption = ConstStr(24)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Interval = 1500
Timer1.Enabled = True
End If
If MsgBoxType = "SAVED" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgok.png")
Label1.Caption = ConstStr(31)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Interval = 1500
Timer1.Enabled = True
End If
If MsgBoxType = "RENAME" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgok.png")
Label1.Caption = ConstStr(25)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Interval = 1500
Timer1.Enabled = True
End If
If MsgBoxType = "SEARCHERR" Then
Image1.Picture = StdPictureEx.LoadPicture(App.path & "\Assets\dec-msgerr.png")
Label1.Caption = ConstStr(36)
Label1.Font.Name = "DinkieBitmap 9pxDemo"
Timer1.Interval = 1500
Timer1.Enabled = True
End If
DoEvents
End Sub

Private Sub Timer1_Timer()
Me.Hide
Unload Me
End Sub
