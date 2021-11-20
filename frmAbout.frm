VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   Caption         =   "ABOUT"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9825
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   9825
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11245
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmAbout.frx":2AFA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = ConstStr(28) & " " & ConstStr(0) & " " & AppVersion
RichTextBox1.FileName = App.Path & "\Assets\about-" & Locale & ".rtf"
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
