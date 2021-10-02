Attribute VB_Name = "XMLHTTP"
'Visual Basic 6 XMLHTTP Script
'https://www.jb51.net/article/53060.htm
Public Enum DataEnum
  ResponseText = 1
  ResponseBody = 2
End Enum
 
Public Function GetData(ByVal Url As String, ByVal DataStic As DataEnum) As Variant
  
  On Error GoTo ERR:
  Dim XMLHTTP As Object
  Dim DataS As String
  Dim DataB() As Byte
  
  Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
  
  XMLHTTP.Open "get", Url, True
  XMLHTTP.send
  
  While XMLHTTP.ReadyState <> 4
    DoEvents
  Wend
  Select Case DataStic
  Case ResponseText
    DataS = XMLHTTP.ResponseText
    GetData = DataS
  Case ResponseBody
    DataB = XMLHTTP.ResponseBody
    GetData = DataB
  Case ResponseBody + ResponseText
    DataS = BytesToStr(XMLHTTP.ResponseBody)
    GetData = DataS
  Case Else
    GetData = ""
  End Select
  Set XMLHTTP = Nothing
  Exit Function
ERR:
  GetData = ""
End Function
 
Public Function PostData(ByVal StrUrl As String, ByVal StrData As String, ByVal DataStic As DataEnum) As Variant
  On Error GoTo ERR:
  
  Dim XMLHTTP As Object
  Dim DataS As String
  Dim DataB() As Byte
  
  Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
  
  XMLHTTP.Open "POST", StrUrl, True
  'XMLHTTP.setRequestHeader "Content-Length", Len(PostData)
 ' XMLHTTP.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
  XMLHTTP.send (StrData)
  Do Until XMLHTTP.ReadyState = 4
    DoEvents
    Sleep (10)
  Loop
  'Select Case DataStic
 ' Case ResponseText
'    DataS = XMLHTTP.ResponseText
  '  PostData = DataS
    'DataB = XMLHTTP.ResponseBody
'    PostData = DataB
  'Case ResponseBody
'  Case ResponseBody + ResponseText
    'DataS = BytesToStr(XMLHTTP.ResponseBody)
'    PostData = DataS
  'Case Else
    'PostData = ""
  'End Select
  DataS = XMLHTTP.ResponseText
    PostData = DataS
  Set XMLHTTP = Nothing
  Exit Function
ERR:
  PostData = ""
End Function
 
Public Function GetDataSWE(ByVal Url As String) As String
  On Error GoTo ERR:
  Dim XMLHTTP As Object
  Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
  XMLHTTP.Open "GET", Url, True
  XMLHTTP.send
  While XMLHTTP.ReadyState <> 4
  Sleep 10
    DoEvents
  Wend
    GetDataSWE = XMLHTTP.ResponseText
  Set XMLHTTP = Nothing
  Exit Function
ERR:
  GetDataSWE = ""
End Function
Public Function PostDataSWE(ByVal StrUrl As String, ByVal StrData As String) As String
  On Error GoTo ERR:
  Dim XMLHTTP As Object
  Dim DataS As String
  Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
  XMLHTTP.Open "POST", StrUrl, True
  XMLHTTP.send (StrData)
  Do Until XMLHTTP.ReadyState = 4
    DoEvents
    Sleep (10)
  Loop
  DataS = XMLHTTP.ResponseText
    PostDataSWE = DataS
  Set XMLHTTP = Nothing
  Exit Function
ERR:
  PostDataSWE = ""
End Function
Public Function BytesToStr(ByVal vIn) As String
  strReturn = ""
  For I = 1 To LenB(vIn)
    ThisCharCode = AscB(MidB(vIn, I, 1))
    If ThisCharCode < &H80 Then
      strReturn = strReturn & Chr(ThisCharCode)
    Else
      NextCharCode = AscB(MidB(vIn, I + 1, 1))
      strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
      I = I + 1
    End If
  Next
  BytesToStr = strReturn
End Function
