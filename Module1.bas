Attribute VB_Name = "Module1"
Public LevelFolder As String
Public Locale As String
Public ConfigFolder As String
Public DesktopFolder As String
Public LevelSourceUrl As String
Public RenameError As String
Public PageNumber As Integer
Public PageNumberMax As Integer
Public Title As String
Public Version As String
Public ErrorText(35) As String
Public GameLabel(28) As String


Public Function CheckFileExists(FilePath As String) As Boolean
    On Error GoTo Err
    If Len(FilePath) < 2 Then CheckFileExists = False: Exit Function
            If Dir$(FilePath, vbAllFileAttrib) <> vbNullString Then CheckFileExists = True
    Exit Function
Err:
    CheckFileExists = False
End Function
'Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
'    On Error GoTo Err
'    Dim JSON As Object
'    Set JSON = CreateObject("MSScriptControl.ScriptControl")
'    JSON.Language = "JScript"
'    JSONParse = JSON.eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
'    Set JSON = Nothing
'    Exit Function
'Err:
'    JSONParse = JSONError
'End Function
