Attribute VB_Name = "Module1"
Public LevelFolder As String
Public Locale As String
Public ConfigFolder As String
Public DesktopFolder As String
Public LevelSourceUrl As String
Public LevelID As String
Public LevelTempName As String
Public RenameError As String
Public PageNumber As Integer
Public PageNumberMax As Integer
Public Title As String
Public Version As String
Public ErrorText(49) As String
Public GameLabel(28) As String
Public Const APIOwner1 = "728005293665026189"
Public Const APIOwner2 = "530177024614989824"
Public Const APIKey1 = "882fa39bc11c46db98c9f9a46fe837ae72ffba24"
Public Const APIKey2 = "9311a9ef9130d88cb6d1620107e5932561f881ae"
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
