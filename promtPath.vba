Sub promptPath()
    'prompt user for path, save to registry. Use registry value for default path.
    Dim path As String
    Dim defaultPath As String
    
    defaultPath = GetSetting("saveAtmtMacro", "pathPrompt", "path")

    path = InputBox(prompt:="Please enter path to save", default:=defaultPath)
    
    If path <> "" Then
        SaveSetting "saveAtmtMacro", "pathPrompt", "path", path
    End If
    
End Sub
