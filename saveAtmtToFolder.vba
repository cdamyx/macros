
Sub promptPath(usrPath)
    'prompt user for path, save to registry. Use registry value for default path.
    'if path does not end with backslash, add it on there
    Dim defaultPath As String
    
    defaultPath = GetSetting("saveAtmtMacro", "pathPrompt", "path")

    usrPath = InputBox(prompt:="Please enter path to save", Default:=defaultPath)
    
    If usrPath <> "" Then
        SaveSetting "saveAtmtMacro", "pathPrompt", "path", usrPath
    End If
    
End Sub

Sub checkBackslash(usrPath, fullPath)
    
    If Right(usrPath, 1) <> "\" Then
        fullPath = usrPath + "\" + "vbaText.txt"
    Else
        fullPath = usrPath + "vbaText.txt"
    End If


End Sub

Sub saveAtmtToFolder()

    Dim primaryFolder As MAPIFolder
    Dim completedFolder As MAPIFolder
    Dim usrPath As String
    Dim fullPath As String


    Set primaryFolder = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("EOM rptg")
    Set completedFolder = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("EOM rptg").Folders("COMPLETED")

    promptPath usrPath

    checkBackslash usrPath, fullPath
    
    'test save something to path
    'if error (i.e. path does not exist) goTo message box
    'fullPath = usrPath + "vbaText.txt"
    primaryFolder.Items(1).Attachments(1).SaveAsFile fullPath



End Sub
