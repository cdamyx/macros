
Sub promptPath(usrPath)
    'prompt user for path, save to registry. Use registry value for default path.
    'need error checking on path, i.e., backslash has to be on end
    Dim defaultPath As String
    
    defaultPath = GetSetting("saveAtmtMacro", "pathPrompt", "path")

    usrPath = InputBox(prompt:="Please enter path to save", Default:=defaultPath)
    
    If Path <> "" Then
        SaveSetting "saveAtmtMacro", "pathPrompt", "path", usrPath
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

    'test save something to path
    fullPath = usrPath + "vbaText.txt"
    primaryFolder.Items(1).Attachments(1).SaveAsFile fullPath



End Sub