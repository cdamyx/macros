
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
        fullPath = usrPath + "\"
    Else
        fullPath = usrPath
    End If


End Sub

Sub createFullPath(fileName, Atmt, fullPath)

    fileName = Atmt.fileName
    fullPath = fullPath + fileName

End Sub

Sub saveAtmtToFolder()

    Dim primaryFolder As MAPIFolder
    Dim completedFolder As MAPIFolder
    Dim usrPath As String
    Dim fullPath As String
    Dim Item As MailItem
    Dim Atmt As Attachment
    Dim fileName As String
    Dim extension As String

    Set primaryFolder = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("EOM rptg")
    Set completedFolder = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("EOM rptg").Folders("COMPLETED")
    
    'log file for items moved

    promptPath usrPath

    checkBackslash usrPath, fullPath
    
    'if error (i.e. path does not exist) goTo message box
    
    For i = primaryFolder.Items.Count To 1 Step -1
        Set Item = primaryFolder.Items(i)
        'Nested loop iterates through all of the attachments in a single email
        For Each Atmt In Item.Attachments
    
            createFullPath fileName, Atmt, fullPath
            
            splitArray = Split(fileName, ".")
            extension = LCase(splitArray(UBound(splitArray)))
            
            'need to check if file exists
            
            If extension = "xlsx" Then
                'excelCount = excelCount + 1
                Atmt.SaveAsFile fullPath
            ElseIf extension = "csv" Then
                'csvCount = csvCount + 1
                Atmt.SaveAsFile fullPath
            ElseIf extension = "pdf" Then
                'pdfCount = pdfCount + 1
                Atmt.SaveAsFile fullPath
            ElseIf extension = "txt" Then
                'txtCount = txtCount + 1
                'get rid of txt elseif after testing is finished
                Atmt.SaveAsFile fullPath
            Else
                'log: could not print attachment
                'MsgBox ("error, not good ext")
            End If
        
        Next
    Item.Move completedFolder
    Next
    'primaryFolder.Items(1).Attachments(1).SaveAsFile fullPath

    Set primaryFolder = Nothing

End Sub

