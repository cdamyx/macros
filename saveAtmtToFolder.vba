
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

Sub createFullPathWithFile(fileName, Atmt, fullPath, fullPathWithFile)

    fileName = Atmt.fileName
    fullPathWithFile = fullPath + fileName

End Sub

Sub getExtension(fileName, plainName, extension)

    splitArray = Split(fileName, ".")
    plainName = LCase(splitArray(LBound(splitArray)))
    extension = LCase(splitArray(UBound(splitArray)))

End Sub

Sub checkIfExists(fullPathWithFile, fileExistence, fullPath, fileName, plainName, extension)
    'handy recursive function for future use
    fileExistence = Dir(fullPathWithFile)

    If fileExistence <> "" Then
        'first rename works great, second rename gives error "stack out of space". Clean this up
        fileName = plainName & Format(Date, "mmddyy") & "." & extension
        fullPathWithFile = fullPath & fileName
        
        checkIfExists fullPathWithFile, fileExistence, fullPath, fileName, plainName, extension
    End If

End Sub

Sub saveAtmtToFolder()

    Dim primaryFolder As MAPIFolder
    Dim completedFolder As MAPIFolder
    Dim usrPath As String
    Dim fullPath As String
    Dim fullPathWithFile As String
    Dim Item As MailItem
    Dim Atmt As Attachment
    Dim fileName As String
    Dim plainName As String
    Dim extension As String
    Dim fileExistence As String

    'set folders to run this program on - would be nice to let user choose this too
    Set primaryFolder = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("EOM rptg")
    Set completedFolder = GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("EOM rptg").Folders("COMPLETED")
    
    'add in: log file for items moved

    'prompt user for save path
    promptPath usrPath

    'if backslash not present at end of path, add it
    checkBackslash usrPath, fullPath
    
    'add in: if error (i.e. path does not exist) goTo message box
    
    'loop through each email in EOM rptg folder, then through each attachment of the current email
    For i = primaryFolder.Items.Count To 1 Step -1
        Set Item = primaryFolder.Items(i)
        For Each Atmt In Item.Attachments
    
            'makes path + filename text string
            createFullPathWithFile fileName, Atmt, fullPath, fullPathWithFile
            
            'separates and saves filename and extension
            getExtension fileName, plainName, extension
            
            'if file already exists, rename it with date
            checkIfExists fullPathWithFile, fileExistence, fullPath, fileName, plainName, extension
            
            'save file
            If extension = "xlsx" Then
                'excelCount = excelCount + 1
                Atmt.SaveAsFile fullPathWithFile
            ElseIf extension = "csv" Then
                'csvCount = csvCount + 1
                Atmt.SaveAsFile fullPathWithFile
            ElseIf extension = "pdf" Then
                'pdfCount = pdfCount + 1
                Atmt.SaveAsFile fullPathWithFile
            Else
                'add in log: could not print attachment
            End If
            
        
        Next
    'move email to completed folder
    Item.Move completedFolder
    Next

    Set primaryFolder = Nothing

End Sub

