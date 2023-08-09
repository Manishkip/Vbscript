Option Explicit

Dim folderPath, latestFile, latestDate, fso, folder, file, fileNameWithoutExtension, fileDate, regex, matches, fileDateStr
folderPath = "Test Folder"  ' Replace with your folder path

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(folderPath)

For Each file In folder.Files
    ' Extract date from file name using regular expression
    ' msgbox file.name 
    Set regex = New RegExp
    regex.Pattern = "(\d{4}-\d{2}-\d{2})"
    Set matches = regex.Execute(file.Name)
    
    If matches.Count > 0 Then
        fileDateStr = matches(0).SubMatches(0)
        fileDate = CDate(fileDateStr)

        ' msgbox fileDate 
        
        ' Compare dates to find the latest file
        If latestFile = "" Or fileDate > latestDate Then
            latestDate = fileDate
            latestFile = file.Path
        End If
    End If
Next

If latestFile <> "" Then
    WScript.Echo "Latest file: " & latestFile
    WScript.Echo "Date in the file name: " & latestDate
Else
    WScript.Echo "No files found in the folder."
End If
