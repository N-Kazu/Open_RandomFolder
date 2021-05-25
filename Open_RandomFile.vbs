Option Explicit
 
Dim strTargetFolderPath,Path
Dim objShell
Dim objFileSys
Dim objFolder,objFolders
Dim intValue, intUpper, intLower
Dim strName,strT

Set objFileSys = CreateObject("Scripting.FileSystemObject")
strTargetFolderPath = objFileSys.getParentFolderName(WScript.ScriptFullName)
Set objFolder = objFileSys.GetFolder(strTargetFolderPath)

Randomize
intUpper = objFolder.Files.Count
intLower = 0
intValue = Int((intUpper * Rnd) + intLower)

For Each objFolders In objFileSys.GetFolder(strTargetFolderPath).Files
    If strT <> "" Then
        strT = strT & "," & objFolders.Name
    Else
        strT = objFolders.Name
    End If
Next
strName = Split(strT,",")

Path = strTargetFolderPath & "\" & strName(intValue)

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run Path
Set objShell = Nothing