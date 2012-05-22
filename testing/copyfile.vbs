Sub CopyFile

	Dim filesys, firstFile, secondFile, date1, date2, sourcePath, targetPath
	
	sourcePath = "C:\Users\Phelan\Desktop\source folder\New Text Document.txt"
	targetPath = "C:\Users\Phelan\Desktop\destination folder\New Text Document.txt"
	
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set firstFile = filesys.GetFile(sourcePath)
	date1 = firstFile.DateLastModified
	
	Set secondFile = filesys.GetFile(targetPath)
	date2 = secondFile.DateLastModified
	
	WScript.Echo "Source File1: " & date1
	WScript.Echo "Destination File2: " & date2
	
	If(date2 < date1) Then
		WScript.Echo "Source is newer. Source is copied to target destination"
		filesys.CopyFile sourcePath,targetPath,True 
	ElseIf(date1 < date2)then
		WScript.Echo "Destination is newer. No copying occurs"
	end If

End Sub

CopyFolder

Sub CopyFolder

	Dim filesys, firstFolder, secondFolder, date1, date2, sourcePath, targetPath

	Set filesys = CreateObject("Scripting.FileSystemObject")
	
	sourcePath = "C:\Users\Phelan\Desktop\source folder"
	targetPath = "C:\Users\Phelan\Desktop\destination folder"
	
	If filesys.FolderExists(sourcePath)then
		Set firstFolder = filesys.GetFolder(sourcePath)
		date1 = firstFolder.DateLastModified
		WScript.Echo "Source Folder Date Last Modified: " & date1
	Else
		WScript.echo "Source folder path does not exist"
	End If
	
	If filesys.FolderExists(targetPath)then
		Set secondFolder = filesys.GetFolder(targetPath)
		date2 = secondFolder.DateLastModified
		WScript.Echo "Destination Folder Date Last Modified: " & date2
	Else
		Set newfolder = filesys.CreateFolder(targetPath)
		filesys.CopyFolder sourcePath,targetPath,True
		WScript.Echo "Target folder path does not exist. Thus, it was created and the contents were copied to it."
	End If
	
	If(date2 < date1) Then
		WScript.Echo "Source is newer. Source is copied to target destination"
		filesys.CopyFolder sourcePath,targetPath,True 
	ElseIf(date1 < date2)then
		WScript.Echo "Destination is newer. No copying occurs"
	end If

End Sub


'If DateDiff("n", date1, date2) < 0 Then
'    'date1 is more recent than date 2, comparison by "minute" (n)
'    WScript.Echo "file 1 (date: " & date1 & ") is more recent than file2 (date: " & date2 & ")"
'Else
'	WScript.Echo "shiiiit"
'End If
