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



'If DateDiff("n", date1, date2) < 0 Then
'    'date1 is more recent than date 2, comparison by "minute" (n)
'    WScript.Echo "file 1 (date: " & date1 & ") is more recent than file2 (date: " & date2 & ")"
'Else
'	WScript.Echo "shiiiit"
'End If