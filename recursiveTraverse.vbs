'uses recursive script to traverse all subfolders and list each file

Dim fso
' create a global copy of the filesystem object
Set fso = CreateObject("Scripting.FileSystemObject")

'set the path of the root directory to be traversed
Dim path
path = "C:\Users\Phelan\Desktop\source folder"

Sub RootFolder()
'Iterate through and display all the files in the root folder
	ShowName(path)
End Sub

RootFolder

' Call the RecurseFolders routine with name of function to be performed
' Takes one argument - in this case, the Path of the folder to be searched
RecurseFolders path, "ShowName"

' echo the job is completed
WScript.Echo "Completed!"

Sub RecurseFolders(sPath, funcName)
	Dim folder
	
	'traverse subfolders and execute whatever function is passed into the RescurseFolders sub
	 With fso.GetFolder(sPath)
	   if .SubFolders.Count > 0 Then
	     For each folder in .SubFolders
	        ' Perform function's operation
	        Execute funcName & " " & chr(34) & folder.Path & chr(34)
	
	        ' Recurse to check for further subfolders
	        RecurseFolders folder.Path, funcName
	     Next
	   End If
	 End With

End Sub

Sub CopyAction(folPath)
'copies all of the files and folders at the level of recursion.
	Dim fil, folder

	For Each fil In fso.GetFolder(folPath).Files
		
	next


End Sub

Sub ShowName(folPath)
Dim fil

  ' go thru each file in the folder
  For Each fil In fso.GetFolder(folPath).Files

  	'echo the name of the file
  	WScript.Echo fil
  Next
end Sub
