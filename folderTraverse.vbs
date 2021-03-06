Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = wscript.CreateObject("wscript.network")
SOURCE_FOLDER = "C:\test_src\"
DESTINATION_FOLDER = "C:\test_dst\"

ReplicateFolders objFSO, SOURCE_FOLDER, DESTINATION_FOLDER
'******************
' Sub ReplicateFolders
'
' This procedure replicates between the source and the destination
' directories at the folder level. A recursive search is done
' between the 2 directories and folders compared. If a folder exists on the 
' destination but not on the source, the source is deleted
'******************
Sub ReplicateFolders (objFSO, strSourceFolderPath, strDestinationFolderPath)
	Dim aFolderArraySource ' Declares an array for folder source
	Dim aFolderArrayDestination ' Declares an array for folder destination
	Dim FolderListSource ' declares a var for folder list source
	Dim FolderListDestination ' declares a var for folder list destination
	Dim oFolderSource
	Dim oFolderDestination
	Dim bSourceExists
	Dim bDestinationExists
	
	On Error Resume Next
		Set aFolderArraySource = objFSO.GetFolder(strSourcefolderpath)
		Set aFolderArrayDestination = objFSO.GetFolder(strDestinationfolderpath)
		Set FolderListSource = aFolderArraySource.SubFolders
		Set FolderListDestination = aFolderArrayDestination.SubFolders
	
	ReplicateFiles objFSO, strSourcefolderpath, strDestinationfolderpath
	
	For Each oFolderDestination in FolderListDestination
		bSourceExists = 0
		For Each oFolderSource in FolderListSource
			If oFolderDestination.Name = oFolderSource.Name then
				bSourceExists = 1
				Exit For
			End If
		Next
	
		If bSourceExists = 0 then
			objFSO.DeleteFolder strDestinationfolderpath & "\" & _
			oFolderDestination.Name, true
		End if
	Next
End Sub
'******************
' Sub ReplicateFiles
'
' This procedure replicates between the source and the destination
' directories at the file level.
' If a particular file on the destination directory
' does not exist on the source at any level then the destination file
' is removed from the destination directory.
'
'******************

Sub ReplicateFiles (objFSO, strSourcefolderpath, strDestinationfolderpath)

Dim aFileArraySource
Dim aFileArrayDestination
Dim FileListSource
Dim FileListDestination
Dim oFileSource
Dim oFileDestination
Dim bSourceExists
Dim bDestinationExists
On Error Resume Next
Set aFileArraySource = objFSO.GetFolder(strSourcefolderpath)
Set aFileArrayDestination = objFSO.GetFolder(strDestinationfolderpath)
Set FileListSource = aFileArraySource.files
Set FileListDestination = aFileArrayDestination.files
For each oFileDestination in FileListDestination
bSourceExists = 0
For each oFileSource in FileListSource
If oFileDestination.Name = oFileSource.Name then
If oFileDestination.DateLastModified = oFileSource.DateLastModified then
bSourceExists = 1
Exit For
End If
End If
Next
If bSourceExists = 0 then
MsgBox strDestinationfolderpath & "\" & oFileDestination.Name
objFSO.DeleteFile strDestinationfolderpath & "\" & _
oFileDestination.Name,true
End If
Next
End Sub