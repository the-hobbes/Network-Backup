'''tortiseBackup.vbs
'''A script used to backup all of the items in a user's Documents and Desktop folder to a remote location, if the timestamp on the local
'''location is newer than that of the network location.

Class BackupObject
'''Class used to create a backup object containing all of the references necessary to perform a backup.
'''Maps the destination drive if necessary and creates the backup folder in that drive if necessary
'''Determines the source user folder as well

	'''Global Properties declared here
	Public userName
	Public homeFolder
	Public destinationDrive
	Public destinationFolder
	Public fso
	Public networkObject
	Public desktopSource
	Public documentsSource
	
	Private Sub Class_Initialize
  	'''default constructor for the Class
		set fso = CreateObject("Scripting.FileSystemObject")
		Set networkObject = CreateObject("wscript.network")		
		userName = GetUser()
		homeFolder = GetSource()
		destinationDrive = GetTarget()
		destinationFolder = GetFolder()	
		desktopSource = homeFolder & "\Desktop"
		documentsSource = homeFolder & "\Documents"
		
	End Sub

	Function GetUser()
	'''A function to retrieve and return the username of the current user who is logged in.
	''' Returns the username of the current user
		GetUser = networkObject.userName
	End Function 

	Function GetSource()
	''' A function to retrieve the home folder of the user who is logged in.
	''' Returns the home folder
		Set oShell = CreateObject("WScript.Shell")
		strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
		GetSource = strHomeFolder
	End Function

	Function GetTarget()
	'''Function used to check to see if the target location for the backup is available, and make it so if it is not.
	'''Target is if the format: \\ipaddress\username, which is then mapped to the desired drive letter if it isn't already
	''' Returns the local path to the mapped drive
		Const TARGET_IP = "192.168.1.4"
		Const TARGET_DRIVE_LETTER = "B"
		
		If IsDriveMapped(TARGET_DRIVE_LETTER) Then
		'actions to take if targetDriveLetter is already mapped
			GetTarget = TARGET_DRIVE_LETTER & ":\"
		Else
		'actions to take if targetDriveLetter is not already mapped
			networkObject.MapNetworkDrive TARGET_DRIVE_LETTER & ":", "\\" & TARGET_IP & "\" & userName
			GetTarget = TARGET_DRIVE_LETTER & ":\"
		end If
	End Function
	
	function IsDriveMapped (byval drive)
	'''Function used to check to see if the drive letter that is passed into it is mapped.
	'''use only the letter, not :, and make sure it's uppercase
	''' Returns false or 3 for drive mapped
    	drive = ucase(left(drive,1))

    	' assume it's not mapped
    	IsDriveMapped = False

    	' if no such drive, return False right now
    	if not fso.DriveExists(drive) then exit function

    	' get Drive object and check its type: 3 = mapped
    	isDriveMapped = (fso.GetDrive(drive).driveType = 3)
	end function

	Function GetFolder()
	'''Function used to check for the existance of the Backup folder in the mapped drive, and create it if it does not exist.
	''' Returns the full path of the destination folder for the backup
		Dim folder 
		folder = "Backup_" & userName
		
		Dim folderPath
		folderPath = destinationDrive & folder
		
		Dim newFolder
		
		If Not fso.FolderExists(folderPath) Then
			Set newFolder = fso.CreateFolder(folderPath) 
		End If 	
		
		GetFolder = folderPath
	End Function

	'''check each folder in the backup source against the backup destination.
		'''if a file or folder is newer, than copy it to the destination

End Class

'Create instance of backup object to set up backup environment
Dim backObjInstance
Set backObjInstance = new BackupObject
Dim source
Dim destination

'set source and destination folder variables
source = backObjInstance.homeFolder
destination = backObjInstance.destinationFolder

'call traverse folder
TraverseFolder source, destination

Function TraverseFolder(sourceFolder, destinationFolder)
' ********
' function to traverse the entirety of the source folder
' takes source and destination folder paths as arguments
' copies folders to destination when they do not exist at destination
' ********
	
	'run through the root folder
	RootFolder sourceFolder, destinationFolder
	
	'traverse subfolders recursively
	' ...
		'get a list of folders at this level
		'if a folder doesn't exist, create/copy it. otherwise, ignore it

End Function

Function RootFolder(sourceFolder, destinationFolder)
' ********
' traverse the root of the source folder and copy newer (and those that don't exist at the destination) folders to destination folder
' ********
	'WScript.Echo sourceFolder 
	'WScript.Echo destinationFolder
	Dim fileSys
	Dim source
	Dim destination
	Dim sourceDate
	Dim destDate
	Dim workingSource
	Dim workingDest
	
	set fileSys = CreateObject("Scripting.FileSystemObject")
	set source = fileSys.GetFolder(sourceFolder)
	set destination = fileSys.GetFolder(destinationFolder)
	
	'get the length of the path for the root source folder and assign it to length
	path = source.Path
	Dim pathLength
	pathLength = Len(path)
	
	'run through each subfolder in the root source
	For Each subFolder In source.SubFolders
		'get a string of path the subfolder, without the root preceeding it.
		subPath = subFolder.path
		subString = right(subPath, Len(subPath) - pathLength) & "\"
		'WScript.Echo subString
		
		'set the working variable to hold the temporary source folder subFolder. Also set a temporary path for the destination to be checked for existance
		Set workingSource = fileSys.GetFolder(subFolder)
		tempDestPath = destination & subString
		
		'check to see if the folder exists at the destination
		If fileSys.FolderExists(tempDestPath) Then
			Set workingDest = fileSys.GetFolder(destination & subString)
			'if it does exist at the destination, check the date modified. If it is newer than the source, dont copy. otherwise, copy
			sourceDate = workingSource.DateLastModified
			destDate = workingDest.DateLastModified
			
			If(destDate < sourceDate) Then
				WScript.Echo "Source is newer. Source is copied to target destination"
				filesys.CopyFolder workingSource & "\",workingDest &"\",True 
			ElseIf(sourceDate < destDate)then
				WScript.Echo "Destination is newer. No copying occurs"
			end If
		'otherwise, the folder doesn't exist at the destination and must be copied to there. 
		Else
			Dim newFolderPath
			Dim newFolder
			newFolderPath = destination & subString
			
			set newFolder = filesys.CreateFolder(newfolderpath)
			'filesys.CopyFolder workingSource,newFolderPath,True
			
			Set objFiles = workingSource.Files
			
			If objFiles.Count <> 0 Then
				On Error Resume Next
				fileSys.CopyFile workingSource & "\" & "*.*", newFolderPath, True
			End If
			
			WScript.Echo "Copied " & newFolder
		End If
	Next
End Function

Function CopyFolders()
End Function

Function CopyFiles()
End Function
