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
		Const TARGET_IP = "192.168.1.5"
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
	End Function

	'''check each folder in the backup source against the backup destination.
		'''if a file or folder is newer, than copy it to the destination

End Class

'Create instance of backup object to set up backup environment
Dim backObjInstance

'catch any errors created by the construction of the object
On Error Resume Next 
Set backObjInstance = new BackupObject

'retrieve home directory from backup object, eg. C:\users\phelan
WScript.Echo backObjInstance.desktopSource
WScript.Echo backObjInstance.documentsSource

'traverse desktop and perform backup
TraverseFolder(backObjInstance.desktopSource)
'traverse documents and perform backup
TraverseFolder(backObjInstance.documentsSource)

Sub TraverseFolder(sourceFolder)
'takes in path of the source folder and destination(backup) folder.


End Sub

'''
'set fso = CreateObject("Scripting.FileSystemObject")
'if fso.FolderExists("c:\windows") Then
'	WScript.echo "There is a folder named c:\windows"
'end If
'''
