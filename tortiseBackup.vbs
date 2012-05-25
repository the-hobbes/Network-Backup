' ********
' tortiseBackup.vbs
' A script used to backup all of the items in a user's Documents and Desktop folder to a remote location, if the timestamp on the local
' location is newer than that of the network location.
' ********

Class BackupObject
' ********
' Class used to create a backup object containing all of the references necessary to perform a backup.
' Maps the destination drive if necessary and creates the backup folder in that drive if necessary
' Determines the source user folder as well
' ********

	'Global Properties declared here
	Public userName
	Public homeFolder
	Public destinationDrive
	Public destinationFolder
	Public fso
	Public networkObject
	'Public desktopSource
	'Public documentsSource
	
	Private Sub Class_Initialize
  	' ********
	' default constructor for the Class
	' ********
		set fso = CreateObject("Scripting.FileSystemObject")
		Set networkObject = CreateObject("wscript.network")		
		userName = GetUser()
		homeFolder = GetSource()
		destinationDrive = GetTarget()
		destinationFolder = GetFolder()	
		'desktopSource = homeFolder & "\Desktop"
		'documentsSource = homeFolder & "\Documents"
		
	End Sub

	Function GetUser()
	' ********
	' A function to retrieve and return the username of the current user who is logged in.
	' Returns the username of the current user
	' ********	
		GetUser = networkObject.userName
	End Function 

	Function GetSource()
	' ********
	' A function to retrieve the home folder of the user who is logged in.
	' Returns the home folder
	' ********	
		Set oShell = CreateObject("WScript.Shell")
		strHomeFolder = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
		GetSource = strHomeFolder
	End Function

	Function GetTarget()
<<<<<<< HEAD
	'''Function used to check to see if the target location for the backup is available, and make it so if it is not.
	'''Target is if the format: \\ipaddress\username, which is then mapped to the desired drive letter if it isn't already
	''' Returns the local path to the mapped drive
		Const TARGET_IP = "192.168.1.3"
=======
	' ********
	' Function used to check to see if the target location for the backup is available, and make it so if it is not.
	' Target is if the format: \\ipaddress\username, which is then mapped to the desired drive letter if it isn't already
	' Returns the local path to the mapped drive
	' ********
		Const TARGET_IP = "192.168.1.4"
>>>>>>> 2a397a22fdbab5039a16cddc65e877a0d46de3af
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
	' ********
	' Function used to check to see if the drive letter that is passed into it is mapped.
	' use only the letter, not :, and make sure it's uppercase
	' Returns false or 3 for drive mapped
	' ********
    		drive = ucase(left(drive,1))

	    	' assume it's not mapped
    		IsDriveMapped = False

    		' if no such drive, return False right now
    		if not fso.DriveExists(drive) then exit function

    		' get Drive object and check its type: 3 = mapped
    		isDriveMapped = (fso.GetDrive(drive).driveType = 3)
	end function

	Function GetFolder()
	' ********
	' Function used to check for the existance of the Backup folder in the mapped drive, and create it if it does not exist.
	' Returns the full path of the destination folder for the backup
	' ********
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
End Class

Sub DoBackup
' ********
' Subroutine to perform the backup. Retrieves the source and destination from the backup object, then uses
' xcopy to perform the necessary copying procedures.
' xcopy usage is as follows:
'	/c = continue copying if error occurs, 
'	/e = copy directories and subdirectories, 
'	/j = copy using unbuffered i/o, 
'	/q = doesn't display filenames when copying, 
'	/y = supresses overwrite confirmation, 
'	/z = network restart mode, 
'	/d = copies only those files whose source time is newer than the destination time
' ********
	'set debugging flag
	Dim debug
	debug = False
	
	'Create instance of backup object to set up backup environment
	Dim backObjInstance
	Set backObjInstance = new BackupObject
	
	'shell object to run xcopy from command shell
	Dim objShell 
	Set objShell = WScript.CreateObject("Wscript.Shell")
	'string to store results of copy execution
	Dim strCommand 
	'use of quotes character for proper command formatting
	Dim chrQuotes
	chrQuotes = Chr(34)
	
	'set source and destination folder variables, add a slash and the quotes through the use of chrQuotes
	Dim strSource
	Dim strDestination
	strSource = chrQuotes & backObjInstance.homeFolder & "\*.*" & chrQuotes
	strDestination = chrQuotes & backObjInstance.destinationFolder & "\" & chrQuotes
	
	'run the command to copy and gather any results
	strCommand = objShell.Run("Xcopy " & strSource & " " & strDestination & " /c /e /j /q /y /z /d", 0, True)
	
	' ******** 
	' used for debugging purposes to print errors or successes
	' ********
	If debug = true Then
		WScript.Echo strSource
		WScript.Echo strDestination
		
		'depending on the results of the command after it has run, display errors or success
		If strCommand <> 0 Then 
			MsgBox "File Copy Error: " & strCommand 
		Else 
			MsgBox "Done Copying" 
		End If
	End If
	
End Sub

'kickoff script
DoBackup
