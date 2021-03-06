Dim objShell 'shell object
Dim strCommand 'string to store results of copy execution
Dim strSource 'source for copy
Dim strDestination 'destination for copy
Dim chrQuotes 'use of quotes character for proper command formatting

chrQuotes = Chr(34)
strSource = chrQuotes & "C:\Users\Phelan\Desktop\source folder\*.*" & chrQuotes
strDestination = chrQuotes & "C:\Users\Phelan\Desktop\destination_folder\" & chrQuotes

'windows shell object to run xcopy	
Set objShell = WScript.CreateObject("Wscript.Shell")
'wshshell.run("cmd /c xcopy C:\Users\Phelan\Desktop\source folder\*.* C:\Users\Phelan\Desktop\destination_folder\ /d /s /e /q /h /r /o /x /y /c"),1,True

'correct xcopy syntax:
'xcopy "C:\Users\Phelan\Desktop\source folder\*.*" "C:\Users\Phelan\Desktop\destination_folder\" /c /e /j /q /s /y /z /d

'strCommand = objShell.Run("Xcopy ""C:\Users\Phelan\Desktop\source folder\*.*"" ""C:\Users\Phelan\Desktop\destination_folder\"" /c /e /j /q /s /y /z /d", 0, True) 


strCommand = objShell.Run("Xcopy " & strSource & " " & strDestination & " /c /e /j /q /s /y /z /d", 0, True)
 

If strCommand <> 0 Then 
	MsgBox "File Copy Error: " & strCommand 
Else 
	MsgBox "Done Copying" 
End If

'xcopy syntax is used as follows: 
'/c = continue copying if error occurs, 
'/e = copy directories and subdirectories, 
'/j = copy using unbuffered i/o, 
'/q = doesn't display filenames when copying, 
'/y = supresses overwrite confirmation, 
'/z = network restart mode, 
'/d = copies only those files whose source time is newer than the destination time