' Function to Backup the files
Function backup(strInput)
	for each copy in cp
		FSO.CopyFile copy, "C:\Users\User Name\Desktop\backup\"&strInput&"\" 	' Change this path according to your need
	Next
	Wscript.echo "Done!"
End Function

' Function to restore the files
Function restore(strInput)
	for each copy in cp
    ' There are two paths here. Change them according to your need. It is copying from first parameter to the second.
		FSO.CopyFile "C:\Users\User Name\Desktop\backup\"&strInput&"\*.*", "C:\Users\User Name\Desktop\restore\"
	Next
	Wscript.echo "Done!"
End Function

Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim copy
' This is the array that stores the file names to be backed up.
' Edit, Delete or add according to your need
cp = array("C:\Users\User Name\Documents\Notes\Blank.jtp", "C:\Users\User Name\Documents\Notes\Graph.jtp")

strInput = InputBox("Who's there? O_o")
strInput = Lcase(strInput)

reason = InputBox("And what do you want to do?")
reason = Lcase(reason)

if strInput = "name1" OR strInput = "name2" OR strInput = "name3" Then 'Edit remove or add these names according to your need
	if reason = "backup" Then
		backup(strInput) 					' Call the backup function.
	ElseIf reason = "restore" Then
		restore(strInput) 					' Call the restore function.
	Else
		Wscript.echo "Huhh..." 				' In case of wrong option.
	End If
Else
	Wscript.Echo "I Don't work for you right now -_- "&vbNewLine&"Tell 7h1n0b1 to add you in the loop!"
End If
