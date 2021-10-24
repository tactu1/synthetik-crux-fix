' Script to edit Synthetik Save file to allow crux to drop

' Read commandline arguments for save file path
Set objMyArgs = WScript.Arguments

' if no arguments or more than 1 arguments are passed notify the user with correct help message
If objMyArgs.Count = 0 Then
	WScript.Echo "Path\Filename of Synthetik save.sav required"
ElseIf objMyArgs.Count > 1 Then
	WScript.Echo "Too many arguments, Path\Filename only"
Else
	' not checking the input, just see if one is there
	Dim strSaveFilePath
	Dim strFileText

	strSaveFilePath = objMyArgs(0)

	' We create a file read object here, because it cannot be used for writing
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Set objSaveFileRead = objFileSystem.OpenTextFile(strSaveFilePath, 1)

	' Loop through the file until the crux object drop chance modifier is found
	Do Until objSaveFileRead.AtEndOfStream
    	Dim strLine
    	strLine = objSaveFileRead.ReadLine

    	' Do nothing until the crux object drop chance modifier is found. Then update drop change
	    If InStr(strLine,"idropchange 149=""-") <> 0 Then
	        strLine = Replace(strLine,strLine,"idropchange 149=""0.000000""")
	        ' debug output
	        ' WScript.Echo strLine
	    End If
   	    ' Recreate file structure with carriage returns since we read it in line by line
	    strFileText = strFileText + strLine + vbCrLf
	Loop
	objSaveFileRead.Close

	' Create file write object and regenerate the entire file with modified contents
	Set objSaveFileWrite = objFileSystem.OpenTextFile(strSaveFilePath, 2)
	objSaveFileWrite.WriteLine strFileText
	objSaveFileWrite.Close

	' debug output
	' WScript.Echo Left(strFileText, 30)

End If