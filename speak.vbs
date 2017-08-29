Dim Speak, Path, key
Path = "string"
Path = "C:\Users\sony\Desktop\TheReunion.txt"
const ForReading = 1
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(Path,ForReading)
strFileText = objFileToRead.ReadAll()
Set Speak=CreateObject("sapi.spvoice")
Do  
	Speak.Speak strFileText,1
	key = InputBox("Enter p for pause , r for resume & t for termination.")
	if key = "p" Then 
		Speak.Pause
	ElseIf key = "r" Then 
		Speak.Resume  
	ElseIf key = "t" Then
		WScript.Quit
	End If
Loop
objFileToRead.Close
Set objFileToRead = Nothing
