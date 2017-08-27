Dim Speak, Path
Path = "string"
Path = "C:\Users\sony\Desktop\TheReunion.txt"
const ForReading = 1
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(Path,ForReading)
strFileText = objFileToRead.ReadAll()
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak strFileText
objFileToRead.Close
Set objFileToRead = Nothing