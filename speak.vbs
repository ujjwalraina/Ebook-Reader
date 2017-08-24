Dim Speak
const ForReading = 1
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\sony\Desktop\TheReunion.txt",ForReading)
strFileText = objFileToRead.ReadAll()
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak strFileText
objFileToRead.Close
Set objFileToRead = Nothing