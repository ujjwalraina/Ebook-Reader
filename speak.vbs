option explicit
dim strpath, fso, strfile, strtxt, user, voice, flag

flag = 1

call init
sub init
do while len(strpath) = 0
strpath = inputbox ("Please enter the full path of txt file", "Txt to Speech")
if isempty(strpath) then
wscript.quit()
end if
loop
'strpath = "C:\Users\???\Desktop\???.txt"

set fso = createobject("scripting.filesystemobject")
on error resume next
set strfile = fso.opentextfile(strpath,1)
if err.number = 0 then
strtxt = strfile.readall()
strfile.close
call ctrl
else
wscript.echo "Error: " & err.number & vbcrlf & "Source: " & err.source & vbcrlf &_
"Description: " & err.description
err.clear
call init
end if
end sub

sub ctrl
user = msgbox("Press ""yes"" to Play, ""no"" for Pause & ""cancel"" to exit", vbyesnocancel + vbexclamation, "Txt to Speech")
select case user
case vbyes
	if flag = 1 Then
		call spk
		call ctrl
	elseif flag = 0 Then
        voice.resume
		call ctrl
    end if
case vbno
		voice.pause
		flag = 0
		call ctrl
case vbcancel
    wscript.quit
end select
end sub 

sub spk
'wscript.echo strtxt
set voice = createobject("sapi.spvoice")
voice.speak strtxt,1
flag = 0
call ctrl
end sub
