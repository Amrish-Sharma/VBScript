'setting up the working shell for VBscript
set WshShell = WScript.CreateObject("WScript.Shell")
'running the command ipconfig and storing the value in IpInfo.txt
WshShell.run "cmd /c ipconfig > IpInfo.txt"
'This will put the vbscript to sleep for 3 seconds, 1000ms eq to 1 second
WScript.Sleep 3000
'variable to store the filname
filename = "IpInfo.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)
'to read all the content of the file and storing it in the variable
strText = f.ReadAll
f.Close
'to output the content of the file
Wscript.Echo strText