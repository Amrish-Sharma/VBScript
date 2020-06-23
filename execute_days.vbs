dtmToday = Date()

dtmDayOfWeek = DatePart("w", dtmToday)

''' Select case to pickup the value of day of the week and call procedure
Select Case dtmDayOfWeek
    Case 1 
    Call Sunday()
    Case 2 
    Call Monday()
    Case 3 
    Call Tuesday()
End Select


'''Sunday procedure will execute from select case
sub Sunday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c ping -n 10 youtube.com >> ping.txt", hidden
wscript.quit
End sub

'''Monday procedure will execute from select case
sub Monday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c netstat >> netstat.txt", hidden
wscript.quit
End sub

'''Tuesday procedure will execute from select case
sub Tuesday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c arp -a >> arp.txt", hidden
wscript.quit
End sub

