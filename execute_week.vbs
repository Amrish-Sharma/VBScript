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
    Case 4
    Call Wednesday()
    Case 5 
    Call Thursday()
    Case 6 
    Call Friday()
    Case 7 
    Call Saturday()
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
WShShell.run "cmd /c ping -n 10 www.devry.edu >> ping.txt", hidden
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

sub Wednesday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c nbstat -n >> nbstat.txt", hidden
wscript.quit
End sub

sub Thursday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c tracert www.devry.edu >> tracert.txt", hidden
wscript.quit
End sub

sub Friday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c ipconfig >> ipconfig.txt",hidden
Set file = fso.OpenTextFile("ipconfig.txt", 1)
text = file.ReadAll
''' to output the content of last command on the screen
Wscript.Echo text
file.Close
wscript.quit
End sub

sub Saturday()
'''defining variables
dim wshShell
dim path
dim fso
'''setting up the environment to run vbscript
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
''' to execute the command and append to the text file using >> if you want to text to be overriden use >
WShShell.run "cmd /c hostname >> hostname.txt",hidden
'''to read all the content stored in hostname.txt
Set file = fso.OpenTextFile("hostname.txt", 1)
text = file.ReadAll
''' to output the content of last command on the screen
Wscript.Echo text
file.Close
End sub

