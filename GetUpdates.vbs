' GetWMIUpdates.vbs
' This script returns all the updates on a computer. On Windows Vista this returns only those supplied by
' component based servicing.
' -------------------------------------------------------' 
Option Explicit

Dim objWMIService, objItem, colItems, strComputer, item
strComputer = "."

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fileOutput
set fileOutput = fso.CreateTextFile("InstalledUpdates.htm",true,false)
fileoutput.WriteLine("<!-- saved from url=(0039)http://msinfluentials.com/blogs/jesper/ -->")
fileOutput.WriteLine("<HTML><HEAD><TITLE>Installed updates</TITLE></HEAD><BODY>")
fileOutput.WriteLine("<TABLE border=""1"">")
fileOutput.WriteLine("<tr style=""background-color:#a0a0ff;font:10pt Tahoma;font-weight:bold;"" align=""left"">")
fileOutput.WriteLine("<TD>Caption</TD><TD>Description</TD><TD>Hotfix ID</TD><TD>KB Link</TD><TD>Installed On</TD><TD>Service Pack in effect</TD><TD>Fix comments</TD></TR>")

'Run the query
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
    
Dim QFEs
Set QFEs = objWMIService.ExecQuery ("Select * from win32_QuickFixEngineering")

Dim strOutput
Dim QFE
For Each QFE in QFEs
	if QFE.HotFixID <> "File 1" then
		strOutput = "<TR style=""background-color:#e0f0f0;font:10pt Tahoma;"">"
		
		strOutput = strOutput + "<TD>" & QFE.Caption & "</TD>" &_
					"<TD>" & QFE.Description & "</TD>" &_
					"<TD>" & QFE.HotFixID & "</TD>"  &_	
					"<TD> <a href=""http://support.microsoft.com/?kbid=" & getKBID(QFE.HotFixID) & """>" & getKBID(QFE.HotFixID) & "</a></TD>" &_		
					"<TD>" & QFE.InstalledOn & "</TD>"  &_
					"<TD>" & QFE.ServicePackInEffect & "</TD>"  &_
					"<TD>" & QFE.FixComments & "</TD>" 
		fileOutput.WriteLine(strOutPut)
		fileOutput.WriteLine("</TR>")

'TODO: The InstalledOn date on Vista is a datetime, not a string like on XP/WS2K3. To fix it we need to use something like this:
'dim dtmInstallDate
'Set dtmInstallDate = CreateObject("WbemScripting.SWbemDateTime")
'dtmInstallDate.value = QFE.InstalledOn
'Wscript.Echo dtmInstallDate.GetVarDate

	end if
Next

					

fileOutput.WriteLine("</TABLE></BODY></HTML>")

WScript.Echo "The updates on your system are now listed in InstalledUpdates.htm"

WScript.Quit

function getKBID(HotFixID)
	' The Hotfix ID is usually the KB number, but it can take several formats
	' On older hotfixes it is often Q123456. Newer ones may be KB123456. Some are just 123456
	' We will just get the right-most six bytes and if those are numeric, assume it is a hotfix number.
	Dim KBNumber
	KBNumber = right(HotFixID,6)
	if isNumeric(KBNumber) then
		getKBID = KBNumber
	else
		getKBID = "No KB Number Found"
	end if
end function