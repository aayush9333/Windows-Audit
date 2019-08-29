Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("freespace.txt", True)

Const HARD_DISK = 3
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk Where DriveType = " & HARD_DISK & "")
For Each objDisk in colDisks
    objFile.WriteLine "DeviceID: "& vbTab &  objDisk.DeviceID       
    objFile.WriteLine "Free Disk Space: "& vbTab & objDisk.FreeSpace
    objFile.WriteLine ""
Next