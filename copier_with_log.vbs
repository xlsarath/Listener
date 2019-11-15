
'Author: psma018 30,October 2019
'Active lister version: 3.6

Option Explicit 
Dim fso,fname,scriptdir,loggingFile,objLog,wshShell,strComputerName,i,name

Set fso = CreateObject("Scripting.FileSystemObject")
Dim sourceFile, destinationFolder,folderIdx,curTIme,folder,qry
scriptdir = fso.GetParentFolderName(WScript.ScriptFullName)
sourceFile = scriptdir&"\results\"
'destinationFolder = "C:\Users\psma018\Documents\Test\"
'loggingFile = scriptdir &"\logs\"
Set wshShell = CreateObject( "WScript.Network" )
strComputerName = wshShell.ComputerName
'WScript.Echo "Computer Name: " & strComputerName

' to insert new line 

Dim bullet
Dim response,destinationResponse,filename,wmi,p,qry2,process
bullet = Chr(10) & "   " & Chr(149) & " " 
Do
    destinationResponse = InputBox("Please give full path of Destination Directory:")
    response = InputBox("Chose the input file directory:" & Chr(10) & bullet & "1.) Detected path :"&sourceFile & bullet & "2.) Enter full path manually" & Chr(10), "Select Directory")
    If (destinationResponse = "") Then WScript.Quit  'Detect Cancel
    If NOT IsNumeric(response) Then Exit Do 'Detect value response.
    MsgBox "Enter valid path", 48, "Invalid Entry"
Loop
'MsgBox "The user chose :" &response, 64, "Hurray!"
  

Function LoggerInstance(loggingFile)
    
    if(fso.FolderExists(loggingFile)) then
    else
        fso.CreateFolder(loggingFile)
    end if  
    if(fso.FileExists(loggingFile&"log.txt")) then
        set objLog = fso.OpenTextFile(loggingFile&"log.txt", 8, true, 0)
        objLog.WriteLine( Now() &" log appending begined now ")
    else 
        Set objLog = fso.CreateTextFile(loggingFile&"log.txt",True)
        objLog.close
        set objLog = fso.OpenTextFile(loggingFile&"log.txt", 8, true, 0)
        objLog.WriteLine( Now()&" Log File created now() " ) 
    end if
End Function

Function FileInUse(qry)
 i=0
    For Each p In wmi.ExecQuery(qry)
                    i = i+1
    Next
End Function


Function FileCopy(name)
fso.CopyFile sourceFile&name, destinationFolder
End Function

Function FileDelete(name)
fso.DeleteFile sourceFile&name
End Function

Dim strPw,strUsr

 
strPw = GetPassword( "Please enter your credentials:" )
WScript.Echo "Your password is: " & strPw & strUsr

Function GetPassword( myPrompt )
' This function uses Internet Explorer to
' create a dialog and prompt for a password.
'
' Version:             2.15
' Last modified:       2015-10-19
'
' Argument:   [string] prompt text, e.g. "Please enter password:"
' Returns:    [string] the password typed in the dialog screen
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
' Error handling code written by Denis St-Pierre
	Dim blnFavoritesBar, blnLinksExplorer, objIE, strHTML, strRegValFB, strRegValLE, wshShell
	
	blnFavoritesBar  = False
	strRegValFB = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\MINIE\LinksBandEnabled"
	blnLinksExplorer = False
	strRegValLE = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\LinksExplorer\Docked"

	Set wshShell = CreateObject( "WScript.Shell" )

	On Error Resume Next
	' Temporarily hide IE's Favorites Bar if it is visible
	If wshShell.RegRead( strRegValFB ) = 1 Then
		blnFavoritesBar = True
		wshShell.RegWrite strRegValFB, 0, "REG_DWORD"
	End If
	' Temporarily hide IE's Links Explorer if it is visible
	If wshShell.RegRead( strRegValLE ) = 1 Then
		blnLinksExplorer = True
		wshShell.RegWrite strRegValLE, 0, "REG_DWORD"
	End If
	On Error Goto 0
	
	' Create an IE object
	Set objIE = CreateObject( "InternetExplorer.Application" )
	' specify some of the IE window's settings
	objIE.Navigate "about:blank"
	' Add string of "invisible" characters (500 tabs) to clear the title bar
	objIE.Document.title = "Password " & String( 500, 7 )
	objIE.AddressBar     = False
	objIE.Resizable      = False
	objIE.StatusBar      = False
	objIE.ToolBar        = False
	objIE.Width          = 420
	objIE.Height         = 280
	' Center the dialog window on the screen
	With objIE.Document.parentWindow.screen
		objIE.Left = (.availWidth  - objIE.Width ) \ 2
		objIE.Top  = (.availheight - objIE.Height) \ 2
	End With
	' Wait till IE is ready
	Do While objIE.Busy
		WScript.Sleep 200
	Loop
	' Insert the HTML code to prompt for a password
	strHTML = "<div style=""text-align: center;"">" _
	        & "<p>" & myPrompt & "</p>" _
                & "<p>username :<input type= ""text"" size=""20"" id = ""username"" value=""Enter your user-id"" onkeyup=" _
                & "</br>" _
	        & "<p>  password :<input type=""password"" size=""20""  id=""Password"" onkeyup=" _
	        & """if(event.keyCode==13){document.all.OKButton.click();}"" /></p>" _
	        & "<p><input type=""hidden"" id=""OK"" name=""OK"" value=""0"" />" _
	        & "<input type=""submit"" value="" OK "" id=""OKButton"" " _
	        & "onclick=""document.all.OK.value=1"" /></p>" _
	        & "</div>"
	objIE.Document.body.innerHTML = strHTML
	' Hide the scrollbars
	objIE.Document.body.style.overflow = "auto"
	' Make the window visible
	objIE.Visible = True
	' Set focus on password input field
        objIE.Document.all.username.focus
	'objIE.Document.all.Password.focus

	' Wait till the OK button has been clicked
	On Error Resume Next
	Do While objIE.Document.all.OK.value = 0 
		WScript.Sleep 200
		' Error handling code by Denis St-Pierre
		If Err Then	' User clicked red X (or Alt+F4) to close IE window
			GetPassword = ""
			objIE.Quit
			Set objIE = Nothing
			' Restore IE's Favorites Bar if applicable
			If blnFavoritesBar Then wshShell.RegWrite strRegValFB, 1, "REG_DWORD"
			' Restore IE's Links Explorer if applicable
			If blnLinksExplorer Then wshShell.RegWrite strRegValLE, 1, "REG_DWORD"
			' Use "WScript.Quit 1" instead of "Exit Function" if you want
			' to abort with return code 1 in case red X or Alt+F4 were used
			Exit Function
		End if
	Loop
	On Error Goto 0

	' Read the password from the dialog window
	GetPassword = objIE.Document.all.Password.value
        strUsr = objIE.Document.all.username.value
	' Terminate the IE object
	objIE.Quit
	Set objIE = Nothing

	On Error Resume Next
	' Restore IE's Favorites Bar if applicable
	If blnFavoritesBar Then wshShell.RegWrite strRegValFB, 1, "REG_DWORD"
	' Restore IE's Links Explorer if applicable
	If blnLinksExplorer Then wshShell.RegWrite strRegValLE, 1, "REG_DWORD"
	On Error Goto 0

	Set wshShell = Nothing
End Function

if response ="" then
    sourceFile = scriptdir&"\results\"
    'loggingFile = scriptdir &"\logs\"
    destinationFolder = destinationResponse&"\"
    Wscript.Echo "destination folder: "&destinationFolder
    loggingFile = fso.GetParentFolderName(destinationFolder) &"\Logs\"
    Wscript.Echo "logging file: "&loggingFile
    call LoggerInstance(loggingFile)
    WScript.Echo ("From directory chosen by user'"&strComputerName&"' :"&sourceFile & Chr(10) & bullet &"logs are available at:"&loggingFile& Chr(10) & bullet &"TO directory chosen : "&destinationResponse )
    objLog.WriteLine("From directory chosen by user'"&strComputerName&"' :"&sourceFile & Chr(10) & bullet &"logs are available at:"&loggingFile& Chr(10) & bullet &"To directory chosen by user: "&destinationResponse )
else 
    sourceFile = response&"\"
    destinationFolder = destinationResponse&"\"
    loggingFile = fso.GetParentFolderName(destinationFolder) &"\Logs\"
    call LoggerInstance(loggingFile)
    WScript.Echo ("From directory chosen by user'"&strComputerName&"' :"&sourceFile & Chr(10) & bullet &"logs are available at:"&loggingFile& Chr(10) & bullet &"To directory chosen by user: "&destinationResponse )
    objLog.WriteLine("From directory chosen by user'"&strComputerName&"' :"&sourceFile & Chr(10) & bullet &"logs are available at:"&loggingFile& Chr(10) & bullet &"To directory chosen by user: "&destinationResponse )
end If    

' Full Computer Name
' can be found by right-clicking My Computer,
' then click Properties, then click the Computer Name tab)
' or use the computer's IP address
'Wscript.Echo (sourceFile)
dim strComputer,objLocator,objWMI,colSwbemObjectSet,objProcess,qrey
'Set folder = fso.getFolder(sourceFile)
strComputer = "cdadrsd01.epridl.com"
set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set wmi = GetObject("winmgmts://./root/cimv2")
set objWMI = objLocator.ConnectServer(strComputer, "root\cimv2", strUsr, strPw)
qrey = "Select * from CIM_DataFile where Path='"&sourceFile&"' and Drive='E:'"
qryPid = "SELECT * FROM Win32_Process WHERE Name = 'WScript.exe'"
WScript.echo qrey
'Wscript.Sleep 10000 'wait for ten seconds 
    Set colSwbemObjectSet = objWMI.ExecQuery(qrey)
bol = true
Do while bol:	
	curTIme = now()
	wscript.Sleep 10000 'sleep for 10 seconds
    For Each objProcess in colSWbemObjectSet
		if objProcess.DateLastModified < curTIme Then
			
			For Each process in wmi.ExecQuery(qryPid)
                If process.CreationDate > datHighest Then
                datHighest = process.CreationDate
                lngMyProcessId = process.ProcessId
                End If    
            next
        	'Wscript.Echo "Process Name: " & objProcess.Name 
        	'Wscript.Echo "\\10.26.176.71\"&(Replace(objProcess.Name,":","")) &"    "&destinationFolder
        	fso.CopyFile "\\10.26.176.71\"&(Replace(objProcess.Name,":","")), destinationFolder, true
        	objLog.WriteLine("Process Name: " & objProcess.Name)
        	'objProcess.delete
        	'wscript.Quit
		End If
	Next
Loop