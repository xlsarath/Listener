
'Author: psma018 30,October 2019
'Active lister version: 3.6

Option Explicit 
Dim fso,fname,scriptdir,loggingFile,objLog,wshShell,strComputerName

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


if response ="" then
    sourceFile = scriptdir&"\results\"
    'loggingFile = scriptdir &"\logs\"
    destinationFolder = destinationResponse&"\"
    loggingFile = fso.GetParentFolderName(destinationFolder) &"\Logs\"
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



'List files in source directory into log file and move'em in sequence
'Note: No folders will be copied/deteled only files will be moved to destination folder

Set folder = fso.getFolder(sourceFile)

'Searching for processes
Dim strScriptName
Dim datHighest
Dim lngMyProcessId

'Which script to look for ? 
strScriptName = "WScript.exe"
'strScriptName = "Notepad.exe"
'Iniitialise 
datHighest = Cdbl(0)

Do while folder.files.count <> -1 
curTIme = now()
WScript.Sleep 10000          'wait for ten seconds
    For each folderIdx In folder.Files
        'objLog.WriteLine(folderIdx.Name)
        if (folderIdx.DateLastModified < curTIme) Then
            'WScript.Echo("File Modified: " &  CDate( fso.DateLastModified))
            'wscript.echo "File Created :" &CDate(fso.DateCreated)
            'wscript.echo " "&fso.DateCreated
            Set wmi = GetObject("winmgmts://./root/cimv2")
            qry = "SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & folderIdx.Name & "%'"
            qry2 = "SELECT * FROM Win32_Process WHERE Name = '" & strScriptName & "'"
            For Each process in wmi.ExecQuery(qry2)
                If process.CreationDate > datHighest Then
                datHighest = process.CreationDate
                lngMyProcessId = process.ProcessId
                End If    
            next
            objLog.WriteLine Now() &", Process Id : "&lngMyProcessId &", File : " &sourceFile&folderIdx.Name 
           'WScript.Echo("process Id = "&lngMyProcessId)    
            For Each p In wmi.ExecQuery(qry)
                    WScript.Echo folderIdx.Name &"file in use, close the File and re-start service" 
                    objLog.WriteLine now()&", Process Id : "&lngMyProcessId &", " &folderIdx.Name &", file is in use!! close the File and re-start service" 
                    objLog.WriteLine Now()&", Process Id : "&lngMyProcessId &", Process Terminated!"
                    objLog.close
                    WScript.Quit 0
            Next
            'WScript.Echo "file is not in use, resume the work flow"   
            'wscript.echo folderIdx.DateCreated
            'wscript.echo folderIdx.DateLastModified
            curTIme = Now()
            fso.CopyFile sourceFile&folderIdx.Name, destinationFolder
            fso.DeleteFile sourceFile&folderIdx.Name
            objLog.WriteLine now()&", Process Id : "&lngMyProcessId &", file moved to " &destinationFolder &", time elapsed : " &DateDiff("s",Now(),curTIme) &"seconds"
        End if 
    Next    
'MsgBox("Still running")    
Loop
'objLog.WriteLine "Files copying begined at " &curTIme &" from source :" &sourceFile
'fso.CopyFile sourceFile&"*.*", destinationFolder
'call MoveEm(sourceFile, destinationFolder)


'fso.DeleteFile sourceFile&"*.*"