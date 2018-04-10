'Init two arrays for tracking program installation
Dim installArray(8)
Dim notInstallArray(8)
'Init index for ease of use

Set fso = CreateObject("Scripting.FileSystemObject")

acrobat = "C:\Program Files (x86)\Adobe\Acrobat 2017\Acrobat"
chrome = "C:\Program Files (x86)\Google\Chrome\Application"
java = "C:\Program Files\Java\jre1.8.0_101\bin"
vlc = "C:\Program Files\VideoLAN\VLC"
office = "C:\Program Files\Microsoft Office\Office16"
firefox = "C:\Program Files (x86)\Mozilla Firefox"
flash = "C:\Windows\System32\Macromed\Flash"

Dim filePaths
Dim executables
Dim programNames

filePaths = Array(acrobat,chrome,java,vlc,office,firefox,flash)
executables = Array("acrobat.exe","chrome.exe","java.exe","vlc.exe","excel.exe","firefox.exe","flashutil_activex.exe")
programNames = Array("Acrobat 2017","Chrome","Java","VLC","Office 2016","Firefox","Flash")

Dim index
index=0

sub executableExists(filePaths)
    for each filePath in filePaths
        If index = 7 then
            Exit For
        End If
        If fso.FolderExists(filePath) then
        For Each f In fso.GetFolder(filePath).Files            
            If LCase(f.Name) = executables(index) Then
                installArray(index) = programNames(index)
                'quit once the EXE is found in the file path
                Exit For
            End If
            Next
            If installArray(index) <> programNames(index) then
                notInstallArray(index) = programNames(index)
                index=index+1
            Else
                index=index+1
            End if
        Else
            notInstallArray(index) = programNames(index)
            index=index+1
        End If
    Next
End sub

executableExists filePaths

Dim installedPrograms
Dim ninstalledPrograms
Dim spacer
Dim separator
Dim downMessage
Dim elemNotInstalled
spacer = "              "
separator = vbCrLf & vbCrLf
downMessage = "Click OK to Shutdown"

Set Shell = CreateObject("Wscript.Shell")

sub checkInstalls(array)
    for each program in installArray
        if program <> "" then
            installedPrograms = installedPrograms + vbCrLf + "-" +program
        End if
    Next
    for each program in notInstallArray
        if program <> "" then
            ninstalledPrograms = ninstalledPrograms + vbCrLf+"-" + program
            elemNotInstalled = elemNotInstalled+1
        End if
    Next
    if elemNotInstalled = 0 then
        result=MsgBox( "Installed Programs" & installedPrograms & separator & downMessage,64,"Program Installations")
    Else
        result=MsgBox("Installed Programs" & installedPrograms & separator & "Programs Not Installed" & ninstalledPrograms & separator & downMessage,64,"Program Installations")
    End if

    if result=1 then
        Shell.Run "shutdown.exe /R /T 5 /C ""Rebooting your computer now!"" "    
    End if
End sub

checkInstalls installArray

