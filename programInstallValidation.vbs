Dim check
Dim installArray(8)
Dim notInstallArray(8)
Dim index

index=0
Set fso = CreateObject("Scripting.FileSystemObject")
 If fso.FolderExists("C:\Program Files (x86)\Adobe\Acrobat 2017\Acrobat") then
    For Each f In fso.GetFolder("C:\Program Files (x86)\Adobe\Acrobat 2017\Acrobat").Files
        If LCase(f.Name) = "acrobat.exe" Then
            installArray(index) = "Acrobat 2017"
            index=index+1
            Exit For
        End If
        Next
        If installArray(0) <> "Acrobat 2017" then
            notInstallArray(index) = "Acrobat 2017"
            index=index+1
        End if
Else
    notInstallArray(index) = "Acrobat 2017"
    index=index+1
End If

If fso.FolderExists("C:\Program Files (x86)\Google\Chrome\Application") then
    For Each f In fso.GetFolder("C:\Program Files (x86)\Google\Chrome\Application").Files
        If LCase(f.Name) = "chrome.exe" Then
            installArray(index) = "Chrome"
            index=index+1
            Exit For
        End If
        Next
        if installArray(1) <> "Chrome" then
            notInstallArray(index) = "Chrome"
            index = index+1
        End if
Else
    notInstallArray(index) = "Chrome"
    index=index+1
End If

If fso.FolderExists("C:\Program Files\Java\jre1.8.0_101\bin") then
    For Each f In fso.GetFolder("C:\Program Files\Java\jre1.8.0_101\bin").Files
        If LCase(f.Name) = "java.exe" Then
            installArray(index) = "Java"
            index=index+1
            Exit For
        End If
        Next
        if installArray(2) <> "Java" then
            notInstallArray(index) = "Java"
            index = index+1
        End if
Else
    notInstallArray(index) = "Java"
    index=index+1
End If

If fso.FolderExists("C:\Program Files\VideoLAN\VLC") then
    For Each f In fso.GetFolder("C:\Program Files\VideoLAN\VLC").Files
        If LCase(f.Name) = "vlc.exe" Then
            installArray(index) = "VLC"
            index=index+1
            Exit For
        End If
        Next
            if installArray(3) <> "VLC" then
                notInstallArray(index) = "VLC"
                index = index+1
            End if
Else
    notInstallArray(index) = "VLC"
    index=index+1
End If

If fso.FolderExists("C:\Program Files\Microsoft Office\Office16") then
    For Each f In fso.GetFolder("C:\Program Files\Microsoft Office\Office16").Files
        If LCase(f.Name) = "excel.exe" Then
            installArray(index) = "Office 2017"
            index=index+1
            Exit For
        End If
        Next
            if installArray(4) <> "Office 2017" then
                notInstallArray(index) = "Office 2017"
                index = index+1
            End if
Else
    notInstallArray(index) = "Office 2017"
    index=index+1
End If

If fso.FolderExists("C:\Program Files (x86)\Mozilla Firefox") then
    For Each f In fso.GetFolder("C:\Program Files (x86)\Mozilla Firefox").Files
        If LCase(f.Name) = "firefox.exe" Then
            installArray(index) = "Firefox"
            index=index+1
            Exit For
        End If
        Next
            if installArray(5) <> "Firefox" Then
                notInstallArray(index) = "Firefox"
                index = index+1
            End If
Else
    notInstallArray(index) = "Firefox"
    index=index+1
End If

If fso.FolderExists("C:\Windows\System32\Macromed\Flash") then
    For Each f In fso.GetFolder("C:\Windows\System32\Macromed\Flash").Files
        If LCase(f.Name) = "flashutil_activex.exe" Then
            installArray(index) = "Flash"
            index=index+1
            Exit For
        End If
        Next
            if installArray(6) <> "Flash" then
                notInstallArray(index) = "Flash"
                index = index+1
            End if
Else
    notInstallArray(index) = "Flash"
    index=index+1
End If
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

