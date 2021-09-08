#################################################################################################################################
#                       USERS CSV FILE MUST CONTAIN TWO HEADER ROWS "UserPrincipalName" and "Mobile"                            #
#################################################################################################################################

Install-Module MSOnline
Install-module Microsoft.Graph
Select-MgProfile -Name Beta


connect-graph -Scopes @("UserAuthenticationMethod.Read.All"; "UserAuthenticationMethod.ReadWrite.All")


Add-Type -AssemblyName System.Windows.Forms
New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }


$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.InitialDirectory = $InitialDirectory
$OpenFileDialog.filter = "CSV (*.csv) | *.csv"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.FileName


if ($OpenFileDialog.FileName) {

    $Users = Import-CSV $OpenFileDialog.FileName

    $i = 0

    ForEach ($user in $users) {
        $i++

        $phone = "+1 " + $user.Mobile

        Write-Host "Working on User" $i "of" $users.Count "-" $user.UserPrincipalName -ForegroundColor Yellow
        Write-Host "Setting Number" $phone -ForegroundColor DarkYellow
    
        $currentmobile = (Get-MgUserAuthenticationPhoneMethod -UserId $user.UserPrincipalName | where { $_.PhoneType -eq "mobile" }).PhoneNumber
        If ($currentmobile) {
            If ($currentmobile -eq $phone) {
                Write-Host "Number already matches - No action taken" -ForegroundColor Red
            }
            If ($currentmobile -ne $phone) {
                Write-Host "Different number is already populated - No action taken" -ForegroundColor Red
            }
        }
        If (!$currentmobile) {
            New-MgUserAuthenticationPhoneMethod -UserId $user.UserPrincipalName -phoneType "mobile" -phoneNumber $phone
            Write-Host "No current mobile number - populated with" $phone -ForegroundColor Green
        }
    }
}
else {
    Write-Warning -Message "No Users CSV selected"
}

Disconnect-Graph