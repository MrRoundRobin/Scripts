[cmdletbinding()]
param (
    [Parameter(Mandatory)]
    [ValidatePattern('^\S{1,}@\S{2,}\.\S{2,}$')]
    [string]$UPN,

    [Parameter(Mandatory)]
    [SecureString]$RoomPassword,

    [Parameter(Mandatory)]
    [string]$DisplayName,

    [ValidateLength(2,2)]
    [string]$UsageLocation = 'DE',

    [switch]$Force
)

$Alias = ($UPN -split '@')[0]

#Requires -Module AzureAD,ExchangeOnlineManagement

#region connect

'Connecting...' | Write-Verbose

Connect-ExchangeOnline -ShowBanner:$false
Connect-AzureAD | Out-Null

#endregion connect

'Checking if Mailbox alreay exists' | Write-Verbose

$Mailbox = Get-Mailbox -Identity $Alias -ErrorAction SilentlyContinue

if ($null -ne $Mailbox) {
    'Mailbox found, checking parameters' | Write-Verbose
    
    if ($Mailbox.MicrosoftOnlineServicesID -ne $UPN) {
        'Mailbox has a different UPN: {0}' -f $Mailbox.MicrosoftOnlineServicesID | Write-Warning
    }
    
    if ($Mailbox.DisplayName -ne $DisplayName) {
        'Mailbox has a different DisplayName: {0}' -f $Mailbox.DisplayName | Write-Warning
    }

    if (-not $Mailbox.IsMailboxEnabled -or -not $Mailbox.RoomMailboxAccountEnabled) {
        'Mailbox is disabled' | Write-Warning
    }

    'Password cannot be checked' | Write-Warning

    if ($Force) {
        $Mailbox | Set-Mailbox -MicrosoftOnlineServicesID $UPN `
                               -DisplayName $DisplayName `
                               -EnableRoomMailboxAccount $true `
                               -RoomMailboxPassword $RoomPassword `
                               -AccountDisabled $false `
                               -Force `
                               -ErrorAction Stop
    } else {
        throw 'Mailbox already exists, if you want to override the propertirs use -Force'
    }
} else {
    'Creating New Mailbox {0}' -f $UPN | Write-Verbose

    New-Mailbox -MicrosoftOnlineServicesID $UPN `
                -Alias $Alias `
                -DisplayName $DisplayName `
                -Room `
                -EnableRoomMailboxAccount $true `
                -RoomMailboxPassword $RoomPassword `
                -ErrorAction Stop `
        | Out-Null
}

'Setting Calendar properties' | Write-Verbose

Set-CalendarProcessing -Identity $UPN `
                       -AutomateProcessing AutoAccept `
                       -AddOrganizerToSubject $false `
                       -AllowConflicts $false `
                       -DeleteComments $false `
                       -DeleteSubject $false `
                       -RemovePrivateProperty $false `
                       -AddAdditionalResponse $true `
                       -AdditionalResponse 'This room is equipped with a Surface Hub' `
                       -ErrorAction Stop

$i = 0
do {
    'Waiting for user to be created ({0})' -f ++$i | Write-Verbose

    Start-Sleep -Seconds 10
    $user = Get-AzureADUser -ObjectId $UPN -ErrorAction SilentlyContinue  
} while ($null -eq $user -and $i -lt 10)
    

if ($null -eq $user) {
    throw 'New User not found after creation. Meybe we need to give the cloud more time to process and repeat in a bit.'
}

'Setting PasswordExpiration and Usage Location' | Write-Verbose

Set-AzureADUser -ObjectId $UPN `
                -PasswordPolicies "DisablePasswordExpiration" `
                -UsageLocation $UsageLocation  `
                -ErrorAction Stop

'Adding License' | Write-Verbose

$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = "6070a4c8-34c6-4937-8dfb-39bbc6397a60"

$AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$AssignedLicenses.AddLicenses = $License

Set-AzureADUserLicense -ObjectId $UPN -AssignedLicenses $AssignedLicenses -ErrorAction Stop
