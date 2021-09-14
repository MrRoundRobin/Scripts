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
    [string]$UsageLocation = 'DE'
)

$Alias = ($UPN -split '@')[0]

#Requires -Module AzureAD

#region Check for modules

$MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)

If ($null -eq $MFAExchangeModule) {
    'Please install Exchange Online MFA Module.' | Write-Warning
    
    Start-Process "https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application"

    Read-Host -Prompt 'Press Enter after installing the Exchange Online MFA Module' | Out-Null

    $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    If ($null -eq $MFAExchangeModule) {
        throw 'Exchange Online MFA module is not available'
    }
}

#endregion Check for MFA module

#region connect

'Connecting...' | Write-Verbose

. "$MFAExchangeModule"
Connect-EXOPSSession -WarningAction SilentlyContinue | Out-Null
Connect-AzureAD

#endregion connect

'Creating New Mailbox {0}' -f $UPN | Write-Verbose

New-Mailbox -MicrosoftOnlineServicesID $UPN `
            -Alias $Alias `
            -Name $DisplayName `
            -Room `
            -EnableRoomMailboxAccount $true `
            -RoomMailboxPassword $RoomPassword `
    | Out-Null

'Setting Calendar properties' | Write-Verbose

Set-CalendarProcessing -Identity $UPN `
                       -AutomateProcessing AutoAccept `
                       -AddOrganizerToSubject $false `
                       -AllowConflicts $false `
                       -DeleteComments $false `
                       -DeleteSubject $false `
                       -RemovePrivateProperty $false `
                       -AddAdditionalResponse $true `
                       -AdditionalResponse "This room is equipped with a Surface Hub"

$i = 0
do {
    'Waiting for user to be created ({0})' -f ++$i | Write-Verbose

    Start-Sleep -Seconds 10
    $user = Get-AzureADUser -ObjectId $UPN -ErrorAction SilentlyContinue  
} while ($null -eq $user -and $i -lt 10)
    

if ($null -eq $user) {
    throw 'New User not found 🤦'
}

'Setting PasswordExpiration and Usage Location' | Write-Verbose

Set-AzureADUser -ObjectId $UPN `
                -PasswordPolicies "DisablePasswordExpiration" `
                -UsageLocation $UsageLocation

'Adding License' | Write-Verbose

$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = "6070a4c8-34c6-4937-8dfb-39bbc6397a60"

$AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$AssignedLicenses.AddLicenses = $License
$AssignedLicenses.RemoveLicenses = @()

Set-AzureADUserLicense -ObjectId $UPN -AssignedLicenses $AssignedLicenses
