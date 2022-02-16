<#
.SYNOPSIS
	This script list all users with MFA Status
.NOTES
	Created on: 	16.02.2022
	Created by: 	Adam Kołodziejski
	Filename:		Get-MFAUsersState.ps1
	Requirements:	Installed MsolService Module and Import-Excel Module
.EXAMPLE
	Get-MFAUsersState.ps1 -FilePath "C:\Temp" -GuestUsers:$true
	This example will get all users with Guest Users and list to file C:\Temp\MFAUsers.xlsx
.EXAMPLE
	Get-MFAUsersState.ps1 -FilePath "C:\Temp" -GuestUsers:$false
	This example will get only member users without Guest Users and list to file C:\Temp\MFAUsers.xlsx
.PARAMETER FilePath
	Is responsible for output file path
.PARAMETER GuestUsers
	Is responsible for mode (with or without Guest Users)
#>
[CmdletBinding()]
param ([Parameter(Mandatory = $true)]
    [string]$FilePath,
    [switch]$GuestUsers
)
Connect-MsolService

$NoMFAList = $null
Write-Host "Finding Azure Active Directory Accounts..."
if ($GuestUsers -eq $true){
    $Users = Get-MsolUser -All
}
else{
    $Users = Get-MsolUser -All | Where-Object {$_.userType -ne 'Guest'}
}

$ReportMFA = [System.Collections.Generic.List[Object]]::new() # Create output file
Write-Host "Processing" $Users.Count "accounts..." 
ForEach ($User in $Users) {
    $MFAEnforced = $User.StrongAuthenticationRequirements.State #Get user MFA Status
    $MFAPhone = $User.StrongAuthenticationUserDetails.PhoneNumber #Get user phone number
    #Get User Authentication Methods
    $DefaultMFAMethod = ($User.StrongAuthenticationMethods | ? { $_.IsDefault -eq "True" }).MethodType
    #If User have enabled MFA Convert value mfa verification to simple word
    If (($MFAEnforced -eq "Enforced") -or ($MFAEnforced -eq "Enabled")) {
        Switch ($DefaultMFAMethod) {
            "OneWaySMS" { $MethodUsed = "One-way SMS" }
            "TwoWayVoiceMobile" { $MethodUsed = "Phone call verification" }
            "PhoneAppOTP" { $MethodUsed = "Hardware token or authenticator app" }
            "PhoneAppNotification" { $MethodUsed = "Authenticator app" }
        }
    }
    Else {
        $MFAEnforced = "Not Enabled"
        Switch ($DefaultMFAMethod) {
            "OneWaySMS" { $MethodUsed = "One-way SMS" }
            "TwoWayVoiceMobile" { $MethodUsed = "Phone call verification" }
            "PhoneAppOTP" { $MethodUsed = "Hardware token or authenticator app" }
            "PhoneAppNotification" { $MethodUsed = "Authenticator app" }
            default { $MethodUsed = "MFA Not Used" }
        }
    }
    
    #create new line report
    $ReportLine = [PSCustomObject] @{
        User        = $User.UserPrincipalName
        Name        = $User.DisplayName
        MFAUsed     = $MFAEnforced
        MFAMethod   = $MethodUsed 
        PhoneNumber = $MFAPhone
    }
     
    $ReportMFA.Add($ReportLine) #Add new line
}

#Export Report

$ReportMFA | Export-Excel -Path $FilePath\MFAUsers.xlsx
