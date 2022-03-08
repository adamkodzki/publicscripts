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
    [switch]$GuestUsers = $false
)
Connect-MsolService

$AppId = "db779784-9d1e-46b0-95e1-5d8273101d1d"
$AppSecret = "vCS7Q~Dkx0RfiF0ZobhsWcmZzRnvt4p.0_VmQ"
$TenantId = "2ad2e312-9958-441e-8a9e-cf1794e1ee08"

# Construct URI and body needed for authentication
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}

$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

# Unpack Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
$Headers = @{
    'Content-Type'  = "application\json"
    'Authorization' = "Bearer $Token" 
}

$NoMFAList = $null
Write-Host "Finding Azure Active Directory Accounts..."
if ($GuestUsers -eq $true) {
    $Users = Get-MsolUser -All
}
else {
    $Users = Get-MsolUser -All | Where-Object { $_.userType -ne 'Guest' }
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
    #Get User Manager
    $UserPrincipalName = $User.UserPrincipalName
    $Params = @{
        "URI"         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/manager"
        "Headers"     = $Headers
        "Method"      = "GET"
        "ContentType" = 'application/json'
    
    }   
    try {
        $Manager = Invoke-RestMethod @Params #invoke request to MS Graph
    }
    catch {
        Write-Host "Don't find Manager for " + $User.DisplayName
    }
    
    #Get User Company Name
    $CompanyParams = @{
        "URI"         = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName`?`$select=companyName"
        "Headers"     = $Headers
        "Method"      = "GET"
        "ContentType" = 'application/json'
    
    }   
    try {
        $Company = Invoke-RestMethod @CompanyParams #invoke request to MS Graph
    }
    catch {
        Write-Host "Don't find Company name for " $User.DisplayName
    }
    #create new line report
    $ReportLine = [PSCustomObject] @{
        User        = $User.UserPrincipalName
        Name        = $User.DisplayName
        Department  = $User.Department
        CompanyName = $Company.companyName
        Manager     = $Manager.displayName
        PhoneNumber = $User.MobilePhone
        MFAUsed     = $MFAEnforced
        MFAMethod   = $MethodUsed 
    }
     
    $ReportMFA.Add($ReportLine) #Add new line
}

#Export Report

$ReportMFA | Export-Excel -Path $FilePath\MFAUsers.xlsx -AutoSize -AutoFilter
