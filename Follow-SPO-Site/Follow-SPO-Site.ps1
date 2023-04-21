# PowerShell Script for Following a Site via GraphAPI in SharePoint Online 
# Copyright (c) 2023 - Manuel Frick - https://www.m365fox.com/
# Sources used: https://charleslakes.com/2021/11/15/graph-api-follow-sharepoint-sites/

# Install the MSAL.PS module if not already installed
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Install-Module MSAL.PS -Scope CurrentUser
}

# Application and tenant information
$AppId = "<Your-Application-ID>"
$TenantId = "<Your-Tenant-ID>"
$ClientSecret = "<Your-Client-Secret>"

# Authenticate and get an access token
$AppCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, (ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force)
$AccessToken = (Get-MsalToken -ClientId $AppId -TenantId $TenantId -ClientCredential $AppCredential -Scopes "https://graph.microsoft.com/.default").AccessToken

# Function to make API requests
function Get-APIResponse {
    param(
        [Parameter(Mandatory)] [System.String] $APIGetRequest
    )
    
    $apiUrl = "https://graph.microsoft.com/v1.0$APIGetRequest"
    $headers = @{
        'Authorization' = "Bearer $AccessToken"
    }

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $headers
    }
    catch [Exception] {
        throw $_
    }
    return $response
}

# Function to make API batch requests
function Get-APIBatchResponse {
    param(
        [Parameter(Mandatory)] [System.Array] $BatchRequests
    )
    
    $apiUrl = "https://graph.microsoft.com/v1.0/$batch"
    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type' = "application/json"
    }

    $body = @{
        "requests" = $BatchRequests
    }

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body (ConvertTo-Json -InputObject $body -Depth 10)
    }
    catch [Exception] {
        throw $_
    }
    return $response
}

# Define function to retrieve the site GUID
function Get-SPOSiteGuid {
    param(
        [Parameter(Mandatory)] [System.String] $RelativePath
    )
 
    [System.String] $thisGuid = ""
    [System.String] $thisReqs = "/sites/contoso.sharepoint.com:$($RelativePath)"
  
    try {
        (Get-APIResponse -APIGetRequest "$thisReqs") | % {
            $thisGuid = "$($_.id)"
        }
    }
    catch [Exception] {}
    return $thisGuid
}

# Define function to retrieve the AD Group ID
function Get-ADGroupId {
    param(
        [Parameter(Mandatory)] [System.String] $GroupName
    )
 
    [System.String] $thisGuid = ""
    [System.String] $thisReqs = "/groups?`$select=id,displayName&`$filter=startswith(displayName, '$GroupName')"
  
    try {
        (Get-APIResponse -APIGetRequest "$thisReqs").value | % {
            $thisGuid = "$($_.id)"
        }
    }
    catch [Exception] {}
    return $thisGuid
}

# Define function to retrieve the AD Group Members
function Get-ADGroupMembers {
    param(
        [Parameter(Mandatory)] [System.String] $GroupGUID
    )
 
    [System.Collections.Hashtable] $listOf = @{}
    [System.String] $thisReqs = "/groups/$($GroupGUID)/members?`$select=id,displayName"
  
    try {
        (Get-APIResponse -APIGetRequest "$thisReqs").value | % {
            $listOf.Add(
                "$($_.displayName)", "$($_.id)"
            )
        }
    }
    catch [Exception] {}
    return $listOf
}

# Get the Site GUID
[System.String] $sPath = "/sites/Intranet"
[System.String] $sGuid = (Get-SPOSiteGuid -RelativePath $sPath)

# Get the AD Group ID
[System.String] $gName = "New Hires"
[System.String] $gGuid = (Get-ADGroupId -GroupName $gName)

# Get the AD Group Members and follow the site
[System.Array] $batchOf = @()
foreach($member in ((Get-ADGroupMembers -GroupGUID $gGuid).GetEnumerator())) {
 
    $batchOf += @{
        "url" = "/users/$($member.Value)/followedSites/add"
        "method" = "POST"
        "id" = "$($batchOf.Count + 1)"
        "body" = @{
            "value" = @(
                @{
                    "id" = "$($sGuid)"
                }
            )
        }
        "headers" = @{
            "Content-Type" = "application/json"
        }
    }
}

# Execute the batch request
if ($batchOf.Length -ne 0) {
    Get-APIBatchResponse -BatchRequests $batchOf
}
