# PowerShell Script for following a Site via GraphAPI in SharePoint Online 
# Copyright (c) 2023 - Manuel Frick - https://www.m365fox.com/

# Install the MSAL.PS module if not already installed
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Install-Module MSAL.PS -Scope CurrentUser
}

# Import the required MSAL.PS module
Import-Module MSAL.PS

# Set your tenant and application information
$tenantId = "<Your-Tenant-ID>"
$appId = "<Your-Application-ID>"
$AppSecret = "<Your-Client-Secret>"
$Scopes = "https://graph.microsoft.com/.default"

# Convert plain text secret to a SecureString
$SecureAppSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force

# Get the access token
$AccessToken = (Get-MsalToken -ClientId $AppId -TenantId $TenantId -ClientSecret $SecureAppSecret -Scopes $Scopes).AccessToken

function Get-APIResponse {
    param(
        [Parameter(Mandatory)] [System.String] $APIGetRequest
    )

    $apiUrl = "https://graph.microsoft.com/v1.0$APIGetRequest"
    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type' = "application/json"
    }
    Write-Host "Request URL: $apiUrl" # Add this line for debugging

    try {
        $response = Invoke-RestMethod -Method Get -Uri $apiUrl -Headers $headers
        Write-Host "Response: $($response | ConvertTo-Json)" # Add this line for debugging
    }
    catch {
        throw $_.Exception
    }
    return $response
}



# Function to make API batch requests
function Get-APIBatchResponse {
    param(
        [Parameter(Mandatory)] [System.Array] $BatchRequests
    )
    
    $batch = $batchRequests.URL
    $body = $BatchRequests.body
    $apiUrl = "https://graph.microsoft.com/v1.0$batch"
    $headers = @{
        'Authorization' = "Bearer $AccessToken"
        'Content-Type' = "application/json"
    }

    #$body = @{
    #    "requests" = $BatchRequests
    #}

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body (ConvertTo-Json -InputObject $body -Depth 10)
    }
    catch [Exception] {
        throw $_
    }
    return $response
}



# Snippet Source: https://charleslakes.com/2021/11/15/graph-api-follow-sharepoint-sites/
# Change the URL for your Microsoft Tenant and replace "contoso"

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
        if ([string]::IsNullOrEmpty($thisGuid)) {
            throw "Error: Site not found or an issue with the Graph API call."
        }
    }
    catch {
        Write-Host $_.Exception.Message
    }
    return $thisGuid
}


function Get-ADGroupId {
    param(
        [Parameter(Mandatory)] [System.String] $GroupName
    )

    [System.String] $thisGuid = ""
    [System.String] $thisReqs = "/groups?`$select=id,displayName&`$filter=displayName eq '$GroupName'"

    try {
        $response = (Get-APIResponse -APIGetRequest "$thisReqs").value
        if ($response.Count -eq 1) {
            $thisGuid = $response[0].id
        } elseif ($response.Count -gt 1) {
            throw "Error: Multiple groups found with the same name. Please ensure the group name is unique."
        } else {
            throw "Error: Group not found."
        }
    }
    catch {
        Write-Host $_.Exception.Message
    }
    return $thisGuid
}


function Get-ADGroupMembers {
    param(
        [Parameter(Mandatory)] [System.String] $GroupID
    )

    [System.Collections.Hashtable] $listOf = @{}
    [System.String] $thisReqs = "/groups/$($GroupID)/members?`$select=id,displayName"

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


#Replace the SharePoint Site-URL and the Azure AD Group Name with your values.

# SharePoint site URL
[System.String] $sPath = "/sites/intranet"

# Get SharePoint Site ID
[System.String] $sid = (Get-SPOSiteGuid -RelativePath $sPath)

# Azure AD Group Name
[System.String] $groupName = "My Group"

# Get Azure AD Group ID
[System.String] $gId = (Get-ADGroupId -GroupName $groupName)

# Get Azure AD Group Members
if (-not [string]::IsNullOrEmpty($gId)) {
    [System.Collections.Hashtable] $groupMembers = (Get-ADGroupMembers -GroupID $gId)
} else {
    Write-Host "Error: Group ID is empty."
    exit
}

# Prepare batch requests
[System.Array] $batchRequests = @()
foreach ($member in $groupMembers.GetEnumerator()) {
    $batchRequests += @{
        "url" = "/users/$($member.Value)/followedSites/add"
        "method" = "POST"
        "id" = "$($batchRequests.Count + 1)"
        "body" = @{
            "value" = @(
                @{
                    "id" = "$($sId)"
                }
            )
        }
        "headers" = @{
            "Content-Type" = "application/json"
        }
    }
}

# Execute batch requests
if ($batchRequests.Length -ne 0) {
    Get-APIBatchResponse -BatchRequests $batchRequests
}
