<#
.SYNOPSIS
    Enable and Add Microsoft 365 email notifications for Audit purposes
.DESCRIPTION
	This will delete any Audit items users, and Audit email notifications to Users configured for these alerts.
	While commands to remove both Audit items, and Audit email notifications have been included, from testing it appeared that this API is the only method that worked - https://helpcenter.veeam.com/docs/vbo365/rest/reference/vbo365-rest.html?ver=70#tag/AuditEmailSettings
	The Audit items API is referenced here - https://helpcenter.veeam.com/docs/vbo365/rest/reference/vbo365-rest.html?ver=70#tag/OrganizationAudit
	
	To update the email notifications users, rerun the script and update the $mailTo variable with each of the email addresses
#>
#Ensures that self signed certificates are ignored
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

#Alter these variables to match the environment
$user = ''
$pass = ''
$vbrServerIP = ":4443"
$mailFrom = "VeeamM365BackupEvents@onmicrosoft.com"
$mailAlerts = "VeeamM365BackupEvents@onmicrosoft.com"
$mailTo = "$mailFrom,$mailAlerts"
$groups = ("VeeamGroup")

#these values are the Display Names from the Microsoft 365 
$users = ("Alex Tan")


$header = @{
            "Content-Type" = "application/x-www-form-urlencoded"
            "accept"       = "application/json"
			}
			
$body = @{
            "grant_type"    = "password"
            "username"      = $user
            "password"      = $pass
			}

$url = "https://$vbrServerIP/v7/"
$veeamAPI = $url + "Token"


$tokenRequest = Invoke-RestMethod -Uri $veeamAPI -Headers $header -Body $body -Method Post -Verbose
Write-Output $tokenRequest.access_token
$bearer = $tokenRequest.access_token

$headers = @{
    "Authorization" = "Bearer $bearer"
}
#$response = Invoke-WebRequest -Uri "$url/Jobs" -Headers $headers

$organizations = Invoke-RestMethod -Method Get -Uri "$url/Organizations?extendedView=false" -Headers $headers
write-host($organizations.content)
$organisationName = $organizations.name
$organisationId = $organizations.Id

$auditItems = Invoke-WebRequest -Uri "$url/Organizations/$organisationId/AuditItems" -Headers $headers
write-host($auditItems.content)


$audititemsConverted = $auditItems | convertfrom-json

$itemsIds = $audititemsConverted[1].id
$response = Invoke-RestMethod -Uri "$url/Organizations/$organisationId/AuditItems/remove" `
    -Method Post `
    -Headers $headers `
    -ContentType "application/json" `
    -Body "{`n    `"itemIds`": [`n      `"$itemsIds`"`n    ]`n  }"