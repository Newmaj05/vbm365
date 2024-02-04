<#
.SYNOPSIS
    Enable and Add Microsoft 365 email notifications for Audit purposes
.DESCRIPTION
	#For email notifications for Audit alerts, using this API appears to be the working method - https://helpcenter.veeam.com/docs/vbo365/rest/reference/vbo365-rest.html?ver=70#tag/AuditEmailSettings
	
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

$emailType = "Microsoft365"
$redirectUrl = "http://localhost"


$subject = "VBM365 Audit Alert - %StartTime% — %OrganizationName% - %DisplayName% - %Action% - %InitiatedByUserName%"

$apiUrl = "$url/AuditEmailSettings/PrepareOAuthSignIn"
$headers = @{
    "Content-Type" = "application/json"
    "Accept" = "application/json"
    "Authorization" = "Bearer $bearer"
}
$body = @{
    "authenticationServiceKind" = $emailType
    "redirectUrl" = $redirectUrl
}

$response = Invoke-RestMethod -Method Post -Uri $apiUrl -Body ($body | ConvertTo-Json) -Headers $headers 
$signInUrl = $response.signInUrl


$prefix = 'http://localhost/'
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Start-Process $signInUrl
Start-Sleep -Seconds 30
$context = $listener.GetContext()
$requestUrl = $context.Request.Url
$listener.Stop()
$params = [System.Web.HttpUtility]::ParseQueryString($requestUrl.Query)
$code = $null
$state = $null
$params.AllKeys | ForEach-Object { 
    if($_ -eq 'code') {
        $code = $params[$_]
    }
    if($_ -eq 'state') {
        $state = $params[$_]
    }
}

$apiUrl = "$url/AuditEmailSettings/CompleteOAuthSignIn"
$headers = @{
    "Content-Type" = "application/json"
    "Accept" = "application/json"
    "Authorization" = "Bearer $bearer"
}
$body = @{
    "code" = $code
    "state" = $state
}
$response = Invoke-RestMethod -Method Post -Uri $apiUrl -Body ($body | ConvertTo-Json) -Headers $headers 
$requestId = $response.requestId
$userId = $response.userId

$apiUrl = "$url/AuditEmailSettings"
$headers = @{
    "Content-Type" = "application/json"
    "Accept" = "application/json"
    "Authorization" = "Bearer $bearer"
}


$body = @{
	"enableNotification" = $true
	"from" = $mailFrom
	"to" = $mailTo
	"subject" = $subject
	"authenticationType" = $emailType
	"userId" = $userId
	"requestId" = $requestId
}
$response = Invoke-RestMethod -Method Put -Uri $apiUrl -Body ($body | ConvertTo-Json) -Headers $headers 


$apiUrl = "$url/AuditEmailSettings"
$headers = @{
    "Accept" = "application/json"
    "Authorization" = "Bearer $bearer"
}
$response = Invoke-RestMethod -Method Get -Uri $apiUrl -Headers $headers

Write-Host "This is the configuration that we have pushed to VB365"
$response


$apiUrl = "$url/AuditEmailSettings/SendTest"
$headers = @{
    "Accept" = "application/json"
    "Authorization" = "Bearer $bearer"
}
$response = Invoke-RestMethod -Method Post -Uri $apiUrl -Headers $headers


$AuditEmailSettings = Invoke-WebRequest -Uri "$url/AuditEmailSettings" -Headers $headers
write-host($AuditEmailSettings.content)

$organizations = Invoke-RestMethod -Method Get -Uri "$url/Organizations?extendedView=false" -Headers $headers
write-host($organizations.content)
$organisationName = $organizations.name
$organisationId = $organizations.Id


$entities = Invoke-RestMethod -Method Get -Uri "$url/Organizations/$organisationId/Users?limit=10000" -Headers $headers

#$user1 = $entities.results |where-object -property displayName -eq "Andrew Newman"
#$user2 = $entities.results |where-object -property displayName -eq "General Admin"

$auditusers = @()
foreach ($user in $users) {
	$user1 = $entities.results |where-object -property displayName -eq $user
	$auditusers += $user1
}

#$users = ($user1,$user2)

$userDetails = $auditusers | ForEach-Object {
        @{
            id          = $_.id
            displayName = $_.displayName
            name        = $_.name
        }
    }
	
	
$headers = @{
	"Content-Type" = "application/json"
	"Accept" = "application/json"
	"Authorization" = "Bearer $bearer"
}


$userDetails | ForEach-Object {
        $bodyData = @{
            type = "user"
            user = @{
                id          = $_.id
                displayName = $_.displayName
                name        = $_.name
            }
        }

        $bodyJson = ConvertTo-Json @($bodyData)

try {
            $response = Invoke-RestMethod -Method Post -Uri "$url/Organizations/$organisationId/AuditItems" -Headers $headers -Body $bodyJson
            Write-Host "Successfully added user $($_.name) to the Audit." -ForegroundColor Green
        } catch {
            Write-Host "Failed to add user $($_.name). Details: $($_.Exception.Message)" -ForegroundColor Red
        }
}

$auditEmailSettings = Invoke-RestMethod -Uri "$url/AuditEmailSettings" -Headers $headers
write-host($auditEmailSettings)
#$auditItems = Invoke-WebRequest -Uri "$url/Organizations/$organisationId /AuditItems" -Headers $headers
$auditItems = Invoke-RestMethod -Uri "$url/Organizations/$organisationId /AuditItems" -Headers $headers
write-host($auditItems)

