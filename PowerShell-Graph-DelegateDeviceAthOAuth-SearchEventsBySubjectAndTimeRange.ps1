# PowerShell-Graph-DelegateDeviceAthOAuth-SearchEventsBySubjectAndTimeRange.ps1
# This script demonstrates how to use the Microsoft Graph API with delegated OAuth authentication to search for calendar events based on a subject and a time range.
# Prerequisites:
# - Register an application in Azure AD and grant it the necessary permissions (e.g., Calendars.Read)
# - Obtain the client ID and tenant ID for your registered application  
# Note: This script uses interactive authentication, which is suitable for testing and development. For production scenarios, consider using a more secure authentication method (e.g., certificate-based or client secret) and running the script in a non-interactive environment.        
 
# ================================
# CONFIGURATION
# ================================
$TenantId  = "YOUR_TENANT_ID"
$ClientId  = "YOUR_CLIENT_ID"

# Date range (UTC recommended)
$StartDateTime = "2020-01-01T00:00:00Z"  # UTC time
$EndDateTime   = "2020-01-10T23:59:59Z"  # UTC time

# Subject search string
$SubjectFilter = "Meeting"

# ================================
# STEP 1: DEVICE CODE AUTH (delegated)
# ================================
$DeviceCodeRequest = @{
    client_id = $ClientId
    scope     = "https://graph.microsoft.com/Calendars.Read"
}

$DeviceCodeResponse = Invoke-RestMethod -Method Post `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $DeviceCodeRequest

Write-Host ""
Write-Host "======================================="
Write-Host $DeviceCodeResponse.message
Write-Host "======================================="
Write-Host ""

# ================================
# STEP 2: POLL FOR TOKEN
# ================================
$TokenBody = @{
    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
    client_id   = $ClientId
    device_code = $DeviceCodeResponse.device_code
}

do {
    Start-Sleep -Seconds $DeviceCodeResponse.interval

    try {
        $TokenResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -ContentType "application/x-www-form-urlencoded" `
            -Body $TokenBody

        $AccessToken = $TokenResponse.access_token
        $TokenReady = $true
    }
    catch {
        $TokenReady = $false
    }

} while (-not $TokenReady)

Write-Host "Access token acquired."

# ================================
# SET HEADERS
# ================================
$Headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Content-Type"  = "application/json"
}

# ================================
# BUILD REQUEST (calendarView + filter)
# ================================
$Url = "https://graph.microsoft.com/v1.0/me/calendarView" +
       "?startDateTime=$StartDateTime&endDateTime=$EndDateTime" +
       "&`$select=id,subject,start,end" +
       "&`$filter=contains(subject,'$SubjectFilter')"

Write-Host "Querying events..."

# ================================
# GET EVENTS (paging)
# ================================
$AllEvents = @()

do {
    $Response = Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers

    if ($Response.value) {
        $AllEvents += $Response.value
    }

    $Url = $Response.'@odata.nextLink'

} while ($Url)

Write-Host "Total matching events: $($AllEvents.Count)"

# ================================
# OUTPUT RESULTS
# ================================
foreach ($Event in $AllEvents) {

    Write-Host "-------------------------------------"
    Write-Host "ID      : $($Event.id)"
    Write-Host "Subject : $($Event.subject)"
    Write-Host "Start   : $($Event.start.dateTime)"
    Write-Host "End     : $($Event.end.dateTime)"
}

Write-Host "Done."
