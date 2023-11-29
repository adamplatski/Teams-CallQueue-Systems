$ClientID = "0000000-0000-0000-0000-000000000"
$TenantID = "0000000-0000-0000-0000-000000000"
$CertificateThumbprint = "0000000-0000-0000-0000-000000000"
$teamId = "0000000-0000-0000-0000-000000000"
$userId = "0000000-0000-0000-0000-000000000"
$Techidentity = "0000000-0000-0000-0000-000000000"
$Techtarget = "0000000-0000-0000-0000-000000000"
$timmylineIdentity = "0000000-0000-0000-0000-000000000"
$timmylineTarget = "0000000-0000-0000-0000-000000000"
$TechidentityL2 = "0000000-0000-0000-0000-000000000"
$TechtargetL2 = "0000000-0000-0000-0000-000000000"
$billy = "0000000-0000-0000-0000-000000000"
$bobby = "0000000-0000-0000-0000-000000000"
$jimmy = "0000000-0000-0000-0000-000000000"
$timmy = "0000000-0000-0000-0000-000000000"
$tommy = "0000000-0000-0000-0000-000000000"
$global:TechgroupIds = @("ABC", "DEF", "GHI")
$global:currentDate = (Get-Date)
$global:shiftsArray = @()
$global:resultz = @()
$global:matchingShift = @()
$global:restOfGroup = @()
$global:groupWorkers = @()
$global:Techids = @()

function Techconnector {

    Connect-MgGraph -ClientID $ClientID -TenantID $TenantID -CertificateThumbprint $CertificateThumbprint
    Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ClientID -TenantID $TenantID
    "Connecting to Tenant"
}

function schedulingGroup {

    "Collecting users"
    #$schedulingGroup = Get-MgTeamScheduleSchedulingGroup -TeamId $teamId -SchedulingGroupId $schedulingGroupId
    #$global:groupWorkers = @($schedulingGroup.UserIds)
    $Techids = foreach ($schedulingGroupId in $TechgroupIds) {
        Invoke-MgGraphRequest -Uri "v1.0/teams/$teamId/schedule/schedulingGroups/$schedulingGroupId" -Headers @{ "MS-APP-ACTS-AS" = $userId }
    }
    $global:Techusers = $Techids.userIds 
    $global:agentPrioritizer = $Techusers | Where-Object { $_ -notin $jimmy, $billy, $bobby, $tommy }
}

function shiftGetter {

    "Collecting all shifts"
    #$teamScheduleShiftIds = @(Get-MgTeamScheduleShift -TeamId $teamId -All)
    $teamScheduleShiftIds = Invoke-MgGraphRequest -Uri "v1.0/teams/$teamId/schedule/shifts/" -Headers @{ "MS-APP-ACTS-AS" = $userId }
    $shiftAgent = $teamScheduleShiftIds.value.Id

    foreach ($shiftID in $shiftAgent) {
        $shiftz = @(Invoke-MgGraphRequest -Uri "v1.0/teams/$teamId/schedule/shifts/$shiftID" -Headers @{ "MS-APP-ACTS-AS" = $userId })
        $sharedShift = $shiftz.SharedShift
        $shiftUserId = $shiftz.userId
        $userSchedule = @($shiftUserId, $sharedShift.startDateTime, $sharedShift.EndDateTime) 
        $global:shiftsArray += , @($userSchedule) 
    } 
}

function shiftParser {
    
    "Creating individual PS object for each shift"
    $global:resultz = $shiftsArray | ForEach-Object {
        $UserId = $_[0]
        $startTime = [DateTime]::ParseExact($_[1], "MM/dd/yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
        $endTime = [DateTime]::ParseExact($_[2], "MM/dd/yyyy HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
    
        [PSCustomObject]@{
            UserID    = $UserId
            StartTime = $startTime
            EndTime   = $endTime
        }    
    }
}

function dateMatcher {
    "Finding current shifts"
    $global:matchingShift = $resultz | Where-Object { $currentDate -ge $_.StartTime -and $currentDate -le $_.EndTime }
    $global:matchingShiftArray = @($matchingShift)
    $global:matchingShiftID = $matchingShift.UserID

    $global:matchingShift | ForEach-Object {
        $user = Get-MgUser -UserId $_.UserID | Select-Object -ExpandProperty DisplayName
        $_ | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $user -Force
    }

    $global:tommyRemover = $matchingShift.UserID | Where-Object { $_ -notin $jimmy }
    $global:johnnyBobbyRemover = $matchingShift.UserID | Where-Object { $_ -notin $jimmy, $timmy }
    $global:jimmyShifts = $matchingShift | Where-Object {$_.UserID -eq $timmy}
    $global:frontLineDisplayer = $matchingShift | Where-Object { $_.UserID -notin $timmy, $jimmy }
    $global:2ndLineDisplayer = $matchingShift | Where-Object { $_.UserID -eq $timmy }
    #$global:nightLineDisplayer = $matchingShift | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $matchingShiftNames.DisplayName -Force -PassThru

    if ($matchingShiftArray.Count -eq 1) {
        $global:matchingShiftID = $matchingShift.UserID
    }
    elseif ($matchingShiftArray.Count -gt 1) {
        $global:multipleShiftIDs = $matchingShift.UserId | ForEach-Object { [System.Guid]::Parse($_) }
    }
}

function daytablePrinter {

    Write-Output "Current DateTime: $currentDate"
    Write-Output `n"--- Matching Day Shifts ---" 
    $matchingShift
    Write-Output `n"--- Tech FrontLine Active Agents ---"`n
    $frontLineDisplayer
    Write-Output `n"--- Tech 2ndLine Active Agents ---"`n
    $2ndLineDisplayer

}

function nighttablePrinter {

    Write-Output "Current DateTime: $currentDate"
    Write-Output `n"--- Matching Night Shifts ---" 
    $matchingShift
    Write-Output `n"--- Tech Active Agents ---" 
    $matchingShift
}

function dayshiftCQassignee {
    
    "`nAssigning dayTech agents to their respective queue's"

    if ($timmyShifts -and $jimmyTimRemover.Count -eq 1) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $jimmyTimRemover -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $timmylineTarget -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $timmylineIdentity -Name "Support-CallQueue" -Users $timmy -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant -AllowOptOut $false    
    }

    elseif ($timmyShifts -and $jimmyTimRemover.Count -eq 2) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $jimmyTimRemover -AgentAlertTime 20 -TimeoutAction Forward -TimeoutThreshold 80 -TimeoutActionTarget $timmylineTarget -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $timmylineIdentity -Name "Support-CallQueue" -Users $timmy -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant -AllowOptOut $false    
    }

    elseif ($timmyShifts -and $jimmyTimRemover.Count -eq 3) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $jimmyTimRemover -AgentAlertTime 20 -TimeoutAction Forward -TimeoutThreshold 120 -TimeoutActionTarget $timmylineTarget -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $timmylineIdentity -Name "Support-CallQueue" -Users $timmy -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant -AllowOptOut $false    
    }

    elseif ($timmyShifts -and $jimmyTimRemover.Count -ge 4) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users 8b12e79d-02cc-4846-8d6b-1484e3a78900 -AgentAlertTime 20 -TimeoutAction Forward -TimeoutThreshold 160 -TimeoutActionTarget $timmylineTarget -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $timmylineIdentity -Name "Support-CallQueue" -Users $timmy -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant -AllowOptOut $false    
    }

    elseif ($matchingShiftArray.Count -eq 1) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $matchingShiftID -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $TechtargetL2 -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod RoundRobin
    }
    
    elseif ($multipleShiftIDs) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $jimmyRemover -AgentAlertTime 30 -TimeoutAction Forward -TimeoutThreshold 180 -TimeoutActionTarget $TechtargetL2 -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod RoundRobin
    }

    else {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $agentPrioritizer -AgentAlertTime 30 -TimeoutAction Forward -TimeoutThreshold 300 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Attendant
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant
    } 
}

function nightshiftCQassignee {

    "`nAssigning nightTech agents to their respective queue's"
    if ($matchingShiftArray.Count -eq 1) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $matchingShift.UserID -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant -AllowOptOut $false
    }

    elseif ($multipleShiftIDs) {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $matchingShiftID -AgentAlertTime 30 -TimeoutAction Forward -TimeoutThreshold 120 -TimeoutActionTarget $TechtargetL2 -RoutingMethod RoundRobin -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $TechtargetL2 -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant -AllowOptOut $false
    }
    else {
        Set-CsCallQueue -Identity $Techidentity -Name "Tech-CallQueue" -Users $agentPrioritizer -AgentAlertTime 30 -TimeoutAction Forward -TimeoutThreshold 240 -TimeoutActionTarget $TechtargetL2 -RoutingMethod Attendant
        Set-CsCallQueue -Identity $TechidentityL2 -Name "Tech-CallQueueL2" -Users $agentPrioritizer -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $Techtarget -TimeoutThreshold 240 -RoutingMethod Attendant
    }
}

Techconnector
schedulingGroup
shiftGetter
shiftParser
dateMatcher
$currentHour = (Get-Date).Hour

switch ($currentHour) {
    { $_ -ge 15 -and $_ -lt 23 } {
        daytablePrinter
        dayshiftCQassignee
        break
    }
    default {
        nighttablePrinter
        nightshiftCQassignee
        break
    }
}
