$ClientID = "0000000-0000-0000-0000-000000000"
$TenantID = "0000000-0000-0000-0000-000000000"
$CertificateThumbprint = "0000000-0000-0000-0000-000000000"
$teamId = "0000000-0000-0000-0000-000000000"
$userId = "0000000-0000-0000-0000-000000000"

$PEOnCallIdentity = "0000000-0000-0000-0000-000000000"
$PEOnCallTarget = "0000000-0000-0000-0000-000000000"
$PEOnCallL2Identity = "0000000-0000-0000-0000-000000000"
$PEOncallL2Target = "0000000-0000-0000-0000-000000000"
$PEOnCallL3Identity = "0000000-0000-0000-0000-000000000"
$PEOncallL3Target = "0000000-0000-0000-0000-000000000"

$adamsDRIdentity = "0000000-0000-0000-0000-000000000"
$adamsDRTarget = "0000000-0000-0000-0000-000000000"

$jonny = "0000000-0000-0000-0000-000000000"
$tommy = "0000000-0000-0000-0000-000000000"
$billy = "0000000-0000-0000-0000-000000000"
$jimmy = "0000000-0000-0000-0000-000000000"
$bobby = "0000000-0000-0000-0000-000000000"
$timmy = "0000000-0000-0000-0000-000000000"
$adam = "0000000-0000-0000-0000-000000000"

$bobbyTeams = "tel:+12345678910"
$bobbyCell = "+12345678910"
$timmyTeams = "tel:+12345678910"
$timmyCell = "+12345678910"
$tommyTeams = "tel:+12345678910"
$tommyCell = "+12345678910"
$jonnyTeams = "tel:+12345678910"
$jonnyCell = "+12345678910"
$jimmyTeams = "tel:+12345678910"
$jimmyCell = "+12345678910"
$adamTeams = "tel:+12345678910"
$adamCell = "+12345678910"
$billyTeams = "tel:+12345678910"

$bobbyEmail = "bobby@adams.io"
$billyEmail = "billy@adams.io"
$tommyEmail = "tommy@adams.io"
$timmyEmail = "timmy@adams.io"

$global:shiftsArray = @()
$global:resultz = @()
$global:matchingShift = @()
$global:restOfGroup = @()
$global:groupWorkers = @()
$global:currentDate = Get-Date

function connector {

    Connect-MgGraph -ClientID $ClientID -TenantID $TenantID -CertificateThumbprint $CertificateThumbprint
    Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ClientID -TenantID $TenantID
}

function shiftGetter {

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
    
    $global:matchingShift = $resultz | Where-Object { $currentDate -ge $_.StartTime -and $currentDate -le $_.EndTime }
    $global:matchingShiftArray = @($matchingShift)

    if ($matchingShiftArray.Count -eq 1) {
        $global:matchingShiftID = $matchingShift.UserID
    }
    elseif ($matchingShiftArray.Count -gt 1) {
        $global:multipleShiftIDs = $matchingShiftID | ForEach-Object { [System.Guid]::Parse($_) }
    }
    $matchingShift
}

function superAgentSorter {

    $global:groupWorkers = 
        
    @(
        "0000000-0000-0000-0000-000000000",
        "0000000-0000-0000-0000-000000000",
        "0000000-0000-0000-0000-000000000",
        "0000000-0000-0000-0000-000000000",
        "0000000-0000-0000-0000-000000000"
    )                     
    foreach ($i in 0..($groupWorkers.Length - 1)) {
        $groupWorkers[$i] = [System.Guid]::Parse($groupWorkers[$i])
    }
    
    $global:restOfGroup = $global:groupWorkers | Where-Object { $_ -ne $matchingShiftID }
    $groupWorkers = foreach ($worker in $groupWorkers) { [System.Guid]::Parse($worker) }
    $matchingIndex = $groupWorkers.IndexOf([System.Guid] $global:matchingShiftID)
    $global:reorderedWorkers = $groupWorkers[$matchingIndex..($groupWorkers.Count - 1)] + $groupWorkers[0..($matchingIndex - 1)]
    $global:reorderedWorkersString = foreach ($worker in $reorderedWorkers) { $worker.ToString() }
    $global:jimmyHandler = @($matchingShiftID) + $restOfGroup   
}

Function getCQassignee {

    if ($matchingShiftID -eq $adam) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        <#
        Set-CsUserCallingSettings -Identity "billyBob@adams.io" -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget $adamCell 
        Set-CsUserCallingSettings -Identity "billyBob@adams.io" -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        #>
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $adam -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 0 -TimeoutActionTarget $adamTeams -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $adam -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 15 -RoutingMethod Serial
    }

    elseif ($matchingShiftID -eq $jimmy) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        <#
        Set-CsUserCallingSettings -Identity "johnnyTom@adams.io" -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget $jimmyCell
        Set-CsUserCallingSettings -Identity "johnnyTom@adams.io" -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        #>
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 45 -TimeoutActionTarget $jimmyTeams -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $reorderedWorkersString -AgentAlertTime 30 -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget "tel:+12363614125" -TimeoutThreshold 150 -RoutingMethod Serial
    }

    elseif ($matchingShiftID -eq $bobby) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        #Set-CsUserCallingSettings -Identity $billyEmail -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget 
        #Set-CsUserCallingSettings -Identity $billyEmail -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 60 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $PEOncallL2Target -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $reorderedWorkersString -AgentAlertTime 60 -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $timmyTeams -TimeoutThreshold 60 -RoutingMethod Serial
    }

    elseif ($matchingShiftID -eq $timmy) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        <#
        Set-CsUserCallingSettings -Identity $timmyEmail -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget $timmyCell 
        Set-CsUserCallingSettings -Identity $timmyEmail -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        #>
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 0 -TimeoutActionTarget $timmyTeams -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $bobby -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $bobbyTeams -TimeoutThreshold 0 -RoutingMethod Attendant
        Set-CsCallQueue -Identity $PEOnCallL3Identity -Name "Network AnalystsL3" -Users $reorderedWorkersString -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 120 -RoutingMethod Attendant
    }

    elseif ($matchingShiftID -eq $bobby) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        <#
        Set-CsUserCallingSettings -Identity $bobbyEmail -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget $bobbyCell 
        Set-CsUserCallingSettings -Identity $bobbyEmail -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        #>
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 0 -TimeoutActionTarget $bobbyTeams -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $tommy -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $tommyTeams -TimeoutThreshold 0 -RoutingMethod Serial
        Set-CsCallQueue -Identity $PEOnCallL3Identity -Name "Network AnalystsL3" -Users $reorderedWorkersString -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 120 -RoutingMethod Attendant
    }

    elseif ($matchingShiftID -eq $tommy) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        <#
        Set-CsUserCallingSettings -Identity $tommyEmail -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget $tommyCell
        Set-CsUserCallingSettings -Identity $tommyEmail -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        #>
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 0 -TimeoutActionTarget $tommyTeams -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $reorderedWorkersString -AgentAlertTime 30 -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallL3Target -TimeoutThreshold 240 -RoutingMethod Serial
    }

    elseif ($matchingShiftID -eq $jonny) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they are scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        <#
        Set-CsUserCallingSettings -Identity $tommyEmail -IsForwardingEnabled $true -ForwardingType Simultaneous -ForwardingTargetType SingleTarget -ForwardingTarget $tommyCell
        Set-CsUserCallingSettings -Identity $tommyEmail -IsUnansweredEnabled $true -UnansweredTargetType SingleTarget -UnansweredTarget "adams_PlatformEngineers_CQ_RA_DR@adams.io" -UnansweredDelay 00:00:45 
        #>
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 60 -TimeoutActionTarget $adamsDRTarget -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $bobby -AgentAlertTime 30 -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $bobbyTeams -TimeoutThreshold 0 -RoutingMethod Serial
        Set-CsCallQueue -Identity $PEOnCallL3Identity -Name "Network AnalystsL3" -Users $reorderedWorkersString -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 120 -RoutingMethod Attendant
    }

    elseif ($matchingShiftArray.Count -eq 1) {
        Write-Output "`nThe matching employee's UserID for $currentDate is $($matchingShift.UserID) and they're scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShiftID -AgentAlertTime 45 -TimeoutAction Forward -TimeoutThreshold 45 -TimeoutActionTarget $adamsDRTarget -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $reorderedWorkersString -AgentAlertTime 45  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 150 -RoutingMethod Serial
    }

    elseif ($multipleShiftIDs) {
        Write-Output "`nThe matching employee's UserID's for $currentDate are $($matchingShift.UserID) and they're scheduled from $($matchingShift.StartTime) to $($matchingShift.EndTime)"
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $matchingShift.UserID -AgentAlertTime 30 -TimeoutAction Forward -TimeoutThreshold 120 -TimeoutActionTarget $PEOncallL2Target -RoutingMethod Serial -AllowOptOut $false -UseDefaultMusicOnHold $true -OverflowAction Forward -OverflowActionTarget $PEOncallL2Target -OverflowThreshold 200 -ConferenceMode $true -LanguageID "en-US"
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $restOfGroup -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 150 -RoutingMethod Serial
    }
    else {
        Set-CsCallQueue -Identity $PEOnCallIdentity -Name "Network Analysts" -Users $groupWorkers -AgentAlertTime 30 -TimeoutAction Forward -TimeoutThreshold 45 -TimeoutActionTarget $PEOncallL2Target -RoutingMethod Serial
        Set-CsCallQueue -Identity $PEOnCallL2Identity -Name "Network AnalystsL2" -Users $restOfGroup -AgentAlertTime 30  -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Forward -TimeoutActionTarget $PEOnCallTarget -TimeoutThreshold 150 -RoutingMethod Serial
    }
}

connector
shiftGetter
shiftParser
dateMatcher 
superAgentSorter
getCQassignee

