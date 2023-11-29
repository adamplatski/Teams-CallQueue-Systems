$ClientID = "0000000-0000-0000-0000-000000000"
$TenantID = "0000000-0000-0000-0000-000000000"
$CertificateThumbprint = "0000000-0000-0000-0000-000000000"
$teamId = "0000000-0000-0000-0000-000000000"
$userId = "0000000-0000-0000-0000-000000000"
$teamsChannelID = "0000000-0000-0000-0000-000000000"
$teamsTeamID = "0000000-0000-0000-0000-000000000"
$userEmail = "adam@adams.io"

$global:NOCgroupIds = @("0000000-0000-0000-0000-000000000", "0000000-0000-0000-0000-000000000", "0000000-0000-0000-0000-000000000")
$global:currentDate = (Get-Date)
$global:shiftsArray = @()
$global:resultz = @()
$global:matchingShift = @()
$global:NOCids = @()

function NOCconnector {

    Connect-MgGraph -ClientID $ClientID -TenantID $TenantID -CertificateThumbprint $CertificateThumbprint
    Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ClientID -TenantID $TenantID
    "Connecting to Tenant"
}

function schedulingGroup {

    "Collecting NOC users"
    #$schedulingGroup = Get-MgTeamScheduleSchedulingGroup -TeamId $teamId -SchedulingGroupId $schedulingGroupId
    #$global:groupWorkers = @($schedulingGroup.UserIds)
    $NOCids = foreach ($schedulingGroupId in $NOCgroupIds) {
        Invoke-MgGraphRequest -Uri "v1.0/teams/$teamId/schedule/schedulingGroups/$schedulingGroupId" -Headers @{ "MS-APP-ACTS-AS" = $userId }
    }
    $global:NOCusers = $NOCids.userIds 
    $global:agentPrioritizer = $NOCusers | Where-Object { $_ -notin $jordan, $bruce, $nathan, $wilson }
}

function shiftGetter {

    "Collecting all NOC shifts"
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
            UPN       = $UPN
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
        $user = Get-MgUser -UserId $_.UserID | Select-Object DisplayName, UserPrincipalName
        $_ | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $user.DisplayName -Force
        $_ | Add-Member -MemberType NoteProperty -Name "UPN" -Value $user.UserPrincipalName -Force
    }
}

function messageSender {
    foreach ($shift in $matchingShift) {
        $shiftUPN = $shift.UPN
        $shiftDN = $shift.DisplayName
        $startWindowEnd = $shift.StartTime.AddMinutes(10)  
        $endWindowStart = $shift.EndTime.AddMinutes(-10)  

        if ($currentDate -le $startWindowEnd -and $currentDate -ge $shift.StartTime) {
            # Send an email to users who just started their shift
            $htmlBody = "<p>Hello $shiftDN,<br><br>Please remember to change your work-mode to logged in.<br><br>Thank you</p>"
            $subject = "login reminder"

        } elseif ($currentDate -ge $endWindowStart -and $currentDate -le $shift.EndTime) {
            # Send an email to users who are about to end their shift
            $htmlBody = "<p>Hello $shiftDN,<br><br>Please remember to change your work mode to logged-off at the end of your shift<br><br>Thank you</p>"
            $subject = "Tendfor work-mode logout reminder"

        } else {
            # Skip sending email if neither condition is met
            continue
        }

        # Common email parameters
        $params = @{
            Message = @{
                Subject = $subject
                Body = @{
                    ContentType = "HTML"
                    Content = $htmlBody
                }
                ToRecipients = @(
                    @{
                        EmailAddress = @{
                            Address = $shiftUPN
                        }
                    }
                )
            }
            SaveToSentItems = "true"
        }

        # Send the email
        Send-MgUserMail -UserId $userId -BodyParameter $params
    }
}

NOCconnector
schedulingGroup
shiftGetter
shiftParser
dateMatcher
messageSender
<#
function messageSender {

    if ($matchingShiftID) {
        foreach ($matchedUser in $matchingShift) {

            $global:matchingShiftUPN = $matchingShift.UPN
            $global:matchingShiftDN = $matchingShift.DisplayName
            $matchingShiftUPN

            $startWindowEnd = $matchingShift.StartTime.AddMinutes(10)  # 10 minutes after StartDateTime
            $endWindowStart = $matchingShift.EndTime.AddMinutes(-10)   # 10 minutes before EndDateTime

            if ($currentDate -le $startWindowEnd -and $currentDate -ge $matchingShift.StartTime) {
                #Hello $matchingShiftDisplayName please remember to change your Tendfor work mode to logged-in at the beginning of your shift
                "Sending Email to users that just started their shift:
            $matchingShiftDN"
                $htmlBody = "<p>Hello $matchingShiftDN,<br><br>Please remember to change your work-mode to logged in.<br><br>Thank you</p>"

                $params = @{
                    Message         = @{
                        Subject      = "login reminder"
                        Body         = @{
                            ContentType = "HTML"  # or "Text" if you prefer plain text
                            Content     = $htmlBody
                        }
                        ToRecipients = @(
                            @{
                                EmailAddress = @{
                                    Address = "adam.adams@adams.io"
                                }
                            }
                        )
                    }
                    SaveToSentItems = "true"
                }
        
                # Send the email
                Send-MgUserMail -UserId $userId -BodyParameter $params
            }
            elseif ($currentDate -ge $endWindowStart -and $currentDate -le $matchingShift.EndTime) {

                $htmlBody = "<p>" + "Hello $matchingShiftDN,<br><br>Please remember to change your work mode to logged-off at the end of your shift<br><br>" + "</p>"
                "Sending an email to users that are about to end their shift:
            $matchingShiftDN"
                $params = @{
                    Message         = @{
                        Subject      = "Tendfor work-mode logout reminder "
                        Body         = @{
                            ContentType = "HTML"  # or "Text" if you prefer plain text
                            Content     = $htmlBody
                        }
                        ToRecipients = @(
                            @{
                                EmailAddress = @{
                                    Address = "adam.adams@adams.io"
                                }
                            }
                        )
                    }
                    SaveToSentItems = "true"
                }
            
                # Send the email
                Send-MgUserMail -UserId $userId -BodyParameter $params
            }
            else { "No one to email" }
        }
    }
}
#>

#$channelMemberIDs = (Get-MgTeamChannelMember -TeamId $teamsTeamID -ChannelId $teamsChannelID)


