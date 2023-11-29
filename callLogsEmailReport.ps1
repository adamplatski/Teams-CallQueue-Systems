$UserId = ""
$ClientID = ""
$TenantID = ""
$CertificateThumbprint = ""
$ClientSecret = ""

Connect-MgGraph -ClientID $ClientID -TenantID $TenantID -CertificateThumbprint $CertificateThumbprint
Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ClientID -TenantID $TenantID

$techSupportBlock = {

    $PWord = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
    $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientID, $PWord
    $global:accessToken = Get-GraphApiAccessToken -Credential $Credential -TenantId $TenantID

    $yesterdaysDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
    $todaysDate = (Get-Date).AddDays(+1).ToString("yyyy-MM-dd")
    $techAttendant = @('test@test.com')
    $techSupportCalls = Get-TeamsPstnCalls -StartDate $yesterdaysDate -EndDate $todaysDate -AccessToken $accessToken | Where-Object -Property userPrincipalName -In $techAttendant
    #$financeCalls = Get-TeamsPstnCalls -StartDate $yesterdaysDate -EndDate $todaysDate -AccessToken $accessToken | Where-Object -Property userPrincipalName -In $financeAttendant
    #$techSupportCalls = Get-TeamsPstnCalls -StartDate $yesterdaysDate -EndDate $todaysDate -AccessToken $accessToken | Where-Object -Property callType -EQ -Value 'ucap_in'
    
    $techSupportCalls | Select-Object userDisplayName, calleeNumber, callerNumber, `
    @{
        Name='StartDateTime';
        Expression={([System.DateTimeOffset]::Parse($_.StartDateTime) - [System.TimeSpan]::FromHours(7)).ToString("yyyy-MM-dd HH:mm:ss") + " PST"}
    },
    @{
        Name='EndDateTime';
        Expression={([System.DateTimeOffset]::Parse($_.EndDateTime) - [System.TimeSpan]::FromHours(7)).ToString("yyyy-MM-dd HH:mm:ss") + " PST"}
    },
    @{
        Name='duration';
        Expression={
            $durationInSeconds = $_.duration
            if ($durationInSeconds -ne $null) {
                $minutes = [math]::Floor($durationInSeconds / 60)
                $seconds = $durationInSeconds % 60
                $formattedDuration = "{0:D2}:{1:D2}" -f [int]$minutes, [int]$seconds
                $formattedDuration
            } else {
                "Duration calculation did not work"
            }
        }
    }
}

$techSupportOutput = Invoke-Command -ScriptBlock $techSupportBlock
$techSupportOutput
$techSupportBody = $techSupportOutput | ConvertTo-Html -As Table -Property userDisplayName, calleeNumber, callerNumber, startDateTime, endDateTime, duration -Head "
<style>
    table {
        border-collapse: collapse;
        width: 100%;
    }
    th, td {
        padding: 8px;
        text-align: left;
        border-bottom: 1px solid #ddd;
    }
    th {
        background-color: #3b3ef5;
        color: white;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
</style>" 

$allCallBlock = {

    $yesterdaysDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
    $todaysDate = (Get-Date).AddDays(+1).ToString("yyyy-MM-dd")
    $allCalls = Get-TeamsPstnCalls -StartDate $yesterdaysDate -EndDate $todaysDate -AccessToken $accessToken | Where-Object -Property callType -EQ -Value 'ucap_in'

    $allCalls | Select-Object userDisplayName, calleeNumber, callerNumber, `
    @{
        Name='StartDateTime';
        Expression={([System.DateTimeOffset]::Parse($_.StartDateTime) - [System.TimeSpan]::FromHours(7)).ToString("yyyy-MM-dd HH:mm:ss") + " PST"}
    },
    @{
        Name='EndDateTime';
        Expression={([System.DateTimeOffset]::Parse($_.EndDateTime) - [System.TimeSpan]::FromHours(7)).ToString("yyyy-MM-dd HH:mm:ss") + " PST"}
    },
    @{
        Name='duration';
        Expression={
            $durationInSeconds = $_.duration
            if ($durationInSeconds -ne $null) {
                $minutes = [math]::Floor($durationInSeconds / 60)
                $seconds = $durationInSeconds % 60
                $formattedDuration = "{0:D2}:{1:D2}" -f [int]$minutes, [int]$seconds
                $formattedDuration
            } else {
                "Duration calculation did not work"
            }
        }
    }
}

$allCallsOutput = Invoke-Command -ScriptBlock $allCallBlock
$allCallsOutput
$allCallsBody = $allCallsOutput | ConvertTo-Html -As Table -Property userDisplayName, calleeNumber, callerNumber, startDateTime, endDateTime, duration -Head "
<style>
    table {
        border-collapse: collapse;
        width: 100%;
    }
    th, td {
        padding: 8px;
        text-align: left;
        border-bottom: 1px solid #ddd;
    }
    th {
        background-color: #3b3ef5;
        color: white;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
</style>" 

$techSupportParams = @{
    Message         = @{
        Subject      = "Tech Support PSTN Report"
        Body         = @{
            ContentType = "HTML"
            Content     =  "$techSupportBody"
        }
       ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = ""
                }
            },
            @{
                EmailAddress = @{
                    Address = ""
                }
            }
        )
    }
    SaveToSentItems = "true"
}


Send-MgUserMail -UserId $userId -BodyParameter $techSupportParams

$allCallsParams = @{
    Message         = @{
        Subject      = "TFN PSTN Report"
        Body         = @{
            ContentType = "HTML"
            Content     =  "$allCallsBody"
        }
       ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = ""
                }
            }
        )
    }
    SaveToSentItems = "true"
}

Send-MgUserMail -UserId $userId -BodyParameter $allCallsParams