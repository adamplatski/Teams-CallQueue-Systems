#Connect-MicrosofTeams
$finance = "0000000-0000-0000-0000-000000000"
$PEOC = "0000000-0000-0000-0000-000000000"
$NOC = "0000000-0000-0000-0000-000000000"
$groupId = $PEOC
$members = Get-MgGroupMember -GroupId $groupId

#------------------------------------------------------------------

#$groupId = 'your-group-id-here'
$members = Get-MgGroupMember -GroupId $groupId

foreach ($member in $members)
{
    if ($member['@odata.type'] -eq "#microsoft.graph.user")
    {
        $user = Get-MgUser -UserId $member.Id
        Write-Host "User ID: " $member.Id "User Name: " $user.DisplayName
    }
}

#----------------------------------------------------------------------------------------

#$groupId = ''
$members = Get-MgGroupMember -GroupId $groupId
$userArray = @()

foreach ($member in $members){
    if ($member['@odata.type'] -eq "#microsoft.graph.user")
    {
        $user = Get-MgUser -UserId $member.Id
        $userObj = New-Object PSObject -Property @{
            UserID = $member.Id
            UserName = $user.DisplayName
        }
        $userArray += $userObj
    }
}

#----------------------------------------------------------------------------------------

# Print the userArray to check the output
$userArray | Format-Table -AutoSize

$policyName = "policy"

foreach ($user in $userArray) {

Grant-CsTeamsComplianceRecordingPolicy -Identity $user.UserID -PolicyName $policyName

    if ($user.TeamsComplianceRecordingPolicy.Name -eq '$policyName') {
        Write-Output ("User {0} has the 'test' TeamsComplianceRecordingPolicy." -f $user.UserPrincipalName)
    } else {
        Write-Output ("User {0} does not have the 'test' TeamsComplianceRecordingPolicy. Current Policy: {1}" -f $user.UserPrincipalName, $user.TeamsComplianceRecordingPolicy)
    }
}




