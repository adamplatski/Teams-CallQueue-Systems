# Login
## You will be prompted to enter your Teams administrator credentials.

$credential = Get-Credential
Connect-MicrosoftTeams -Credential $credential
Connect-MsolService -Credential $credential

# Create After Hours Schedules
$timerangeMoFr = New-CsOnlineTimeRange -Start 08:30 -end 17:00 
$afterHoursSchedule = New-CsOnlineSchedule -Name "After Hours Schedule" -WeeklyRecurrentSchedule -MondayHours @($timerangeMoFr) -TuesdayHours @($timerangeMoFr) -WednesdayHours @($timerangeMoFr) -ThursdayHours @($timerangeMoFr) -FridayHours @($timerangeMoFr) -Complement

# Create Address and Email Information Prompt
$addressPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "To repeat this information at any time press the * key. Our mailing address is: 100 Timson street, St John's, Newfoundland, Canada."
$addressPromptOption = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone4 -Prompt $addressPrompt

$financeGroupID = Find-CsGroup -SearchQuery "Finance" | ForEach-Object { $_.Id }
$supportGroupID = Find-CsGroup -SearchQuery "Support" | ForEach-Object { $_.Id }

# Dial By Name Auto Attendant - Resource Account Creation
## Note: Creating resource account here so it can be referenced on the main auto attendant. The actual Dial By Name auto attendant will be created later.
# Get license types
Get-MsolAccountSku
<#
#>

# -- Auto Attendant: ce933385-9390-45d1-9512-c8d228074e07
# -- Call Queue: 11cd3e2e-fccb-42ad-ad00-878b93575e07

# Adam's DBN Auto Attendant
New-CsOnlineApplicationInstance -UserPrincipalName adam_DialByName_RA@adam.io -DisplayName "Adam's Dial By Name RA" -ApplicationID "0000000-0000-0000-0000-000000000" 
Set-MsolUser -UserPrincipalName "adam_DialByName_RA@adam.io" -UsageLocation CA 
Set-MsolUserLicense -UserPrincipalName "adam_DialByName_RA@adam.io" -AddLicenses "reseller-account:PHONESYSTEM_VIRTUALUSER" 
$dialByNameApplicationInstanceID = (Get-CsOnlineUser "adam_DialByName_RA@adam.io").Identity
$dialByNameTarget = (Get-CsOnlineUser -Identity "adam_DialByName_RA@adam.io")
$dialByNameEntity = New-CsAutoAttendantCallableEntity -Identity $dialByNameTarget.Identity -Type applicationendpoint
$dialByNameOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $dialByNameEntity
$dialByNameMenuOption9Target = (Get-CsOnlineUser "adam_DialByName_RA@adam.io").Identity
$dialByNameMenuOption9Entity = New-CsAutoAttendantCallableEntity -Identity $dialByNameMenuOption9Target -Type applicationendpoint
$dialByNameMenuOption9 = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone9 -CallTarget $dialByNameMenuOption9Entity
$dialByNameMenuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Please say or enter the name of the person you would like to reach. To return to the previous menu press 9"
$dialByNameMenu = New-CsAutoAttendantMenu -Name "Adam's Dial By Name Menu" -MenuOptions $dialByNameMenuOption9 -Prompt $dialByNameMenuPrompt -EnableDialByName -DirectorySearchMethod ByName
$dialByNameCallFlow = New-CsAutoAttendantCallFlow -Name "Adam's Dial By Name Call Flow" -Menu $dialByNameMenu
$dialScope = New-CsAutoAttendantDialScope -GroupScope -GroupIds @("0000000-0000-0000-0000-000000000")
$dialByNameAutoAttendant = New-CsAutoAttendant -Name "Adam's Dial By Name" -DefaultCallFlow $dialByNameCallFlow -LanguageId "en-US" -TimeZoneId "UTC" -Operator $dialByNameEntity -EnableVoiceResponse -InclusionScope $dialScope 
New-CsOnlineApplicationInstanceAssociation -Identities @($dialByNameApplicationInstanceID) -ConfigurationID $dialByNameAutoAttendant.Id -ConfigurationType AutoAttendant

# Create and assign Resource Account
New-CsOnlineApplicationInstance -UserPrincipalName adam_Main_RA@adam.io -DisplayName "Adam's Main RA" -ApplicationID "0000000-0000-0000-0000-000000000"
Set-MsolUser -UserPrincipalName "adam_Main_RA@adam.io" -UsageLocation CA 
Set-MsolUserLicense -UserPrincipalName "adam_Main_RA@adam.io" -AddLicenses "reseller-account:PHONESYSTEM_VIRTUALUSER" 
$mainInstanceID = (Get-CsOnlineUser -Identity "adam_Main_RA@adam.io").Identity 
$mainEntity = New-CsAutoAttendantCallableEntity -Identity $mainInstanceID -Type ApplicationEndpoint 

$mainMenuSupportOptionTarget = (Get-CsOnlineUser -Identity "Support_CQ_RA@adam.io").Identity
$mainMenuSupportOptionEntity = New-CsAutoAttendantCallableEntity -Identity $mainMenuSupportOptionTarget -Type ApplicationEndpoint
$mainMenuSupportOption1 = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $mainMenuSupportOptionEntity
$mainMenuFinanceOptionTarget = (Find-CsGroup -SearchQuery "Finance").Id
$mainMenuFinanceOptionEntity = New-CsAutoAttendantCallableEntity -Identity $mainMenuFinanceOptionTarget -Type SharedVoicemail
$mainMenuFinanceOption2 = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone2 -CallTarget $mainMenuFinanceOptionEntity
$mainMenuDBNoption3 = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone3 -Prompt $dialByNameMenuPrompt
$mainMenuOperator0 = New-CsAutoAttendantMenuOption -Action TransferCallToOperator -DtmfResponse Tone0

$mainMenuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Thank you for calling Adam's. For Technical Support press 1. For Finance press 2. If you know the name of the person you would like to reach, press 3. For our address and email information, press 4. For all other inquiries please press 0 to speak with the operator." 
$mainMenuOptions = New-CsAutoAttendantMenu -Name "Main Menu" -MenuOptions @($mainMenuSupportOption1, $mainMenuFinanceOption2, $mainMenuDBNoption3, $addressPromptOption, $mainMenuOperator0) -Prompts @($mainMenuPrompt)
$mainMenuCallFlow = New-CsAutoAttendantCallFlow -Name "Main Menu Call Flow" -Menu $mainMenuOptions

$afterHoursGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Thank you for calling Adam's. Our offices are now closed. Regular business hours are Monday through Friday from 8:30 am to 5:00 pm eastern time."
$afterHoursMenuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "If you know the name of the person you would like to leave a voicemail for, press 1. For our address and email information press 4."

$afterHoursMenu = New-CsAutoAttendantMenu -Name "After Hours Menu" -MenuOptions @($mainMenuSupportOption1, $mainMenuFinanceOption2, $dialByNameOption, $addressPromptOption) -Prompts @($afterHoursMenuPrompt)
$afterHoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours Call Flow" -Greetings @($afterHoursGreetingPrompt) -Menu $afterHoursMenu
$afterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $afterHoursSchedule.Id -CallFlowId $afterHoursCallFlow.Id

# Create Holiday Prompts and Menu Options
$christmasSchedule = Get-CsOnlineSchedule | where {$_.Name -eq 'Christmas'}
$christmasGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Thank you for calling Adam's. Our offices are currently closed for the Christmas holiday. Regular business hours are Monday through Friday from 8:30 am to 5:00 pm eastern time."
$christmasMenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$christmasMenu = New-CsAutoAttendantMenu -Name "Christmas Menu" -MenuOptions @($christmasMenuOption)
$christmasCallFlow = New-CsAutoAttendantCallFlow -Name "Christmas" -Greetings @($christmasGreetingPrompt) -Menu $christmasMenu
$christmasCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $christmasSchedule.Id -CallFlowId $christmasCallFlow.Id

$newyearSchedule = Get-CsOnlineSchedule | where {$_.Name -eq 'New Years Day'}
$newyearGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Thank you for calling Adam's. Our offices are currently closed for the New Year's holiday. Regular business hours are Monday through Friday from 8:30 am to 5:00 pm eastern time."
$newyearMenuOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
$newyearMenu = New-CsAutoAttendantMenu -Name "New Year Menu" -MenuOptions @($newyearMenuOption)
$newyearCallFlow = New-CsAutoAttendantCallFlow -Name "New Year" -Greetings @($newyearGreetingPrompt) -Menu $newyearMenu
$newyearCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $newyearSchedule.Id -CallFlowId $newyearCallFlow.Id

$mainAutoAttendant = New-CsAutoAttendant -Name "Adam's Main" -DefaultCallFlow $mainMenuCallFlow -CallFlows @($afterHoursCallFlow, $christmasCallFlow, $newyearCallFlow) -CallHandlingAssociations @($afterHoursCallHandlingAssociation, $christmasCallHandlingAssociation, $newyearCallHandlingAssociation ) -LanguageId "en-US" -TimeZoneId "Eastern Standard Time" -Operator $mainEntity 
$mainAutoAttendantID = (Get-CsAutoAttendant -Identity $mainAutoAttendant.id)
New-CsOnlineApplicationInstanceAssociation -Identities @($mainInstanceID) -ConfigurationID $mainAutoAttendantID.Id -ConfigurationType AutoAttendant 

New-CsOnlineApplicationInstance -UserPrincipalName adam_Support_RA@adam.io -DisplayName "Adam's Support RA" -ApplicationID "0000000-0000-0000-0000-000000000" 
Set-MsolUser -UserPrincipalName "adam_Support_RA@adam.io" -UsageLocation CA 
Set-MsolUserLicense -UserPrincipalName "adam_Support_RA@adam.io" -AddLicenses "reseller-account:PHONESYSTEM_VIRTUALUSER" 
$supportInstanceID = (Get-CsOnlineUser -Identity "adam_Support_RA@adam.io").Identity 
$supportEntity = New-CsAutoAttendantCallableEntity -Identity $supportInstanceID -Type ApplicationEndpoint 
$supportMenuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt "Thank you for calling the Adam's technical support line. To speak with a support specialist please press 1. If you know the name of the person you would like to reach, press 2." 
$supportMenuCQ = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Tone1 -CallTarget $supportCQ
$supportMenuDBN = New-CsAutoAttendantMenuOption -Action Announcement -DtmfResponse Tone2 -Prompt $dialByNameMenuPrompt -Action TransferCallToTarget $dialByNameAutoAttendant
$supportMenuOptions = New-CsAutoAttendantMenu -Name "Main Menu" -MenuOptions @($supportMenuCQ, $supportMenuDBN) -Prompts @($supportMenuPrompt)
$supportMenuCallFlow = New-CsAutoAttendantCallFlow -Name "Main Menu Call Flow" -Greetings $supportMenuPrompt -Menu $supportMenuOptions
$supportAutoAttendant = New-CsAutoAttendant -Name "Adam's Support" -DefaultCallFlow $supportMenuCallFlow -LanguageId "en-US" -TimeZoneId "Eastern Standard Time" -Operator $supportEntity 
$supportAutoAttendantID = (Get-CsAutoAttendant -Identity $supportAutoAttendant.id) 
New-CsOnlineApplicationInstanceAssociation -Identities @($supportInstanceID) -ConfigurationID $supportAutoAttendantID.Id -ConfigurationType AutoAttendant

New-CsOnlineApplicationInstance -UserPrincipalName Support_CQ_RA@adam.io -DisplayName "Support Call Queue RA" -ApplicationID "0000000-0000-0000-0000-000000000"
Set-MsolUser -UserPrincipalName "Support_CQ_RA@adam.io" -UsageLocation CA
Set-MsolUserLicense -UserPrincipalName "Support_CQ_RA@adam.io" -AddLicenses "reseller-account:PHONESYSTEM_VIRTUALUSER"
New-CsCallQueue -Name "Support Call Queue" -AgentAlertTime 20 -AllowOptOut $true -DistributionLists $supportGroupID -UseDefaultMusicOnHold $true -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Disconnect -TimeoutThreshold 2700 -RoutingMethod LongestIdle -ConferenceMode $true -LanguageID "en-US"
$supportCQ = (Get-CsOnlineUser -Identity "Support_CQ_RA@adam.io").Identity
$supportCallQueueID = (Get-CsCallQueue -NameFilter "Support_CQ_RA@adam.io").Identity
New-CsOnlineApplicationInstanceAssociation -Identities @($supportCQ) -ConfigurationID $supportCallQueueID -ConfigurationType CallQueue

New-CsOnlineApplicationInstance -UserPrincipalName Finance_CQ_RA@adam.io -DisplayName "Finance Call Queue RA" -ApplicationID "0000000-0000-0000-0000-000000000"
Set-MsolUser -UserPrincipalName "Finance_CQ_RA@adam.io" -UsageLocation CA
Set-MsolUserLicense -UserPrincipalName "Finance_CQ_RA@adam.io" -AddLicenses "reseller-account:PHONESYSTEM_VIRTUALUSER"
New-CsCallQueue -Name "Finance Call Queue" -AgentAlertTime 20 -AllowOptOut $true -DistributionLists $financeGroupID -UseDefaultMusicOnHold $true -OverflowAction DisconnectWithBusy -OverflowThreshold 200 -TimeoutAction Disconnect -TimeoutThreshold 2700 -RoutingMethod LongestIdle -ConferenceMode $true -LanguageID "en-US"
$financeCQ = (Get-CsOnlineUser -Identity "Finance_CQ_RA@adam.io").Identity
$financeCallQueueID = (Get-CsCallQueue -NameFilter "Finance").Identity
New-CsOnlineApplicationInstanceAssociation -Identities @($financeCQ) -ConfigurationID $financeCallQueueID -ConfigurationType CallQueue
