#Migrate-ExchangeDistributionGroup-To-Office365
#Create-Office365MailGroup

<#
.SYNOPSIS
 
 Run script with Administrator privileges.
 
.DESCRIPTION
 Run script with Office365/Microsoft administrator privileges.

 Usage:
    .\Create-Office365MailGroup.ps1 -User "admin@company.onmicrosoft.com" -Password "AdminPassword"
    .\Create-Office365MailGroup.ps1 -User "admin@company.com" -Password "AdminPassword" -OtherGroupOwner "user1@company.com" -XMLFile "myDistributionGroup.xml"


.PARAMETER User
    Office 365 admin user name

    Example: -User "admin@domain.com"


.PARAMETER Password
    Office 365 admin user password

    Example: -Password "AdminPassword"


.PARAMETER OtherGroupOwner

    Example: -OtherGroupOwner "user@domain.com"


.PARAMETER XMLFile
    Optional parameter
    Example: -XMLFile


.PARAMETER LogFile
    Optional parameter
    Set to log-file path

    Example: -LogFile "C:\Migrate-DistributionGroup\Log.txt"


.PARAMETER GroupNotes
    Optional parameter
    Set exchange group description

    Example: -GroupNotes "Mail distribution group has been migrated to Office 365 using the script`r`nhttps://github.com/zarevych/Migrate-ExchangeDistributionGroup-To-Office365"


.EXAMPLE
   .\.ps1 -ComputerName

   Description
   -----------
   Run for 


.EXAMPLE
   .\.ps1 -User "admin@domain.com"

   Description
   -----------
   Run for
   Enter user password
   

.NOTES
   File Name  : Config-Servers-DSC.ps1
   Ver.       : 1.2002
   Written by : Andriy Zarevych

   Find me on :
   * My Blog  :	https://angry-admin.blogspot.com/
   * LinkedIn :	https://linkedin.com/in/zarevych/
   * Github   :	https://github.com/zarevych

   Change Log:
   V1.1908    : Initial version
   V1.2002    : Add progress bar
              : Minor fixes

#>


[CmdletBinding()]

param(
    # User cred
    [string]$User,
    [string]$Password,

    #
    [string]$OtherGroupOwner,

    # Set the XML file Name
    [Parameter(Mandatory=$false, HelpMessage="Enter XML file name, example: .\DistributionGroups-Info.xml", ValueFromPipelineByPropertyName=$true)]    
    [string]$XMLFile = ".\DistributionGroups-Info.xml",

    # Set the Log file Name
    [Parameter(Mandatory=$false, HelpMessage="Enter log file Name, example: .\Log.txt", ValueFromPipelineByPropertyName=$true)]    
    [string]$LogFile = ".\Log_$(Get-Date -f 'yyyy-MM-dd').txt",

    #
    [string]$GroupNotes = "Mail distribution group has been migrated to Office 365 using the script`r`nhttps://github.com/zarevych/Migrate-ExchangeDistributionGroup-To-Office365"

)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

[System.Collections.ArrayList]$GroupOwners=@()
#[System.Collections.ArrayList]$GroupMembers=@()

#----------------------------------------------------------[Functions]-------------------------------------------------------------

function Write-Log ([String]$LogFile, [String]$Message, [String]$ForegroundColor) {
    Add-Content $LogFile $Message -ErrorAction SilentlyContinue
    if ($ForegroundColor){
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    else {
        Write-Host $Message
    }
}

#----------------------------------------------------------[Execution]-------------------------------------------------------------

$host.ui.rawui.windowtitle="PowerShell | Create-Office365MailGroup.ps1 | $([char]0x00A9) Andriy Zarevych"

# Get Script Path. Change default values to script directory path. Needed for remote execution.
$MyPath = (Get-Variable MyInvocation).Value
$MyDirectoryPath = Split-Path $MyPath.MyCommand.Path


# Set the Log file Name
if ($LogFile.StartsWith(".\Log_")){
    $LogFile = $MyDirectoryPath + "\Log_$(Get-Date -f 'yyyy-MM-dd').txt"
}

# Set the XML file Name
if ($XMLFile.StartsWith(".\DistributionGroups-Info.xml")){
    $XMLFile = $MyDirectoryPath + "\DistributionGroups-Info.xml"
}

$ScriptName = $MyPath.MyCommand.Name
$ScriptPath = $MyPath.InvocationName

Write-host
Write-host "Run: $ScriptName"
Write-host "LogFile Name: $LogFile"
Write-Host

Add-Content $LogFile "- Start -"
Add-Content $LogFile "$ScriptName"
Add-Content $LogFile $(Get-Date)

# Check xml-file
if (Test-Path "$XMLFile")
{
    $MsgText = "XML File: "+ $XMLFile
    Write-Log $LogFile $MsgText
    [xml]$XMLGroupInfo = Get-Content ".\DistributionGroups-Info.xml"
}
else
{
    $MsgText = "Cannot find XML file: "+ $XMLFile
    Write-Log $LogFile $MsgText "Red"
    Add-Content $LogFile "- End -"
    Add-Content $LogFile ""
    Break
}

Write-Host
Write-Host "Notes:"
Write-Host $GroupNotes
Write-Host

# Get and check user cred
if ($User -eq "" -And $Password -eq ""){
    try {
        $Cred = Get-Credential
    }
    catch {
        Write-host "Missing user credential parameters"
        Break
    }
}
else{
    if ($User -And $Password){
        $NewPassword=$Password
        $NewPassword = ConvertTo-SecureString $NewPassword -AsPlainText -Force
        $Cred = New-Object System.Management.Automation.PSCredential($User,$NewPassword)
    }
    if ($User -And $Password -eq ""){
        $Cred = Get-Credential -UserName $User -Message "Enter $User password please"
        if (!($Cred)) {
            Write-Host "Parameter -Password empty"
            Write-Host
            break
            }
        }
}
if ($User -eq "" -And $Password){
    Write-Host "Don't use parameter -Password with out -User"
    break
}



# Connecting to Office 365 - Exchange Online
Write-Log $LogFile "Connecting to Exchange online..." "Green"
try {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection -Name "Exchange Online" -ErrorAction Stop
}
catch
{
    Write-Log $LogFile "Connecting to remote server outlook.office365.com failed." "Red"
    Add-Content $LogFile "- End -"
    Add-Content $LogFile ""
    Write-Host
    Break
}
try {
    Import-PSSession $Session -DisableNameChecking
}
catch {
    $MsgText = $_.Exception
    #$MsgText = $_.Exception.Message
    Write-Log $LogFile $MsgText "Red"
}

$domain = Get-AcceptedDomain | Where-Object Default -EQ 'True'
Write-Host
Write-Host "*** Welcome to Exchange Online for the domain $domain ***"
Write-Host

foreach ($Group in $XMLGroupInfo.Groups.Group){

    #Show Progress Info
    $MsgText = $Group.Name + " " + $Group.xmlns + "/" + $XMLGroupInfo.Groups.Group.Count
    try {
        Write-Progress -Activity $MsgText -Status "Progress:" -PercentComplete ($Group.xmlns/$XMLGroupInfo.Groups.Group.Count*100) -ErrorAction SilentlyContinue
    }
    catch {  
    }

    $GroupOwners=@()

    Write-Log $LogFile ""
    $MsgText = "Group: "+ $Group.Name
    Write-Log $LogFile $MsgText
        
    if ($Group.GroupType -eq "MailUniversalDistributionGroup"){
        $MsgText = "Creating MailUniversalDistributionGroup - DisplayName:" + $Group.DisplayName + " Alias:" + $Group.Alias
        Write-Log $LogFile $MsgText
        New-DistributionGroup -Name $Group.Name -DisplayName $Group.DisplayName -Alias $Group.Alias -type Distribution -IgnoreNamingPolicy -ErrorAction SilentlyContinue -Notes $GroupNotes
    }
    if ($Group.GroupType -eq "MailUniversalSecurityGroup"){
        $MsgText = "Creating MailUniversalSecurityGroup - DisplayName:" + $Group.DisplayName + " Alias:" + $Group.Alias
        Write-Log $LogFile $MsgText
        New-DistributionGroup -Name $Group.Name -DisplayName $Group.DisplayName -Alias $Group.Alias -type Security -IgnoreNamingPolicy -ErrorAction SilentlyContinue -Notes $GroupNotes
    }

    #PrimarySmtpAddress
    $MsgText = "PrimarySmtpAddress: "+ $Group.PrimarySmtpAddress
    Write-Log $LogFile $MsgText
    Set-DistributionGroup -Identity $Group.Alias -EmailAddresses @{Add=$Group.PrimarySmtpAddress}
    Set-DistributionGroup -Identity $Group.Alias -PrimarySmtpAddress $Group.PrimarySmtpAddress

    #HiddenFromAddressLists
            if ($Group.HiddenFromAddressListsEnabled.ToLower() -eq "false"){
                Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -HiddenFromAddressListsEnabled $false -ErrorAction Continue
            }
            if ($Group.HiddenFromAddressListsEnabled.ToLower() -eq "true"){
                Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -HiddenFromAddressListsEnabled $true -ErrorAction Continue
            }

    # Add SMTP Addresses
    foreach ($SmtpAddress in $Group.SmtpAddress){
        Set-DistributionGroup -Identity $Group.Alias -EmailAddresses @{Add=$SmtpAddress}
    }
    
    #
    #
    # AcceptMessagesOnlyFromSendersOrMembers
    foreach ($AcceptMessagesFrom in $Group.AcceptMessagesFrom){
        Set-DistributionGroup $Group.Alias -BypassSecurityGroupManagerCheck -AcceptMessagesOnlyFromSendersOrMembers @{Add=$AcceptMessagesFrom}
    }
    
    # RejectMessagesFromSendersOrMembers
    foreach ($RejectMessagesFrom in $Group.RejectMessagesFrom){
        Set-DistributionGroup $Group.Alias -BypassSecurityGroupManagerCheck -RejectMessagesFromSendersOrMembers @{Add=$RejectMessagesFrom}
    }

    #
    # SendAs
    foreach ($SendAs in $Group.SendAs){
        try {
            Add-RecipientPermission $Group.Alias -AccessRights SendAs -Trustee $SendAs -Confirm:$false -ErrorAction SilentlyContinue
        }
        catch {
            Continue
        }
    }
    #
    # GrantSendOnBehalfTo

    foreach ($SendOnBehalf in $Group.GrantSendOnBehalf){
        Set-DistributionGroup -Identity $Group.Alias -GrantSendOnBehalfTo @{Add=$SendOnBehalf}
    }

    #
    #

    #--- Group Information --- Set Group Owners
    foreach ($GroupManager in $group.ManagedBy){
        $Recipient = Get-Recipient $GroupManager -ErrorAction SilentlyContinue
        if (($Recipient.RecipientType -eq "UserMailbox") -or ($Recipient.RecipientType -eq "MailUser")){
            $GroupOwners.Add($GroupManager) | Out-Null
        }
    }
    Write-Log $LogFile "ManagedBy:"
    if ($GroupOwners){
        Set-DistributionGroup -Identity $Group.Alias -ManagedBy $GroupOwners -ErrorAction SilentlyContinue
        Write-Log $LogFile $GroupOwners
    }
    else {
        if ($OtherGroupOwner){
            $Recipient = Get-Recipient $OtherGroupOwner -ErrorAction SilentlyContinue
            if (($Recipient.RecipientType -eq "UserMailbox") -or ($Recipient.RecipientType -eq "MailUser")){
                Set-DistributionGroup -Identity $Group.Alias -ManagedBy $OtherGroupOwner -ErrorAction SilentlyContinue
                Write-Log $LogFile $OtherGroupOwner "Yellow"
            }
        }
    }

    # --- Add Group Members --- UserMailbox/MailUser
    foreach ($Member in $Group.GroupMemberUPN){
        $Recipient = Get-Recipient $Member -ErrorAction SilentlyContinue
        if (($Recipient.RecipientType -eq "UserMailbox") -or ($Recipient.RecipientType -eq "MailUser")){
            Add-DistributionGroupMember -Identity $Group.Alias -Member $Member -ErrorAction SilentlyContinue
        }
        else {
            $MsgText = "Unknown member: " + $Member
            Write-Log $LogFile $MsgText "Yellow"
        }
    }

    # --- Add Group Members --- Distribution/Security Group
    foreach ($Member in $Group.GroupMemberDL){
        $Recipient = Get-Recipient $Member -ErrorAction SilentlyContinue
        if (($Recipient.RecipientType -eq "MailUniversalDistributionGroup") -or ($Recipient.RecipientType -eq "MailUniversalSecurityGroup")){
            Add-DistributionGroupMember -Identity $Group.Alias -Member $Member -ErrorAction SilentlyContinue
        }
        else {
            $MsgText = "Unknown member - Distribution or Security group: " + $Member
            Write-Log $LogFile $MsgText "Yellow"
        }
    }


    # RequireSenderAuthenticationEnabled
    if ($Group.RequireSenderAuthenticationEnabled.ToLower() -eq "false"){
        try {
            Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -RequireSenderAuthenticationEnabled $false -ErrorAction Continue
        }
        catch {
            $MsgText = $_.Exception
            Write-Log $LogFile $MsgText "Yellow"
            #Continue        
        }
    }
    if ($Group.RequireSenderAuthenticationEnabled.ToLower() -eq "true"){
        try {
            Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -RequireSenderAuthenticationEnabled $true -ErrorAction Continue
        }
        catch {
            $MsgText = $_.Exception
            Write-Log $LogFile $MsgText "Yellow"
            #Continue
        }
    }

    # --- Group Membership Approval --- Join
    try {
        Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -MemberJoinRestriction $Group.MemberJoinRestriction -ErrorAction Stop
    }
    catch {
        #$_.Exception
        #$_.Exception.Message
        $MsgText = $_.Exception.Message
        Write-Log $LogFile $MsgText "Yellow"
        $MsgText = "WARNING: Check Directory Sync or on-premises group " + $Group.Alias
        Write-Log $LogFile $MsgText "Red"
        Continue
    }
    # --- Group Membership Approval --- Leave
    Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -MemberDepartRestriction $Group.MemberDepartRestriction -ErrorAction Continue
    
    # --- Group Membership Approval --- Join/Leave
    Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -MemberJoinRestriction $Group.MemberJoinRestriction -MemberDepartRestriction $Group.MemberDepartRestriction -ErrorAction Continue

    # --- Advanced --- ReportTo
    #Set-DistributionGroup -Identity $Group.Alias -BypassSecurityGroupManagerCheck -ReportToManagerEnabled $Group.ReportToManager -ReportToOriginatorEnabled $Group.ReportToOriginator

    sleep 1

}


write-host
Write-Log $LogFile "Disconnecting from Exchange on-line"

if (Get-PSSession -Name "Exchange Online")
	{
		Remove-PSSession -Name "Exchange Online"
		Write-Log $LogFile "Session disconnected"
	}


Add-Content $LogFile "- End -"
Add-Content $LogFile ""

Remove-PSSession $Session
