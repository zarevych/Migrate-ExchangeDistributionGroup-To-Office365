#Migrate-ExchangeDistributionGroup-To-Office365
#Migrating-ExchangeDistributionGroup-To-Office365
#Get-DistributionGroup-Info
#Get-ExchangeGroup-Info

<#
.SYNOPSIS
 Get Exchange Distribution Group Info.
 
  
.DESCRIPTION
 Get Exchange Distribution Group Info.
 Prepare xml-file with group info.
 
 Run on Exchange Management Shell or 

 
 Usage:
    .\Get-ExchangeGroup-Info.ps1 -DistributionGroup "myMailGroup"
    .\Get-ExchangeGroup-Info.ps1 -DistributionGroup "myMailGroup*" -AcceptedDomain "company.com, mail.company.ua"
    .\Get-ExchangeGroup-Info.ps1 -AcceptedDomain "company.com, mail.company.ua"

 In Section Initialisations you may set default value:


.PARAMETER DistributionGroup

    Example: -DistributionGroup "MyDistributionGroup"
    Example: -DistributionGroup "Office*", "Department*"
    Example: -DistributionGroup "MyDistributionGroup", "Office*"


.PARAMETER AcceptedDomain
    Optional parameter
    Check accepted domain in distribution group
    By default uses Exchange default domain

    Example: -AcceptedDomain "company.com"
    Example: -AcceptedDomain "company.com, mail.company.ua"


.PARAMETER TargetOU

    Example: -TargetOU "OU=MigratedDG365,DC=company,DC=com"


.PARAMETER Confirm
    To confirm disable DistributionGroup from Exchange

    Example: -Confirm:$true


.PARAMETER XMLFile
    Set xml-file
    
    Example: -XMLFile "C:\Migrate-DistributionGroup\DistributionGroups-Info.xml"


.PARAMETER LogFile
    Optional parameter
    Set to log-file path

    Example: -LogFile "C:\Migrate-DistributionGroup\Log.txt"


.EXAMPLE
   .\Get-ExchangeGroup-Info.ps1 

   Description
   -----------
   Run for 


.EXAMPLE
   .\Get-ExchangeGroup-Info.ps1 -Confirm $false -XMLFile "All-DistributionGroups.xml"

   Description
   -----------
   Run for
   


.NOTES
   File Name  : Get-ExchangeGroup-Info.ps1
   Ver.       : 1.2002
   Written by : Andriy Zarevych

   Find me on :
   * My Blog  :	https://angry-admin.blogspot.com/
   * LinkedIn :	https://linkedin.com/in/zarevych/
   * Github   :	https://github.com/zarevych

   Change Log:
   V1.1908    : Initial version
   V1.2002    : Add progress bar

#>


#---------------------------------------------------------[Initialisations]--------------------------------------------------------


[CmdletBinding()]

param(
    
    # DistributionGroup
    [string[]]$DistributionGroup,

    # AcceptedDomain
    [string]$AcceptedDomain,

    # TargetOU
    [string]$TargetOU,

    # Set the XML file Name
    [Parameter(Mandatory=$false, HelpMessage="Enter XML file name, example: .\DistributionGroups-Info.xml", ValueFromPipelineByPropertyName=$true)]    
    [string]$XMLFile = ".\DistributionGroups-Info.xml",
    
    # Set the Log file Name
    [Parameter(Mandatory=$false, HelpMessage="Enter log file name, example: .\Log.txt", ValueFromPipelineByPropertyName=$true)]    
    [string]$LogFile = ".\Log_$(Get-Date -f 'yyyy-MM-dd').txt",
 
    # Confirm
    [bool]$Confirm = $false
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#[System.Collections.ArrayList]$DistributionGroups=@()

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

$host.ui.rawui.windowtitle="PowerShell | Get-ExchangeGroup-Info.ps1 | $([char]0x00A9) Andriy Zarevych"

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

Add-Content $LogFile "- Start -"
Add-Content $LogFile "$ScriptName"
Add-Content $LogFile $(Get-Date)


# Import Modules
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;
Import-Module ActiveDirectory

$MsgText = "Report file: " + $XMLFile
Write-Log $LogFile $MsgText

# Get AcceptedDomain
$SMTPDomains = $AcceptedDomain
if ($SMTPDomains) {
    $SMTPDomains = $SMTPDomains -replace (' ')
    $SMTPDomains = $SMTPDomains -split ","
}
else {
    $SMTPDomains = Get-AcceptedDomain
    foreach ($SMTPDomain in $SMTPDomains) {
        if ($SMTPDomain.default) {
            $SMTPDomains = $SMTPDomain.DomainName.Address         
        }
    }
}
$MsgText = $SMTPDomains
Write-Log $LogFile "Accepted Domain:"
Write-Log $LogFile $SMTPDomains

# Get Exchange Distribution or Security Group
if ($DistributionGroup){
    $DistributionGroup = $DistributionGroup -split ","
    $DistributionGroup
    if ($DistributionGroup.Count -eq 1) {
        $DistribGroup = [system.String]::Join(" ", $DistributionGroup)
        $DistributionGroups = Get-DistributionGroup $DistribGroup -ErrorAction SilentlyContinue
    }
    else{
        $DistributionGroups = $DistributionGroup | Get-DistributionGroup -ErrorAction SilentlyContinue
    }
}
else {
    $DistributionGroups = Get-DistributionGroup -ResultSize unlimited | sort Name
}

if ($DistributionGroups -eq $null){
    $MsgText = "Object ""$DistributionGroup"" couldn't be found on Exchange Distribution or Security Group."
    Write-Log $LogFile $MsgText "Red"
    Write-Log $LogFile ""
    break
}
$DistributionGroups
#$DistributionGroups.Count
Write-Host


# Create The Document
$XmlWriter = New-Object System.XMl.XmlTextWriter($XMLFile, $Null)

# Set The Formatting
$xmlWriter.Formatting = "Indented"
$xmlWriter.Indentation = "4"

# Write the XML Decleration
$xmlWriter.WriteStartDocument()

# Write Root Element
$xmlWriter.WriteStartElement("Groups")

$GroupCount = 1

foreach ($DistribGroup in $DistributionGroups){
    $GroupDisplayName=$DistribGroup.DisplayName
    
    #Show Progress Info
    $MsgText = $GroupDisplayName + " " + $GroupCount + "/" + $DistributionGroups.Count
    Write-Progress -Activity $MsgText -Status "Progress:" -PercentComplete ($GroupCount/$DistributionGroups.Count*100)

    $MsgText = "Group: "+ $DistribGroup
    Write-Log $LogFile $MsgText

    #Group Properties
    $xmlWriter.WriteStartElement('Group', $GroupCount)
    $xmlWriter.WriteElementString("DisplayName", $DistribGroup.DisplayName)
    
    $xmlWriter.WriteElementString("Name", $DistribGroup.Name)
    $xmlWriter.WriteElementString("SamAccountName", $DistribGroup.SamAccountName)
    
    $xmlWriter.WriteElementString("GroupType", $DistribGroup.RecipientType)
    Write-Log $LogFile $DistribGroup.RecipientType

    # --- Email Address ---
    $xmlWriter.WriteElementString("EmailAddressPolicyEnabled",$DistribGroup.EmailAddressPolicyEnabled)

    # Get Alias
    $xmlWriter.WriteElementString("Alias",$DistribGroup.Alias)

    # Get PrimarySmtpAddress
    $xmlWriter.WriteElementString("PrimarySmtpAddress",$DistribGroup.PrimarySmtpAddress)
    Write-Log $LogFile $DistribGroup.PrimarySmtpAddress

    # Get SMTP Addresses
    $SmtpAddresses=$DistribGroup.EmailAddresses | select ProxyAddressString
    foreach ($SmtpAddress in $SmtpAddresses){
        if ($SmtpAddress.ProxyAddressString.Substring(0, [Math]::Min(($SmtpAddress.ProxyAddressString.Length - 0), 5)) -eq "smtp:"){
            $SmtpAddress = $SmtpAddress.ProxyAddressString
            if ($DistribGroup.PrimarySmtpAddress -eq $SmtpAddress.TrimStart("SMTP:")){
                Continue
            }
            $SmtpAddress = $SmtpAddress.TrimStart("smtp:")
            foreach ($SMTPDomain in $SMTPDomains){
                if ($SmtpAddress.Substring($SmtpAddress.Length - $SMTPDomain.Length) -eq $SMTPDomain){
                    $xmlWriter.WriteElementString("SmtpAddress",$SmtpAddress)
                }
            }
        }
    }

    #--- RequireSenderAuthentication ---
    $xmlWriter.WriteElementString("RequireSenderAuthenticationEnabled", $DistribGroup.RequireSenderAuthenticationEnabled)
    $MsgText = "RequireSenderAuthentication: " + $DistribGroup.RequireSenderAuthenticationEnabled
    Write-Log $LogFile $MsgText

    #--- Advanced --- HiddenFromAddressListsEnabled
    $xmlWriter.WriteElementString("HiddenFromAddressListsEnabled",$DistribGroup.HiddenFromAddressListsEnabled)
    $MsgText = "HiddenFromAddressLists: " + $DistribGroup.HiddenFromAddressListsEnabled
    Write-Log $LogFile $MsgText
 
    #--- Advanced --- ReportTo
    $xmlWriter.WriteElementString("ReportToManager", $DistribGroup.ReportToManagerEnabled)
    $xmlWriter.WriteElementString("ReportToOriginator", $DistribGroup.ReportToOriginatorEnabled)

    #--- Group Membership Approval --- Join
    $xmlWriter.WriteElementString("MemberJoinRestriction",$DistribGroup.MemberJoinRestriction)

    #--- Group Membership Approval --- Leave
    $xmlWriter.WriteElementString("MemberDepartRestriction",$DistribGroup.MemberDepartRestriction)

    #--- Mail Flow ---
    #--- Mail Flow --- AcceptMessagesOnlyFrom
    Write-Host "AcceptMessagesFrom:"
    # AcceptMessagesFrom User
    foreach ($AcceptMessagesFrom in $DistribGroup.AcceptMessagesOnlyFrom){
        $User = Get-ADUser -Filte {Name -eq $AcceptMessagesFrom.Name} | Select-Object SamAccountName, UserPrincipalName
        $xmlWriter.WriteElementString("AcceptMessagesFrom", $User.UserPrincipalName)
        $User.UserPrincipalName
    }
    # AcceptMessagesFrom DL
    foreach ($AcceptMessagesFrom in $DistribGroup.AcceptMessagesOnlyFromDLMembers){
        $Group = Get-DistributionGroup $AcceptMessagesFrom -ErrorAction SilentlyContinue
        $xmlWriter.WriteElementString("AcceptMessagesFrom", $Group.PrimarySmtpAddress.Address)
        $Group.PrimarySmtpAddress.Address
    }
  
    # --- Mail Flow --- RejectMessagesFromSendersOrMembers
    Write-Host "RejectMessagesFrom:"
    $DistribGroup.RejectMessagesFrom #| select Name
    foreach ($RejectMessagesFrom in $DistribGroup.RejectMessagesFrom){
        $User = Get-ADUser -Filte {Name -eq $RejectMessagesFrom.Name} | Select-Object SamAccountName, UserPrincipalName
        $xmlWriter.WriteElementString("RejectMessagesFrom", $User.UserPrincipalName)
    }
    # Write-Host "AcceptMessagesFrom DL:"
    foreach ($RejectMessagesFrom in $DistribGroup.RejectMessagesFromDLMembers){
        $Group = Get-DistributionGroup $RejectMessagesFrom -ErrorAction SilentlyContinue
        $xmlWriter.WriteElementString("RejectMessagesFrom", $Group.PrimarySmtpAddress.Address)
    }


    Write-Host "SendAs:"
    $Users = Get-ADPermission -identity $DistribGroup.sAMAccountName | where {($_.ExtendedRights -like "*Send-As*")}

    #$Users
    foreach ($User in $Users){
        try {
            $user = Get-ADUser -Identity $user.User.SecurityIdentifier.Value -ErrorAction SilentlyContinue
        }
        catch {
            continue
        }
        $xmlWriter.WriteElementString("SendAs", $user.UserPrincipalName)
        $user.UserPrincipalName
    }

    Write-Host "SendOnBehalf:"
    # --- Mail Flow --- GrantSendOnBehalfTo
    $users = $DistribGroup.GrantSendOnBehalfTo
    foreach ($user in $users){
        $user = get-aduser $user.DistinguishedName
        $xmlWriter.WriteElementString("SendOnBehalf", $user.UserPrincipalName)
        $user.UserPrincipalName
    }

    #--- Group Information --- Group Owners
    foreach ($GroupManagedBy in $DistribGroup.ManagedBy) {
        $MsgText = "ManagedBy: " + $GroupManagedBy.Name
        Write-Log $LogFile $MsgText
        $GroupManager = Get-ADUser -Filte {Name -eq $GroupManagedBy.Name} | Select-Object SamAccountName, UserPrincipalName
        $GroupManagerUPN = $GroupManager.UserPrincipalName
        $xmlWriter.WriteElementString("ManagedBy",$GroupManagerUPN)
    }

    # --- Group Members ---
    $Members = Get-DistributionGroupMember $DistribGroup -ResultSize unlimited | sort RecipientType -ErrorAction Continue
    Foreach ($Member in $Members){
        
        # RecipientType: MailUser or UserMailbox
        if (($Member.RecipientType -eq "MailUser") -or ($Member.RecipientType -eq "UserMailbox")){
            $MemberName = $Member.SamAccountName
            $MName = Get-ADUser -Filte {SamAccountName -eq $MemberName} | Select-Object SamAccountName, UserPrincipalName
            $GroupMemberUPN = $MName.UserPrincipalName
            $MsgText = "Member - " + $Member.RecipientType + " - " + $GroupMemberUPN
            Write-Log $LogFile $MsgText
            $xmlWriter.WriteElementString("GroupMemberUPN",$GroupMemberUPN)
        }

        # RecipientType: MailUniversalDistributionGroup/"MailUniversalSecurityGroup"
        if (($Member.RecipientType -eq "MailUniversalDistributionGroup") -or ($Member.RecipientType -eq "MailUniversalSecurityGroup")) {         
            $xmlWriter.WriteElementString("GroupMemberDL",$Member.PrimarySmtpAddress)
            if ($Member.RecipientType -eq "MailUniversalDistributionGroup"){
                $MsgText = "Member - MailUniversalDistributionGroup: " + $Member.PrimarySmtpAddress
                Write-Log $LogFile $MsgText "Yellow"
                #Write-Host "Member - MailUniversalDistributionGroup:" $Member.Name $Member.PrimarySmtpAddress -ForegroundColor Yellow
            }
            if ($Member.RecipientType -eq "MailUniversalSecurityGroup"){
                $MsgText = "Member - MailUniversalSecurityGroup: " + $Member.PrimarySmtpAddress
                Write-Log $LogFile $MsgText "Yellow"
                #Write-Host "Member - MailUniversalSecurityGroup:" $Member.Name $Member.PrimarySmtpAddress -ForegroundColor Yellow
            }
        }    

        # RecipientType: MailContact
        if ($Member.RecipientType -eq "MailContact"){
            Write-Host "Member - MailContact:" $Member.PrimarySmtpAddress -ForegroundColor Yellow
            $xmlWriter.WriteElementString("GroupMemberMailContact",$Member.PrimarySmtpAddress)
        }
    }


    # Get Group CustomAttribute
    $xmlWriter.WriteElementString("CustomAttribute1",$DistribGroup.CustomAttribute1)
    $xmlWriter.WriteElementString("CustomAttribute2",$DistribGroup.CustomAttribute2)
    $xmlWriter.WriteElementString("CustomAttribute3",$DistribGroup.CustomAttribute3)
    $xmlWriter.WriteElementString("CustomAttribute4",$DistribGroup.CustomAttribute4)
    $xmlWriter.WriteElementString("CustomAttribute5",$DistribGroup.CustomAttribute5)

    #
    $xmlWriter.WriteEndElement() # <-- Closing Group


    #
    # Disable-DistributionGroup and move to $TargetOU
    #
    if ($Confirm){
        $MsgText = "Disable distribution group " + $DistribGroup
        Write-Log $LogFile $MsgText
        Disable-DistributionGroup -Identity $DistribGroup -Confirm:$false
        
        # Move DistributionGroup to $TargetOU
        if ($TargetOU){
            if([adsi]::Exists("LDAP://$TargetOU")) {
                $MsgText = "Move AD group " + $DistribGroup + " to " + $TargetOU
                Write-Log $LogFile $MsgText
                Get-ADGroup $DistribGroup.Name | Move-ADObject -identity {$_.objectguid} -TargetPath $TargetOU
           }
           else{
                $MsgText = "Warning. Object not found: " + $TargetOU
                Write-Log $LogFile $MsgText "Yellow"
           }
        }
    }

    Write-Host
    $GroupCount = $GroupCount + 1

    sleep 1
}

# Write the Document

# Write Close Tag for Root Element
$xmlWriter.WriteEndElement() # <-- Closing RootElement

# End the XML Document
$xmlWriter.WriteEndDocument() | Out-Null

# Finish The Document
$xmlWriter.Finalize | Out-Null
$xmlWriter.Flush | Out-Null
$xmlWriter.Close()

#
Add-Content $LogFile "- End -"
Add-Content $LogFile ""
#[System.IO.File]::AppendAllText($LogFile,"- End -")
#[System.IO.File]::AppendAllText($LogFile,"")
