<#  
.NOTES
===========================================================================
Created on:     3/6/2024
Created by:     Michael Lubert
Updated on:     3/14/2024
Version:        1.0.1
Version History:        
        1.0.1 - 2024-03-14: Clarified debug variable names
                            Added debug variable descriptions
                            Fixed sorting and categorizing
        1.0.0 - 2024-03-06: Initial Version
===========================================================================
.DESCRIPTION
This script will send an email report of the size of all active user mailboxes and mailbox archvies.
#>

##################################################################
##################################################################
########## BEGIN VARIABLES. CONFIGURATION OPTIONS BELOW ##########
##################################################################
##################################################################

#################################
### BEGIN DEBUG MODE SETTINGS ###
#################################

$DEBUGNOSEND=$False             # Prevent Sending Emails, set to $False for production
$DEBUGMSG=$False                # Display Debug details
$DEBUGDUMP=$False               # Spit out HTML file version of Report
$DEBUGDUMPFILE="test.html"      # HTML File Name
$DEBUGLIMIT=$False;             # Don't process every user.
$DEBUGCOUNT=10;                 # Number of users to process

###############################
### END DEBUG MODE SETTINGS ###
###############################

######################################
### BEGIN SERVER SPECIFIC SETTINGS ###
######################################

$DIRPATH = "C:\Scripts\M365UserMailboxReport\"

######################################
### BEGIN SERVER SPECIFIC SETTINGS ###
######################################

############################
### BEGIN EMAIL SETTINGS ###
############################

$COMPANY="COMPANY"

$REPORTTIME = Get-Date -format "yyyy-MM-dd"
$REPORTEMAILSUBJECT="Daily $COMPANY M365 User Report - $REPORTTIME"
$FROMEMAIL = "administrator@company.com"
$PRIORITY="Normal"
$TORECIPIENTS = @(
    @{EmailAddress = @{Address = "user@company.com"}}
    @{EmailAddress = @{Address = "administrator@company.com"}}
)
$TO=@()
$TO+=@{EmailAddress = @{Address = "user@company.com"}}
$TO+=@{EmailAddress = @{Address = "administrator@company.com"}}

############################
### END EMAIL SETTINGS ###
############################

###################################
### BEGIN ENTRA APP ID SETTINGS ###
###################################

$CERTTHUMB = "0000000000000000000000000000000000000000"                         # Thumbprint for Self-Signed Cert used for Entra Authentication
$EARAPPID = "00000000-0000-0000-0000-000000000000"                              # Entra App Registration Application ID
$EARTENANTID = "00000000-0000-0000-0000-000000000000"                           # Entra Tenant ID
$EARORG = "company.onmicrosoft.com"                                           # Entra Organization Name

###################################
### END ENTRA APP ID SETTINGS ###
###################################

####################################
### BEGIN EMAIL MESSAGE TEMPLATE ###
####################################

$HTMLHEAD="<html><style>
    BODY{font-family: Arial; font-size: 8pt;}
    H1{font-size: 16px;}
    H2{font-size: 14px;}
    H3{font-size: 12px;}
    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
    TD{border: 1px solid black; padding: 5px; }
    td.pass{background: #C1E1C1;}
    td.warn{background: #FAA0A0;}
    td.fail{background: #FAA0A0; color: #000000;}
    td.info{background: #85D4FF;
    </style><body>
    <h3>"+$COMPANY+" M365 User Report - "+$REPORTTIME+"</h3>"

$MBXTABLE="<table><tr><th>User</th><th>Mailbox Size</th><th>Number of Items</th><th>Archive Size</th><th>Number of Archived Items</th></tr>"

##################################
### END EMAIL MESSAGE TEMPLATE ###
##################################

#################################################################################################
#################################################################################################
#################################################################################################
########## END VARIABLES. THERE IS NOTHING MORE FOR YOU TO CONFIGURE BEYOND THIS POINT ##########
#################################################################################################
#################################################################################################
#################################################################################################

###############################
########## FUNCTIONS ##########
###############################

function SendMessage
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)]
        [string]$From,
        [Parameter(Mandatory, Position=1)]
        [object]$To,
        [Parameter(Mandatory, Position=2)]
        [string]$Subject,
        [Parameter(Mandatory, Position=3)]
        [string]$Priority,
        [Parameter(Mandatory, Position=4)]
        [string]$Message
    )

    if($DebugMode){
        Write-Host "FROM: $From" -ForegroundColor Green
        Write-Host "TO: $From" -ForegroundColor Green
        Write-Host "SUBJECT: $Subject" -ForegroundColor Green
        Write-Host "PRIORITY: $Priority" -ForegroundColor Green
        Write-Host "MESSAGE: $Message" -ForegroundColor Green
    }
    $Email = @{
        Message = @{
            Subject = $Subject
            ToRecipients = $To # Array of @{EmailAddress = @{Address = $ToAddress}}
            Body = @{
                contentType = "HTML";
                content = $Message
            }
            Importance = $Priority
        }
        #SaveToSentItems = "true"
    }
    Send-MGUserMail -UserId $From -BodyParameter $Email
}

###################################
########## END FUNCTIONS ##########
###################################

if(test-path -Path $DIRPATH){}
else{$DIRPATH="./"}

#######################
### CONNECT TO M365 ###
#######################

Try
{
    Connect-ExchangeOnline -AppId $EARAPPID -CertificateThumbprint $CERTTHUMB -Organization $EARORG -ShowBanner:$false
    Connect-MgGraph -ClientId $EARAPPID -TenantId $EARTENANTID -CertificateThumbprint $CERTTHUMB -nowelcome
    #echo "A"
}
Catch
{
   $_ | Out-File ($DIRPATH + "\" + "Log.txt") -Append
   exit
}

#################################
### GET ACTIVE USER MAILBOXES ###
#################################

$ULIST=@()
$USERS=get-mailbox -resultsize unlimited | Where-Object {$_.IsShared -eq $False -And $_.IsResource -eq $False}
foreach ($USER in $USERS)
{
    if($DEBUGMSG){$USER}
	$USERZ=(get-mailboxstatistics -Identity $USER); 
	$USERAZ=(Get-MailboxStatistics -Identity $user -Archive)
	$USERTIS=$USERZ.TotalItemSize.Value -replace '.*\(| bytes\).*|,' | % {'{0:N2}' -f ($_ / 1gb)}
    $USERTISBYTE=[int]$USERZ.TotalItemSize.Value -replace '.*\(| bytes\).*|,' | % {'{0:N2}' -f ($_)}
	$USERATIS=$USERAZ.TotalItemSize.Value -replace '.*\(| bytes\).*|,' | % {'{0:N2}' -f ($_ / 1gb)}
    $USERATISBYTE=[int]$USERAZ.TotalItemSize.Value -replace '.*\(| bytes\).*|,' | % {'{0:N2}' -f ($_)}
    if($DEBUGMSG){
        $USERZ.TotalItemSize.Value
        $USERTIS;
        $USERTISBYTE;
        $USERAZ.TotalItemSize.Value
        $USERATIS;
        $USERATISBYTE;

    }
    if($USERTISBYTE -gt 35000000000)
    {
        if($DEBUGMSG){echo "Should fail"}
        $class="fail"
    }
    elseif($USERTISBYTE -lt 35)
    {
        if($DEBUGMSG){echo "Should pass";}
        $class="pass"
    }
    else
    {
        if($DEBUGMSG){echo "Should warn"}
        $class="warn"
    }
    $TABLEROW="<tr><td class=$class>"+$USER.DisplayName+"</td><td class=$class>"+$USERTIS+" GB</td><td class=$class>"+$USERZ.ItemCount+"</td><td class=$class>"+$USERATIS+" GB</td><td class=$class>"+$USERAZ.ItemCount+"</td></tr>"
    $ULIST+=$USERTIS.PadLeft(10,"0")+","+$TABLEROW
    if($DEBUGLIMIT)
    {
        $DEBUGCOUNT--;
        if($DEBUGCOUNT -eq 0){break}
    }
}

#########################################
### SORT MAILBOXES BY DESCENDING SIZE ###
#########################################

$SORTULIST=($ULIST | sort-object -descending)
foreach ($ENTRY in $SORTULIST)
{
    $ENTRY.split(",")[1]
    $MBXTABLE+=$ENTRY.split(",")[1]
}
$MBXTABLE+="</table>"

##################
### CLOSE HTML ###
##################

$HTMLTAIL = "</body></html>"
$HTMLREPORT = $HTMLHEAD+$MBXTABLE+$HTMLTAIL

##############################
### CREATE DEBUG HTML FILE ###
##############################

if($DEBUGDUMP)
{
    $HTMLREPORT > $DEBUGDUMPFILE
}

####################
### SEND MESSAGE ###
####################

if($DEBUGNOSEND -eq $False){SendMessage -From $FROMEMAIL -To $TO -Subject $REPORTEMAILSUBJECT -Priority $PRIORITY -Message $HTMLREPORT}

Disconnect-ExchangeOnline -Confirm:$false
