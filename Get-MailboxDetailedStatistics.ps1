<##############################################################################
# Bit.Expert - SCRIPT - POWERSHELL
# NAME: EXO_SendStatistics.ps1
# 
# AUTHOR:  Christopher McLendon
# DATE:  May 8, 2018
# EMAIL: christopher@themclendons.com
# 
# COMMENT:  This script will connect to your Exhange Online tenant and,
# given a date range, calculate the total number of e-mails along with payload
# size for each Exchange user.  It uses the Message ID for each message to
# avoid counting duplicate messages where a user may have sent the same 
# message to multiple users (e.g. multiple users on the to: line)
#
# VERSION HISTORY
# 1.0 2018.05.08 Initial Version
# 1.1 2018.05.09 Refactoring and logic to recycle remote session
# 1.2 2018.05.10 Comments, fixes and additional counters
##############################################################################>

<#
.SYNOPSIS
Create a remoting session for Office 365
.DESCRIPTION
Checks for an existing remoting session to Office 365 and and Exchange Online.
If it finds an open session it will recycle that session, otherwise it will call 
Get-PSSession to establish a new session.
#>
function Connect-ExchangeOnline
{
    $remoteSession = $null

    Write-Host "Connecting session..." -ForegroundColor Green

    foreach($session in Get-PSSession) {
        if ($session.ComputerName -eq "outlook.office365.com" -and $session.ConfigurationName -eq "Microsoft.Exchange" -and $session.State -eq "Opened" ) {
            #re-use the remote session that was previously opened
            $remoteSession = $session
            Write-Host "Re-connecting remote session to Office 365:" -ForegroundColor Green
            Write-Host $remoteSession -ForegroundColor Yellow
            break
        }
    }

    if ($remoteSession -eq $null) {
        #no remote session to recycle -- need to create a new session
        Write-Host "Enter administrative credentials to connect to Office 365:" -ForegroundColor Red
        $remoteSession =  New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
        
        #import the remote session
        Import-PSSession $remoteSession
    }

    Write-Host "...Session connected." -ForegroundColor Green
}

<#
.SYNOPSIS
Get the list of Exchange Recipients from Office 365 
#>
function Get-ExchangeRecipients
{
    BEGIN {
        $recipientdetails = @{}

        <#
            Get a list of accounts from Exchange Online - we only filter against the Primary SMTP Address
        #>

        Write-Host "Getting recipient addresses from Exchange Online..." -ForegroundColor Green -NoNewline
        $recipients = Get-Recipient -ResultSize Unlimited | Select-Object PrimarySMTPAddress

        $I = 0

        <#
            Iterate through each recipient address and build a new collection using the primary SMTP
            address as the hash table key; we will store a custom object on each address for the purposes
            of calculating send/receive statistics
        #>
        foreach($recipient in $recipients) {
            <#
                This could be a long running activity depending on the number of recipients in the organization
                so we wil show a progress bar
            #>

            $I++
            Write-Progress -Activity "Initializing recipients" -Status "Progress:" -PercentComplete ($I/$recipients.count*100)

            <# 
                Initialize a new custom property object and initialize a bunch of counters we will later use when we
                parse through all of our messages; this object will be stored on our MailRecipients hash table using
                the primary SMTP address as the hash table key for each instance
            #>

            $recipientProperties = [PSCustomObject]@{
                MailAddress = $recipient.PrimarySMTPAddress.ToLower()
                Sent = 0
                SentSize = 0
                SentInternal = 0
                SentInternalSize = 0
                SentExternal = 0
                SentExternalSize = 0
                Received = 0
                ReceivedSize = 0
                ReceivedInternal = 0
                ReceivedInternalSize = 0
                ReceivedExternal = 0
                ReceivedExternalSize = 0
                SentUnique = 0
                SentUniqueSize = 0
                SentUniqueInternal = 0
                SentUniqueInternalSize = 0
                SentUniqueExternal = 0
                SentUniqueExternalSize = 0
            }
            
            <#
                Assign the property object to a mail recipient
            #>

            $recipientdetails[$recipientProperties.MailAddress] = $recipientProperties
        }
    }

    END {
        Write-Host "$($recipients.count) recipients loaded." -ForegroundColor Gray
        Write-Output $recipientdetails
    }
}

<#
.SYNOPSIS
Get's the Exchange Online Message Trace using a start and end date
.PARAMETER startdate
Start date 
.PARAMETER enddate
End date
#>
function Get-SmtpLogFile([string] $startdate, [string] $enddate) {
    <#
        This could be a long running activity depending on the number of messages being parsed
        so we wil show a progress bar
    #>
    
    Write-Host "Getting messages from Exchange Online between $($startdate) and $($enddate)..." -ForegroundColor Green -NoNewline
    $messagetrace = Get-MessageTrace -StartDate $startdate -EndDate $enddate
    Write-Host "...Done." -ForegroundColor White
    
    Write-Host "$($messagetrace.count) total messages loaded..." -ForegroundColor Gray -NoNewline

    #Send messagetrace to standard output
    Write-Output $messagetrace
}

<#
.SYNOPSIS
Calculate send and receive statistics for each of the Exchange recipient accounts
.PARAMETER messagetrace
Message log to parse
.PARAMETER stdin
Recipient table with initialized statistics from Get-ExchangeRecipients
#>
function Initialize-RecipientStatistics
{
    [CmdletBinding()]
	param(
        [object[]] $messagetrace,
		[Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True)]
        $recipients
    )

    BEGIN {
        $uniquemessages = @{}
    }
    PROCESS {

    }
    END {
        $I = 0

        Write-Host "Processing statistics for $($recipients.Count) accounts..." -ForegroundColor Green -NoNewline

        foreach ($message in $messagetrace) {
            $I++
            Write-Progress -Activity "Parsing through messages" -Status "Progress:" -PercentComplete ($I/$messagetrace.count*100)

            $address = $null
            $internalSender = $recipients.ContainsKey($message.SenderAddress)
            $internalReceiver = $recipients.ContainsKey($message.RecipientAddress)
            
            if ($uniquemessages.Contains($message.MessageId)) {
                $firstMessageInstance = $false;
                $uniquemessages[$message.MessageId]++
            }
            else {
                $firstMessageInstance = $true;
                $uniquemessages.Add($message.MessageId, 1)
            }
            

            if ($internalSender -eq $true) {
                <# 
                    Update sender statistics 
                #>

                $address = $recipients[$message.SenderAddress]
                
                <#
                    Sent and SentSize will be compounded -- if the user sends an e-mail to two people
                    at the same time that are part of the same organization it will be counted twice;
                    this counter includes the aggregate of both internal and external e-mail
                #>
                $address.Sent++
                $address.SentSize += $message.Size

                if ($firstMessageInstance -eq $true) {
                    $address.SentUnique++
                    $address.SentUniqueSize += $message.Size
                }

                if ($internalReceiver -eq $true) {
                    <# 
                        SentInternal and SentInternalSize will be compounded -- if the user sends an e-mail to two people
                        at the same time that are part of the same organization it will be counted twice
                    #>
                    $address.SentInternal++
                    $address.SentInternalSize += $message.Size

                    if ($firstMessageInstance -eq $true) {
                        $address.SentUniqueInternal++
                        $address.SentUniqueInternalSize += $message.Size
                    }
                }
                else {                
                    <# 
                        SentExternal and SentExternalSize will be compounded -- if the user sends an e-mail to two people
                        at the same time that outside the organization it will be counted twice
                    #>
                    $address.SentExternal++
                    $address.SentExternalSize += $message.Size

                    if ($firstMessageInstance -eq $true) {
                        $address.SentUniqueExternal++
                        $address.SentUniqueExternalSize += $message.Size
                    }
                }
            }
            
            if ($internalReceiver -eq $true) {
                <# 
                    Update internal recipient statistics for inbound e-mails
                #>

                $address = $recipients[$message.RecipientAddress]

                $address.Received++
                $address.ReceivedSize += $message.Size

                if ($internalSender -eq $true) {
                    $address.ReceivedInternal++
                    $address.ReceivedInternalSize += $node.Size
                }
                else {
                    $address.ReceivedExternal++
                    $address.ReceivedExternalSize += $node.Size
                }
            }
        }

        Write-Host "...Done." -ForegroundColor White
        Write-Output $recipients.Values
    }
}

<#
    SCRIPT START
#>

Write-Host "You entered for -Hello " $startdate
Connect-ExchangeOnline
$messagelog = Get-SmtpLogFile -StartDate "05/07/2018 00:00:00 AM" -EndDate "05/07/2018 11:59:59 PM"
Get-ExchangeRecipients | Initialize-RecipientStatistics -MessageTrace $messagelog
