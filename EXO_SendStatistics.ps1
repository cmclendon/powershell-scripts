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
# 1.0 2018.05.09 Added additional recipient detail and some refactoring
##############################################################################>

function getExchangeRecipients ([hashtable] $MailRecipients)
{
    #Get a list of accounts from Exchange Online - we only filter against the Primary SMTP Address
    $Recipients = Get-Recipient -ResultSize Unlimited | Select-Object PrimarySMTPAddress

    foreach($Recipient in $Recipients)
    {
        $props = [PSCustomObject]@{
            DataSize = 0
            MailCount = 0
        }
        
        $MailRecipients[$Recipient.PrimarySMTPAddress.ToLower()] = $props
    }
}

function parseExchangeSMTPLogFile ([Hashtable] $MailRecipients, [Hashtable] $MessageList, [string] $StartDate, [string] $EndDate)
{
    $MessageTrace = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate

    foreach ($node in $MessageTrace)
    {
        #verify uniqueness of message (have we already captured it?)
        if (!$MessageList.Contains($node.MessageId))
        {
            #update sender mail stats
            if ($SmtpAddress.ContainsKey($node.SenderAddress))
            {
                #Initialize message object
                $message = [PSCustomObject]@{
                    MessageSize = $node.Size
                    SenderAddress = $node.SenderAddress
                }

                #Add to our message list and key off the Message Id
                $MessageList.Add($node.MessageId, $message)

                #Update MailRecipient mail send statistics
                $SenderObject = $MailRecipients[$node.SenderAddress]

                $SenderObject.DataSize = $SenderObject.DataSize + $node.Size
                $SenderObject.MailCount = $SenderObject.MailCount + 1
            }
        }
    }
}

#Remove existing PowerShell sessions to Office 365 and create a new one
Remove-PSSession *

#Create a new O365 session and request credential input
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber

$MailRecipients = @{}
$MessageList = @{}

Write-Output "Building recipient collection..."
getExchangeRecipients($MailRecipients)

#MAKE SURE you update your start and end dates here -- IMPORTANT to use the format in the example
Write-Output "Parsing Exchange SMTP log file..."
parseExchangeSMTPLogFile -MailRecipients $MailRecipients -MessageList $MessageList -StartDate "05/07/2018 00:00:00 AM" -EndDate "05/07/2018 11:59:59 PM"

#write column header and output statistics
Write-Output "Sender|Count|Size"
foreach($Recipient in $MailRecipients.Keys)
{
    Write-Output "$($Recipient)|$($MailRecipients[$Recipient].DataSize)|$($MailRecipients[$Recipient].MailCount)"
}
