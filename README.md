
.SYNOPSIS
    Parses active mailboxes and searches for hidden rules.

.DESCRIPTION
    Sign into Exchange server, parse mailboxes, export to CSV in C:\Windows\Temp\MailboxRulesCount.csv

.PARAMETER Name
    Will Prompt for for 1,2,0

.EXAMPLE
    PS> .\emailRuleChecker.ps1

.EXAMPLE
    PS> 25 of 337
    PS> Press 1 to search enabled hidden rule mailboxes, Press 0 to exit.

.LINK
