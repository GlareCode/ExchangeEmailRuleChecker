<#
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

#>

$path = $env:windir

# Check if ran as Administrator----------------------------------------------------------

$gotchya = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")

if ($gotchya -eq $False){
    Write-Host "**************Please reopen PowerShell as Administrator***************"
    Write-Host "**************Please reopen PowerShell as Administrator***************"
    Write-Host "**************Please reopen PowerShell as Administrator***************"
    
    $exit = Read-Host "Press 0 to Exit"

    if ($exit -ne 0){
        stop-process -id $PID
    }Else{
        stop-process -id $PID
    }
}

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine | Out-Null

Import-Module ExchangeOnlineManagement

function Get-HiddenRules {
    $table = New-Object System.Data.Datatable
    [void]$table.Columns.Add("Name")
    [void]$table.Columns.Add("Hidden Rules")

    foreach ($box in $allMailboxes){

        Write-Host $allmailboxes.IndexOf($box) "of" $allmailboxes.count " Press CTRL+C to Quit" -NoNewLine
        $result = Get-InboxRule -Mailbox $box -IncludeHidden | Where-Object {($_.Name -ne "Junk E-Mail Rule") -and ($_.Enabled -ne $False)}

        # Avoid Exchange Rate Limiting
        $rando = Get-Random -Minimum 502 -Maximum 660
        start-sleep -Milliseconds $rando

        $isnull = [string]::IsNullOrEmpty($result)

        if ($isnull -eq $True){
            [void]$table.Rows.Add("$box","$False")
            Clear-Host
        }Elseif ($isnull -eq $False){
            [void]$table.Rows.Add("$box","$True")
            Clear-Host
            continue
        }Else{
            Clear-Host
            continue
        }
    }
    Write-Host "Find the output file in C:\Windows\Temp\MailboxHiddenRuleCheck.csv"
    $table | Export-Csv -LiteralPath $path\Temp\MailboxHiddenRuleCheck.csv
    return Get-Choice
}


function Get-TotalNumber {
    $table = New-Object System.Data.Datatable
    [void]$table.Columns.Add("Name")
    [void]$table.Columns.Add("Total Rules")
    [void]$table.Columns.Add("Total with Hidden Rules")
    [void]$table.Columns.Add("Rule Difference")

    foreach ($box in $allmailboxes){
        Write-Host "Finding Rules: Mailbox" $allmailboxes.IndexOf($box) "of" $allmailboxes.count " Press CTRL+C to Quit" -NoNewLine
        $result = Get-InboxRule -Mailbox $box

        # Avoid Exchange Rate Limiting
        $rando = Get-Random -Minimum 502 -Maximum 660
        start-sleep -Milliseconds $rando

        $result2 = Get-InboxRule -Mailbox $box -IncludeHidden

        $isnull = [string]::IsNullOrEmpty($result)
        $isnull2 = [string]::IsNullOrEmpty($result2)

        if ($isnull -eq $true -and $isnull2 -eq $true){
            [void]$table.Rows.Add("$box","0","0","0")
            Clear-Host
            Continue

        }Elseif ($isnull -eq $true -and $isnull2 -eq $false){
            for ($numb = 0; $numb -lt $result2.count; $numb ++){
                Write-Host "Counting Rules for: " $box " Press CTRL+C to Quit" -NoNewline
            }
            [void]$table.Rows.Add("$box","0","$numb","$numb")
            Clear-Host
            Continue

        }Elseif ($isnull -eq $false -and $isnull2 -eq $false){
            for ($numb1 = 0; $numb1 -lt $result.count; $numb1 ++){
                Write-Host "Counting Rules for: " $box " Press CTRL+C to Quit" -NoNewline
            }
            for ($numb2 = 0; $numb2 -lt $result2.count; $numb2 ++){
                Write-Host "Counting Rules for: " $box " Press CTRL+C to Quit" -NoNewline
            }
            [void]$table.Rows.Add("$box","$numb1","$numb2",$numb1 - $numb2)
            Clear-Host
            Continue
    
        }ElseIf ($isnull -eq $false -and $isnull2 -eq $true){
            for ($numb = 0; $numb -lt $result.count; $numb ++){
                Write-Host "Counting Rules for: " $box " Press CTRL+C to Quit" -NoNewline
            }
            [void]$table.Rows.Add("$box","$numb","0","$numb")
            Clear-Host
            Continue
        }
    }

    Write-Host "Find the output file in C:\Windows\Temp\MailboxRulesCount.csv"
    $table | Export-Csv -LiteralPath $path\Temp\MailboxRulesCount.csv
    return Get-Choice
}

function Get-Choice {
    $choose = Read-Host "`nPress 1 to count enabled hidden rules for each mailboxes,`nPress 2 to count the total number of hidden and non-hidden rules for each mailbox,`nPress 0 to Exit."

    If ($choose -eq "1"){
        Get-HiddenRules
    }Elseif ($choose -eq "2"){
        Get-TotalNumber
    }Elseif ($choose -eq "0"){
        Write-Host "Goodbye"
        return
        Write-Host "*****Type this to list the hidden rules of specific user*****`n`nGet-InboxRule -Mailbox user.mailbox@company.com -IncludeHidden | Where-Object {($_.Name -ne "Junk E-Mail Rule") -and ($_.Enabled -ne $False)}`n"
    }Else {
        Get-Choice
    }
}


function Get-Connection {
    $choose = Read-Host "`nPress 1 to connect to Exchange Online`nPress 0 to Exit`n"

    if ($choose -eq "1"){
    Connect-ExchangeOnline
    $allMailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object -ExpandProperty UserPrincipalName
    Write-Host "`n*****Total Mailboxes*****`n"
    $allmailboxes.count
    Write-Host "`n*****Total Mailboxes*****`n"
    start-sleep -Seconds 2

    }Elseif ($choose -eq "0"){
        Write-Host "Goodbye"
        return
    }Else{
        Get-Connection
    }

    Get-Choice
}


Get-Connection
