$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
$report = @()
$i = 0
ForEach ($mailbox in $mailboxes) {
    $i++
    $LastProcessed = $Null
    Write-Progress -Activity "Scanning Mailbox $($mailbox.DisplayName)" -Status "Scanned: $i of $($mailboxes.Count)"
    $Log = Export-MailboxDiagnosticLogs -Identity $mailbox.UserPrincipalName -ExtendedProperties
    $xml = [xml]($Log.MailboxLog)  
    $LastProcessed = ($xml.Properties.MailboxTable.Property | ? {$_.Name -like "*ELCLastSuccessTimestamp*"}).Value   
    $ItemsDeleted  = $xml.Properties.MailboxTable.Property | ? {$_.Name -like "*ElcLastRunDeletedFromRootItemCount*"}
    If ($LastProcessed -eq $Null) {
        $LastProcessed = "Not processed"}

    $reportLine = [PSCustomObject]@{
            User          = $mailbox.DisplayName
            LastProcessed = $LastProcessed
            ItemsDeleted  = $ItemsDeleted.Value}      
        $report += $reportLine
    }
#$report | Select User, LastProcessed, ItemsDeleted
#$report | Export-CSV C:\Scripts\Inboxsizes.csv
$report | Out-GridView
