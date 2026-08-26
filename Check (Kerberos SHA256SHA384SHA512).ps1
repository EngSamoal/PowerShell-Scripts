Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4768} -MaxEvents 50 |
    ForEach-Object {
        $xml = [xml]$_.ToXml()
        $etype = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TicketEncryptionType' }).'#text'
        [PSCustomObject]@{
            Time    = $_.TimeCreated
            Account = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
            EncType = $etype
        }
    } | Sort-Object Time -Descending | Format-Table -AutoSize