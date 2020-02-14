$InactiveDays = New-TimeSpan -days 56
$LastLogonTimeMark = (get-date) – $InactiveDays

Get-ADComputer -Filter * -Properties lastLogontimeStamp,whenCreated |`
?{[DateTime]::FromFileTime($_.lastLogonTimestamp) -lt $LastLogonTimeMark}|`
select Name, @{Name=”LastLogon”;Expression={[datetime]::FromFileTime($_.lastLogontimeStamp)}}, whenCreated|`
sort LastLogon| ft Name,LastLogon, whenCreated –AutoSize