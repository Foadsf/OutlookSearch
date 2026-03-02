# Test individual DASL filter strings against Outlook
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$items = $inbox.Items

$filters = @(
    @{ Name = "Subject DASL ci_phrasematch"; Filter = "@SQL=""urn:schemas:httpmail:subject"" ci_phrasematch 'test'" }
    @{ Name = "Subject DASL ci_startswith"; Filter = "@SQL=""urn:schemas:httpmail:subject"" ci_startswith 'test'" }
    @{ Name = "Subject DASL LIKE"; Filter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%test%'" }
    @{ Name = "Body DASL ci_phrasematch"; Filter = "@SQL=""urn:schemas:httpmail:textdescription"" ci_phrasematch 'test'" }
    @{ Name = "Body DASL LIKE"; Filter = "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%test%'" }
    @{ Name = "Date DASL >="; Filter = "@SQL=""urn:schemas:httpmail:datereceived"" >= '2025-01-01T00:00:00Z'" }
    @{ Name = "Date DASL >= local fmt"; Filter = "@SQL=""urn:schemas:httpmail:datereceived"" >= '01/01/2025 00:00:00'" }
    @{ Name = "From DASL ci_phrasematch"; Filter = "@SQL=""urn:schemas:httpmail:fromname"" ci_phrasematch 'foad'" }
    @{ Name = "From DASL LIKE"; Filter = "@SQL=""urn:schemas:httpmail:fromname"" LIKE '%foad%'" }
    @{ Name = "Subject+Date DASL AND"; Filter = "@SQL=""urn:schemas:httpmail:subject"" ci_phrasematch 'test' AND ""urn:schemas:httpmail:datereceived"" >= '2025-01-01T00:00:00Z'" }
    @{ Name = "Subject+Body DASL AND"; Filter = "@SQL=""urn:schemas:httpmail:subject"" ci_phrasematch 'test' AND ""urn:schemas:httpmail:textdescription"" ci_phrasematch 'test'" }
    @{ Name = "Subject OR Body DASL"; Filter = "@SQL=(""urn:schemas:httpmail:subject"" ci_phrasematch 'test' OR ""urn:schemas:httpmail:textdescription"" ci_phrasematch 'test')" }
    @{ Name = "JET ReceivedTime"; Filter = "[ReceivedTime] >= '1/1/2025 12:00 AM'" }
    @{ Name = "JET SenderName"; Filter = "[SenderName] = 'foad'" }
    @{ Name = "JET Subject"; Filter = "[Subject] = 'test'" }
    @{ Name = "HasAttachment JET"; Filter = "[HasAttachment] = True" }
)

foreach ($f in $filters) {
    Write-Host -NoNewline ("$($f.Name): ".PadRight(40))
    try {
        $result = $items.Restrict($f.Filter)
        Write-Host "OK ($($result.Count) results)" -ForegroundColor Green
    }
    catch {
        Write-Host "FAIL: $($_.Exception.Message)" -ForegroundColor Red
    }
}
