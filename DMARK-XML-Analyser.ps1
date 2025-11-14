Write-Host "############################################################" -ForegroundColor Gray
Write-Host "#*                                                        *#" -ForegroundColor Green
Write-Host "#****************   Hallo, $env:USERNAME :)   *******************#" -ForegroundColor Yellow
Write-Host "#*                                                        *#" -ForegroundColor Magenta
Write-Host "#**   DMARK-XML-Analyser Powershell Script von DaUfooo   **#" -ForegroundColor Cyan
Write-Host "#*                                                        *#" -ForegroundColor Green
Write-Host "############################################################" -ForegroundColor Gray
Start-Sleep -Seconds 1
$directoryPath = ".\XML-Reports\"

if (-Not (Test-Path $directoryPath)) {
    Write-Host "Fehler: Das Verzeichnis wurde nicht gefunden." -ForegroundColor Red
    exit
}

$xmlFiles = Get-ChildItem -Path $directoryPath -Filter *.xml

if ($xmlFiles.Count -eq 0) {
    Write-Host "Fehler: Es wurden keine XML-Dateien gefunden." -ForegroundColor Red
    exit
}

$processedFilesCount = 0
$allOutput = @()

foreach ($xmlFile in $xmlFiles) {
    Write-Host "----------------------------------------------------------------------" -ForegroundColor White
    Write-Host "`nVerarbeite Datei: $($xmlFile.Name)" -ForegroundColor Cyan
    $processedFilesCount++

    try {
        [xml]$xml = Get-Content -Path $xmlFile.FullName -ErrorAction Stop
    } catch {
        Write-Host "Fehler: Die Datei konnte nicht geladen werden. $_" -ForegroundColor Red
        continue
    }

    if ($xml -ne $null) {
        $orgName = $xml.feedback.report_metadata.org_name
        $email = $xml.feedback.report_metadata.email
        $reportId = $xml.feedback.report_metadata.report_id
        $domain = $xml.feedback.policy_published.domain
        $pct = $xml.feedback.policy_published.pct
        $beginTimestamp = [int64]$xml.feedback.report_metadata.date_range.begin
        $endTimestamp = [int64]$xml.feedback.report_metadata.date_range.end
        $begin = [System.DateTime]::Parse("1970-01-01 00:00:00").AddSeconds($beginTimestamp).ToLocalTime()
        $end = [System.DateTime]::Parse("1970-01-01 00:00:00").AddSeconds($endTimestamp).ToLocalTime()
        $beginFormatted = $begin.ToString('yyyy-MM-dd HH:mm:ss')
        $endFormatted = $end.ToString('yyyy-MM-dd HH:mm:ss')
                        
        $tempOutput = @()

        foreach ($record in $xml.feedback.record) {
            $currentTime = [System.DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')

            foreach ($authResult in $record.auth_results) {
                foreach ($dkimResult in $authResult.dkim) {
                    $dkimDomain = $dkimResult.domain
                    $dkimResultValue = $dkimResult.result
                    $dkimColor = if ($dkimResultValue -eq 'pass') { 'Green' } elseif ($dkimResultValue -eq 'fail') { 'Red' } else { 'Yellow' }
                }

                foreach ($spfResult in $authResult.spf) {
                    $spfDomain = $spfResult.domain
                    $spfResultValue = $spfResult.result
                    $spfColor = if ($spfResultValue -eq 'pass') { 'Green' } elseif ($spfResultValue -eq 'fail') { 'Red' } else { 'Yellow' }
                }
            }

            $outputObj = New-Object PSObject -property @{
                ReportTime     = $currentTime
                Organisation   = $orgName
                SourceIP       = $record.row.source_ip
                SPF            = $record.row.policy_evaluated.spf
                SPFResult      = ($record.auth_results.spf | ForEach-Object { $_.result }) -join ", "
                SPFDomain      = ($record.auth_results.spf | ForEach-Object { $_.domain }) -join ", "
                DKIM           = $record.row.policy_evaluated.dkim
                DKIMResult     = ($record.auth_results.dkim | ForEach-Object { $_.result }) -join ", "
                DKIMDomain     = ($record.auth_results.dkim | ForEach-Object { $_.domain }) -join ", "
                Count          = $record.row.count
            }

            $tempOutput += $outputObj
        }

        $allOutput += $tempOutput

Write-Host "Date Range: $beginFormatted to $endFormatted" -ForegroundColor Yellow
Write-Host "Organisation: $orgName" -ForegroundColor Magenta
Write-Host "Email: $email" -ForegroundColor Gray
Write-Host "Source IP: $($record.row.source_ip)" -ForegroundColor Red
Write-Host "DKIM: $($record.row.policy_evaluated.dkim)" -ForegroundColor $dkimColor
Write-Host "SPF: $($record.row.policy_evaluated.spf)" -ForegroundColor $spfColor
Write-Host "DKIM Result: $($record.auth_results.dkim | ForEach-Object { $_.result })" -ForegroundColor $dkimColor
Write-Host "SPF Result: $($record.auth_results.spf | ForEach-Object { $_.result })" -ForegroundColor $spfColor
Write-Host "Percentage: $pct" -ForegroundColor Gray

    } 
    else {
        Write-Host "Fehler: XML konnte nicht geladen werden." -ForegroundColor Red
    }
}

Write-Host "`nErgebnisse der Verarbeitung aller Dateien:" -ForegroundColor Cyan
Write-Host "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------" -ForegroundColor White
$allOutput | Format-Table -Property ReportTime, Organisation, SourceIP, SPF, SPFResult, SPFDomain, DKIM, DKIMResult, DKIMDomain, Count -AutoSize
Write-Host "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------" -ForegroundColor White

$csvPath = ".\Ergebniss-Auswertung.csv"
$allOutput | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8


Write-Host "##################################################################" -ForegroundColor Gray
Write-Host "# *                                                             *#" -ForegroundColor Magenta
Write-Host "# *  DMARK-XML-Analyser hat $processedFilesCount XML-Files verabeitet             *#" -ForegroundColor Cyan
Write-Host "# *  Die Ergebnisse wurden exportiert.                          *#" -ForegroundColor Yellow
Write-Host "# *  Ordner des Exports: '$csvPath'           *#" -ForegroundColor Red
Write-Host "# *                                                             *#" -ForegroundColor Green
Write-Host "##################################################################" -ForegroundColor Gray

Write-Host "Drücke eine beliebige Taste zum" -ForegroundColor Yellow
Read-Host "Beenden"
