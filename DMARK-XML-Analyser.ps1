Write-Host "##################################################################" -ForegroundColor Cyan
Start-Sleep -Milliseconds 500
Write-Host "#     Hallo $env:USERNAME - Virtueller DaUfooo wird geladen!     #" -ForegroundColor Yellow
Start-Sleep -Milliseconds 500
Write-Host "##################################################################" -ForegroundColor Cyan
Start-Sleep -Milliseconds 2000
Write-Host "############################################################" -ForegroundColor Red
Start-Sleep -Milliseconds 500
Write-Host "#            !Virtueller DaUfooo wurde geladen!               #" -ForegroundColor Green
Start-Sleep -Milliseconds 500
Write-Host "#          !Super Fast Reading wird aktiviert!             #" -ForegroundColor Yellow
Start-Sleep -Milliseconds 500
Write-Host "# Ahrg,Brtzl,Brtzl,quhf,hmm, rumsprfnns,hmmmm.... fje ARG! #" -ForegroundColor Cyan
Start-Sleep -Milliseconds 500
Write-Host "#############################################################" -ForegroundColor Red
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
    Write-Host "----------------------------------------------------------------------" -ForegroundColor Red
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

        Write-Host "Date Range: $beginFormatted to $endFormatted" -ForegroundColor Yellow
        Write-Host "Organisation: $orgName" -ForegroundColor Green
        Write-Host "Email: $email" -ForegroundColor Green
        Write-Host "Percentage: $pct" -ForegroundColor Green

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
    } 
    else {
        Write-Host "Fehler: XML konnte nicht geladen werden." -ForegroundColor Red
    }
}

Write-Host "#############################################################" -ForegroundColor Yellow
Start-Sleep -Milliseconds 500
Write-Host "#  !Virtueller DaUfooo hat sich alle XML Files durchgelesen!   #" -ForegroundColor Cyan
Start-Sleep -Milliseconds 500
Write-Host "#      Nun male ich Dir ein paar Pixel aufn Screen          #" -ForegroundColor Red
Start-Sleep -Milliseconds 500
Write-Host "# Wisch, Wasch, Schwupp da Wupp DKIM und SPF zeig Dich nun! #" -ForegroundColor Green    
Start-Sleep -Milliseconds 500
Write-Host "#############################################################" -ForegroundColor Yellow
Start-Sleep -Milliseconds 500
Write-Host "`nErgebnisse der Verarbeitung aller Dateien:" -ForegroundColor Cyan
Start-Sleep -Milliseconds 500
$allOutput | Format-Table -Property ReportTime, Organisation, SourceIP, SPF, SPFResult, SPFDomain, DKIM, DKIMResult, DKIMDomain, Count -AutoSize
Start-Sleep -Milliseconds 500

$csvPath = ".\Ergebniss-Auswertung.csv"
$allOutput | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Start-Sleep -Milliseconds 500
Write-Host "#############################################################" -ForegroundColor Yellow
Start-Sleep -Milliseconds 500
Write-Host "#   Virtueller DaUfooo ist nun kaputt und legt sich schlafen   #" -ForegroundColor Cyan
Start-Sleep -Milliseconds 500
Write-Host "#     Aiiiiiijjjjaaaaaaa!! Der Process wurde gekillt!       #" -ForegroundColor Red
Start-Sleep -Milliseconds 500
Write-Host "#   Virtueller DaUfooo ist nun eingeschlafen - Bitte Flüstern  #" -ForegroundColor Green
Start-Sleep -Milliseconds 500
Write-Host "#############################################################" -ForegroundColor Yellow
Start-Sleep -Milliseconds 500
Write-Host "Es wurden insgesamt $processedFilesCount XML-Dateien verarbeitet." -ForegroundColor Cyan
Write-Host "Die Ergebnisse wurden erfolgreich in '$csvPath' exportiert." -ForegroundColor Green
Start-Sleep -Milliseconds 500
Read-Host "Drücke eine beliebige Taste zum Beenden"
