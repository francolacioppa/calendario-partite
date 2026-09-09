# ==============================================================================
# Script per aggiornare il calendario partite (Serie A, UCL, UEL) con emittenti TV
# Genera il file calendario_partite.csv formattato per Excel (UTF-8 con BOM, delimitatore ;)
# ==============================================================================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$wc = New-Object System.Net.WebClient
$wc.Encoding = [System.Text.Encoding]::UTF8

$competitions = @(
    @{ OrigName = "Serie A"; ShortName = "Serie A"; Url = "https://www.calciointv.com/indexfi.php?comp=Serie%20A" },
    @{ OrigName = "Champions League"; ShortName = "UCL"; Url = "https://www.calciointv.com/indexfi.php?comp=Champions%20League" },
    @{ OrigName = "Europa League"; ShortName = "UEL"; Url = "https://www.calciointv.com/indexfi.php?comp=Europa%20League" }
)

$allMatches = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenKeys = [System.Collections.Generic.HashSet[string]]::new()

foreach ($comp in $competitions) {
    Write-Host "Scaricamento $($comp.ShortName)..." -ForegroundColor Cyan
    try {
        $bytes = $wc.DownloadData($comp.Url)
        $html = [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    catch {
        Write-Warning "Errore nello scaricamento di $($comp.ShortName): $_"
        continue
    }
    
    $blocks = [regex]::Matches($html, '(?s)<div class="div_partido.*?</div>\s*<div class="c"></div>')
    
    foreach ($b in $blocks) {
        $val = $b.Value
        $mDate = [regex]::Match($val, '&dates=(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})')
        $mText = [regex]::Match($val, '&text=([^\r\n&]+)')
        $mTv = [regex]::Match($val, 'TV:\s*([^\r\n&]+)')
        
        if ($mDate.Success -and $mText.Success) {
            $year = [int]$mDate.Groups[1].Value
            $month = [int]$mDate.Groups[2].Value
            $day = [int]$mDate.Groups[3].Value
            $hour = [int]$mDate.Groups[4].Value
            $min = [int]$mDate.Groups[5].Value
            
            $dtStart = [datetime]::new($year, $month, $day, $hour, $min, 0)
            $dtEnd = $dtStart.AddHours(2)
            
            $rawMatch = [System.Net.WebUtility]::HtmlDecode($mText.Groups[1].Value.Trim())
            $rawTv = if ($mTv.Success) { [System.Net.WebUtility]::HtmlDecode($mTv.Groups[1].Value.Trim()) } else { "Da definire" }
            
            # Normalizzazione nomi squadre per la Serie A (e coppe)
            $matchClean = $rawMatch -replace '\bInter Milan\b', 'Inter' -replace '\bAC Milan\b', 'Milan'
            
            # Chiave univoca per evitare duplicati desktop/mobile
            $key = "$year$month$day-$hour$min-$matchClean"
            if ($seenKeys.Contains($key)) { continue }
            $seenKeys.Add($key) | Out-Null
            
            # Logica assegnazione emittente
            $emittente = "Da definire"
            if ($comp.ShortName -eq "Serie A") {
                if ($rawTv -match "DAZN" -and ($rawTv -match "Sky" -or $rawTv -match "NOW")) {
                    $emittente = "DAZN / SKY"
                } elseif ($rawTv -match "DAZN") {
                    $emittente = "DAZN"
                } elseif ($rawTv -match "Sky" -or $rawTv -match "NOW") {
                    $emittente = "Sky"
                } else {
                    $emittente = "DAZN"
                }
            } else {
                # UCL e UEL
                if ($rawTv -match "Prime Video") {
                    $emittente = "Amazon Prime Video"
                } elseif ($rawTv -match "Sky" -or $rawTv -match "NOW") {
                    if ($rawTv -match "TV8") {
                        $emittente = "Sky / TV8"
                    } else {
                        $emittente = "Sky"
                    }
                } elseif ($rawTv -match "TV8") {
                    $emittente = "TV8"
                } else {
                    $emittente = $rawTv
                }
            }
            
            $allMatches.Add([PSCustomObject]@{
                PlannedStartDate = $dtStart.ToString("dd/MM/yyyy HH:mm")
                PlannedEndDate = $dtEnd.ToString("dd/MM/yyyy HH:mm")
                Competizione = $comp.ShortName
                Partita = $matchClean
                Emittente = $emittente
                Stato_Programmazione = if ($hour -eq 1 -or $hour -eq 0) { "Orario da definire" } else { "Confermato" }
                SortDate = $dtStart
            })
        }
    }
}

# Ordinamento cronologico
$sorted = $allMatches | Sort-Object SortDate

# Costruzione righe CSV
$csvLines = [System.Collections.Generic.List[string]]::new()
$csvLines.Add("Planned Start Date;Planned End Date;Competizione;Partita;Emittente;Stato_Programmazione")

foreach ($m in $sorted) {
    $partita = '"' + $m.Partita.Replace('"', '""') + '"'
    $emittente = '"' + $m.Emittente.Replace('"', '""') + '"'
    $line = "$($m.PlannedStartDate);$($m.PlannedEndDate);$($m.Competizione);$partita;$emittente;$($m.Stato_Programmazione)"
    $csvLines.Add($line)
}

# Salvataggio file con codifica UTF-8 BOM per compatibilità Excel
$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$csvPath = Join-Path $baseDir "calendario_partite.csv"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($csvPath, $csvLines, $utf8Bom)

Write-Host "Aggiornamento completato con successo!" -ForegroundColor Green
Write-Host "File salvato in: $csvPath"
Write-Host "Totale partite esportate: $($sorted.Count)"
