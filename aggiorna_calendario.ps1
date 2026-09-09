# ==============================================================================
# Script per aggiornare il calendario partite (Serie A, UCL, UEL) con emittenti TV
# Mantiene lo storico delle partite passate e unisce i nuovi aggiornamenti
# Output CSV: Delimitatore ",", Competizione come prima colonna, UTF-8 BOM
# ==============================================================================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$wc = New-Object System.Net.WebClient
$wc.Encoding = [System.Text.Encoding]::UTF8

$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$csvPath = Join-Path $baseDir "calendario_partite.csv"

# 1. Carica lo storico esistente (se presente)
$masterMatches = [System.Collections.Generic.Dictionary[string, PSCustomObject]]::new()

if (Test-Path $csvPath) {
    try {
        $firstLine = (Get-Content $csvPath -First 1)
        $del = if ($firstLine -match ";") { ";" } else { "," }
        $existing = Import-Csv -Path $csvPath -Delimiter $del -Encoding UTF8
        
        foreach ($row in $existing) {
            $comp = $row.Competizione
            $match = $row.Partita
            $start = $row.'Planned Start Date'
            $end = $row.'Planned End Date'
            $emittente = $row.Emittente
            $stato = $row.Stato_Programmazione
            
            if ($comp -and $match) {
                $key = "$comp|$match"
                
                # Parse datetime per ordinamento
                $dtSort = [datetime]::MinValue
                if ($start -and [datetime]::TryParseExact($start, "dd/MM/yyyy HH:mm", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dtSort)) {
                    # Parsed successfully
                }
                
                $masterMatches[$key] = [PSCustomObject]@{
                    Competizione = $comp
                    PlannedStartDate = $start
                    PlannedEndDate = $end
                    Partita = $match
                    Emittente = $emittente
                    Stato_Programmazione = $stato
                    SortDate = $dtSort
                }
            }
        }
        Write-Host "Caricate $($masterMatches.Count) partite dallo storico esistente." -ForegroundColor Cyan
    } catch {
        Write-Warning "Avviso nella lettura dello storico esistente: $_"
    }
}

# 2. Scarica i dati aggiornati dalle competizioni
$competitions = @(
    @{ OrigName = "Serie A"; ShortName = "Serie A"; Url = "https://www.calciointv.com/indexfi.php?comp=Serie%20A" },
    @{ OrigName = "Champions League"; ShortName = "UCL"; Url = "https://www.calciointv.com/indexfi.php?comp=Champions%20League" },
    @{ OrigName = "Europa League"; ShortName = "UEL"; Url = "https://www.calciointv.com/indexfi.php?comp=Europa%20League" }
)

$scrapedCount = 0

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
            
            # Normalizzazione nomi squadre
            $matchClean = $rawMatch -replace '\bInter Milan\b', 'Inter' -replace '\bAC Milan\b', 'Milan'
            
            # Logica emittente
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
                if ($rawTv -match "Prime Video") {
                    $emittente = "Amazon Prime Video"
                } elseif ($rawTv -match "Sky" -or $rawTv -match "NOW") {
                    if ($rawTv -match "TV8") { $emittente = "Sky / TV8" } else { $emittente = "Sky" }
                } elseif ($rawTv -match "TV8") {
                    $emittente = "TV8"
                } else {
                    $emittente = $rawTv
                }
            }
            
            $stato = if ($hour -eq 1 -or $hour -eq 0) { "Orario da definire" } else { "Confermato" }
            $key = "$($comp.ShortName)|$matchClean"
            
            # Aggiorna o aggiunge nel master dictionary (preserva lo storico)
            $masterMatches[$key] = [PSCustomObject]@{
                Competizione = $comp.ShortName
                PlannedStartDate = $dtStart.ToString("dd/MM/yyyy HH:mm")
                PlannedEndDate = $dtEnd.ToString("dd/MM/yyyy HH:mm")
                Partita = $matchClean
                Emittente = $emittente
                Stato_Programmazione = $stato
                SortDate = $dtStart
            }
            $scrapedCount++
        }
    }
}

# 3. Ordinamento cronologico complessivo
$sorted = $masterMatches.Values | Sort-Object SortDate

# 4. Generazione righe CSV con delimitatore "," e Competizione per prima colonna
$csvLines = [System.Collections.Generic.List[string]]::new()
$csvLines.Add("Competizione,Planned Start Date,Planned End Date,Partita,Emittente,Stato_Programmazione")

foreach ($m in $sorted) {
    $comp = '"' + $m.Competizione.Replace('"', '""') + '"'
    $partita = '"' + $m.Partita.Replace('"', '""') + '"'
    $emittente = '"' + $m.Emittente.Replace('"', '""') + '"'
    $stato = '"' + $m.Stato_Programmazione.Replace('"', '""') + '"'
    
    $line = "$comp,$($m.PlannedStartDate),$($m.PlannedEndDate),$partita,$emittente,$stato"
    $csvLines.Add($line)
}

# 5. Salvataggio con UTF-8 BOM
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($csvPath, $csvLines, $utf8Bom)

Write-Host "Aggiornamento completato con successo!" -ForegroundColor Green
Write-Host "Partite scaricate in questo ciclo: $scrapedCount"
Write-Host "Totale partite mantenute (incluso storico): $($sorted.Count)"
Write-Host "File salvato in: $csvPath"
