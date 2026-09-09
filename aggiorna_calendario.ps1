# ==============================================================================
# Script Ibrido: CalcioInTV (Primario per Emittenti) + TheSportsDB (Arbitro Orari)
# Formato output personalizzato:
# Number,Planned Start Date,Planned End Date,Short Description,State,Type,Impatto,Note
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
            $comp = if ($row.Number) { $row.Number } else { $row.Competizione }
            $match = if ($row.'Short Description') { $row.'Short Description' } else { $row.Partita }
            $start = $row.'Planned Start Date'
            $end = $row.'Planned End Date'
            $emittente = if ($row.Note) { $row.Note } else { $row.Emittente }
            $state = if ($row.State) { $row.State } else { "" }
            $type = if ($row.Type) { $row.Type } else { "" }
            $impatto = if ($row.Impatto) { $row.Impatto } else { "" }
            
            if ($comp -and $match) {
                $key = "$comp|$match"
                
                $dtSort = [datetime]::MinValue
                if ($start -and [datetime]::TryParseExact($start, "dd/MM/yyyy HH:mm", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dtSort)) {
                    # parsed
                }
                
                $masterMatches[$key] = [PSCustomObject]@{
                    Number = $comp
                    PlannedStartDate = $start
                    PlannedEndDate = $end
                    ShortDescription = $match
                    State = $state
                    Type = $type
                    Impatto = $impatto
                    Note = $emittente
                    SortDate = $dtSort
                }
            }
        }
        Write-Host "Caricate $($masterMatches.Count) partite dallo storico esistente." -ForegroundColor Cyan
    } catch {
        Write-Warning "Avviso nella lettura dello storico esistente: $_"
    }
}

# 2. Scarica i dati aggiornati da TheSportsDB per convalidare date e orari ufficiali
Write-Host "Interrogazione TheSportsDB per orari e date ufficiali..." -ForegroundColor Cyan
$officialLookup = @{}
$romeZone = [TimeZoneInfo]::FindSystemTimeZoneById("W. Europe Standard Time")

# Interroghiamo i turni imminenti di Serie A (es. giornate 4..12)
foreach ($r in 4..12) {
    $u = "https://www.thesportsdb.com/api/v1/json/3/eventsround.php?id=4332&r=$r&s=2026-2027"
    try {
        $respRound = (Invoke-RestMethod -Uri $u -TimeoutSec 5).events
        if ($respRound) {
            foreach ($ev in $respRound) {
                if ($ev.dateEvent -and $ev.strTime -and $ev.strTime -ne "00:00:00") {
                    $t1 = $ev.strHomeTeam -replace '\bAC Milan\b', 'Milan' -replace '\bInter Milan\b', 'Inter'
                    $t2 = $ev.strAwayTeam -replace '\bAC Milan\b', 'Milan' -replace '\bInter Milan\b', 'Inter'
                    $pair = "$t1 - $t2"
                    
                    try {
                        $utcStr = "$($ev.dateEvent) $($ev.strTime)"
                        $utcDt = [datetime]::ParseExact($utcStr, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
                        $localDt = [TimeZoneInfo]::ConvertTimeFromUtc($utcDt, $romeZone)
                        $officialLookup["Serie A|$pair"] = $localDt
                    } catch {}
                }
            }
        }
    } catch {}
}
Write-Host "Mappate $($officialLookup.Count) partite ufficiali da TheSportsDB per la convalida." -ForegroundColor Green

# 3. Scarica i dati primari da CalcioInTV (Emittenti TV e calendario completo)
$competitions = @(
    @{ OrigName = "Serie A"; ShortName = "Serie A"; Url = "https://www.calciointv.com/indexfi.php?comp=Serie%20A" },
    @{ OrigName = "Champions League"; ShortName = "UCL"; Url = "https://www.calciointv.com/indexfi.php?comp=Champions%20League" },
    @{ OrigName = "Europa League"; ShortName = "UEL"; Url = "https://www.calciointv.com/indexfi.php?comp=Europa%20League" }
)

$scrapedCount = 0
$correctedCount = 0

foreach ($comp in $competitions) {
    Write-Host "Scaricamento $($comp.ShortName) da CalcioInTV..." -ForegroundColor Cyan
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
            $rawMatch = [System.Net.WebUtility]::HtmlDecode($mText.Groups[1].Value.Trim())
            $rawTv = if ($mTv.Success) { [System.Net.WebUtility]::HtmlDecode($mTv.Groups[1].Value.Trim()) } else { "Da definire" }
            
            # Normalizzazione nomi squadre
            $matchClean = $rawMatch -replace '\bInter Milan\b', 'Inter' -replace '\bAC Milan\b', 'Milan'
            $key = "$($comp.ShortName)|$matchClean"
            
            # Controllo e riconciliazione con TheSportsDB
            if ($officialLookup.ContainsKey($key)) {
                $officialDt = $officialLookup[$key]
                if ($officialDt -ne $dtStart) {
                    # Orario ufficiale trovato e diverso dal segnaposto
                    $dtStart = $officialDt
                    $correctedCount++
                }
            }
            
            $dtEnd = $dtStart.AddHours(2)
            
            # Logica emittente da CalcioInTV (primaria)
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
            
            $existingState = if ($masterMatches.ContainsKey($key)) { $masterMatches[$key].State } else { "" }
            $existingType = if ($masterMatches.ContainsKey($key)) { $masterMatches[$key].Type } else { "" }
            $existingImpatto = if ($masterMatches.ContainsKey($key)) { $masterMatches[$key].Impatto } else { "" }
            
            $masterMatches[$key] = [PSCustomObject]@{
                Number = $comp.ShortName
                PlannedStartDate = $dtStart.ToString("dd/MM/yyyy HH:mm")
                PlannedEndDate = $dtEnd.ToString("dd/MM/yyyy HH:mm")
                ShortDescription = $matchClean
                State = $existingState
                Type = $existingType
                Impatto = $existingImpatto
                Note = $emittente
                SortDate = $dtStart
            }
            $scrapedCount++
        }
    }
}

# 4. Ordinamento cronologico complessivo
$sorted = $masterMatches.Values | Sort-Object SortDate

# 5. Generazione righe CSV secondo il formato template
$csvLines = [System.Collections.Generic.List[string]]::new()
$csvLines.Add("Number,Planned Start Date,Planned End Date,Short Description,State,Type,Impatto,Note")

foreach ($m in $sorted) {
    $num = '"' + $m.Number.Replace('"', '""') + '"'
    $desc = '"' + $m.ShortDescription.Replace('"', '""') + '"'
    $state = if ($m.State) { '"' + $m.State.Replace('"', '""') + '"' } else { "" }
    $type = if ($m.Type) { '"' + $m.Type.Replace('"', '""') + '"' } else { "" }
    $impatto = if ($m.Impatto) { '"' + $m.Impatto.Replace('"', '""') + '"' } else { "" }
    $note = '"' + $m.Note.Replace('"', '""') + '"'
    
    $line = "$num,$($m.PlannedStartDate),$($m.PlannedEndDate),$desc,$state,$type,$impatto,$note"
    $csvLines.Add($line)
}

# 6. Salvataggio con UTF-8 BOM
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($csvPath, $csvLines, $utf8Bom)

Write-Host "Aggiornamento completato con successo!" -ForegroundColor Green
Write-Host "Partite scaricate: $scrapedCount"
Write-Host "Date/Orari corretti da TheSportsDB: $correctedCount"
Write-Host "Totale partite mantenute nel CSV: $($sorted.Count)"
Write-Host "File salvato in: $csvPath"
