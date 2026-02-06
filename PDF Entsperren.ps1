<#
.SYNOPSIS
   Entblockt alle PDF-Dateien in einem angegebenen Pfad, sofern der Benutzer die nÃ¶tigen Berechtigungen hat.

.DESCRIPTION
   PrÃ¼ft fÃ¼r jede PDF-Datei, ob der Benutzer Lese- und Schreibrechte besitzt.
   Entblockt nur Dateien mit ausreichenden Rechten und gibt eine Ãœbersicht aus.

.PARAMETER Path
   Der Basisordner, in dem gesucht werden soll (z. B. "P:\").
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

Write-Host "Starte PrÃ¼fung und Entblocken von PDF-Dateien unter: $Path" -ForegroundColor Cyan
Write-Host ""

# PrÃ¼fe, ob Pfad existiert
if (-not (Test-Path $Path)) {
    Write-Host "Fehler: Der angegebene Pfad existiert nicht." -ForegroundColor Red
    exit
}

# Ergebnisse speichern
$Results = @()

# Alle PDF-Dateien finden
$files = Get-ChildItem -Path $Path -Recurse -Filter *.pdf -ErrorAction SilentlyContinue

if ($files.Count -eq 0) {
    Write-Host "Keine PDF-Dateien gefunden." -ForegroundColor Yellow
    exit
}

foreach ($file in $files) {
    $hasRead = $false
    $hasWrite = $false
    $canUnblock = $false

    try {
        # Teste Leseberechtigung
        $stream = $null
        try {
            $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Read')
            $hasRead = $true
        } catch {}
        finally {
            if ($stream) { $stream.Close() }
        }

        # Teste Schreibberechtigung (Datei Ã¶ffnen im Schreibmodus)
        $stream = $null
        try {
            $stream = [System.IO.File]::Open($file.FullName, 'Open', 'Write')
            $hasWrite = $true
        } catch {}
        finally {
            if ($stream) { $stream.Close() }
        }

        $canUnblock = $hasRead -and $hasWrite

        if ($canUnblock) {
            Unblock-File -Path $file.FullName -ErrorAction SilentlyContinue
            $status = "Entblockt"
            $color = "Green"
        } else {
            $status = "Keine Berechtigung"
            $color = "Red"
        }

        Write-Host ("[{0}] {1}" -f $status, $file.FullName) -ForegroundColor $color

        $Results += [pscustomobject]@{
            Datei = $file.FullName
            Lesen = $hasRead
            Schreiben = $hasWrite
            Status = $status
        }
    } catch {
        Write-Host ("[Fehler] {0}" -f $file.FullName) -ForegroundColor Yellow
    }
}

# Zusammenfassung
Write-Host ""
Write-Host "Fertig. Zusammenfassung:" -ForegroundColor Cyan
$Results | Group-Object -Property Status | ForEach-Object {
    Write-Host ("{0}: {1}" -f $_.Name, $_.Count)
}

# Optional: Ergebnisse als CSV exportieren
#$Results | Export-Csv -Path "$env:USERPROFILE\UnblockResults.csv" -NoTypeInformation -Encoding UTF8

