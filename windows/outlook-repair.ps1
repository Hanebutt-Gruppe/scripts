# Initialize results tracking
$results = @{
    ProcessesStopped = $false
    AADBrokerRegistered = $false
    OutlookAppRegistered = $false
    ServicesRestarted = $false
    Errors = @()
    Warnings = @()
}

# 1. Outlook-Prozesse sicher beenden
Write-Host "`n[1/3] " -NoNewline -ForegroundColor Cyan
Write-Host "Sicheres Beenden von Outlook-Prozessen..." -ForegroundColor White
try {
    $processes = @("OUTLOOK", "OLK", "MSOIDSVC", "MSOIDSVCM")
    $stoppedCount = 0
    
    foreach ($processName in $processes) {
        $runningProcesses = Get-Process $processName -ErrorAction SilentlyContinue
        if ($runningProcesses) {
            Write-Host "  → Beende $processName Prozess(e)..." -ForegroundColor Yellow
            $runningProcesses | Stop-Process -Force -ErrorAction Stop
            Write-Host "  ✓ $processName Prozess(e) beendet" -ForegroundColor Green
            $stoppedCount++
        }
    }
    
    if ($stoppedCount -eq 0) {
        Write-Host "  ℹ Keine Outlook-Prozesse liefen" -ForegroundColor Gray
    }
    
    Start-Sleep -Seconds 3
    $results.ProcessesStopped = $true
    Write-Host "  ✓ Schritt 1 erfolgreich abgeschlossen" -ForegroundColor Green
} catch {
    $errorMsg = "Fehler beim Beenden der Prozesse: $($_.Exception.Message)"
    Write-Host "  ✗ $errorMsg" -ForegroundColor Red
    $results.Errors += $errorMsg
}

# 2. AAD Broker Plugin und Outlook-App registrieren
Write-Host "`n[2/3] " -NoNewline -ForegroundColor Cyan
Write-Host "Microsoft-Apps werden neu registriert..." -ForegroundColor White

try {
    # Erst alle relevanten Broker-Prozesse beenden
    Write-Host "  → Beende Broker-Prozesse..." -ForegroundColor Yellow
    $brokerProcesses = @("RuntimeBroker", "BackgroundTaskHost", "ApplicationFrameHost")
    $brokerStopped = 0
    
    foreach ($processName in $brokerProcesses) {
        $processes = Get-Process $processName -ErrorAction SilentlyContinue
        if ($processes) {
            $processes | Stop-Process -Force -ErrorAction SilentlyContinue
            $brokerStopped++
        }
    }
    
    if ($brokerStopped -gt 0) {
        Write-Host "  ✓ $brokerStopped Broker-Prozess(e) beendet" -ForegroundColor Green
    }
    
    Start-Sleep -Seconds 2
    
    # AAD Broker Plugin mit verbesserter Fehlerbehandlung
    Write-Host "  → Registriere AAD BrokerPlugin..." -ForegroundColor Yellow
    try {
        # Erst versuchen zu resetten falls bereits installiert
        $aadPackage = Get-AppxPackage Microsoft.AAD.BrokerPlugin -ErrorAction SilentlyContinue
        if ($aadPackage) {
            Reset-AppxPackage Microsoft.AAD.BrokerPlugin -ErrorAction SilentlyContinue
            Write-Host "    ✓ AAD BrokerPlugin zurückgesetzt" -ForegroundColor Green
            Start-Sleep -Seconds 2
        }
        
        # Dann neu registrieren
        $aadBrokerPath = "C:\Windows\SystemApps\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AppxManifest.xml"
        if (Test-Path $aadBrokerPath) {
            Add-AppxPackage -Register $aadBrokerPath -DisableDevelopmentMode -ErrorAction Stop
            Write-Host "  ✓ AAD BrokerPlugin erfolgreich registriert" -ForegroundColor Green
            $results.AADBrokerRegistered = $true
        } else {
            $warningMsg = "AAD BrokerPlugin-Pfad nicht gefunden"
            Write-Host "  ⚠ $warningMsg" -ForegroundColor Yellow
            $results.Warnings += $warningMsg
        }
    } catch {
        $errorMsg = "AAD BrokerPlugin Registrierung fehlgeschlagen: $($_.Exception.Message)"
        Write-Host "  ✗ $errorMsg" -ForegroundColor Red
        $results.Errors += $errorMsg
        Write-Host "  ℹ Ein Neustart könnte erforderlich sein" -ForegroundColor Gray
    }
    
    # Outlook-App neu registrieren mit verbesserter Fehlerbehandlung
    Write-Host "  → Registriere Outlook-App..." -ForegroundColor Yellow
    $outlookAppPaths = @(
        "C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_*\AppxManifest.xml",
        "C:\Windows\SystemApps\Microsoft.OutlookForWindows_*\AppxManifest.xml"
    )
    
    $outlookRegistered = $false
    foreach ($pattern in $outlookAppPaths) {
        $paths = Get-ChildItem $pattern -ErrorAction SilentlyContinue
        foreach ($path in $paths) {
            try {
                # Erst versuchen zu resetten
                $appName = ($path.Directory.Name -split "_")[0]
                $existingApp = Get-AppxPackage $appName -ErrorAction SilentlyContinue
                if ($existingApp) {
                    Reset-AppxPackage $appName -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 1
                }
                
                Add-AppxPackage -Register $path.FullName -DisableDevelopmentMode -ErrorAction Stop
                Write-Host "  ✓ Outlook-App registriert: $($path.Directory.Name)" -ForegroundColor Green
                $outlookRegistered = $true
            } catch {
                $errorMsg = "Outlook-App Registrierung fehlgeschlagen: $($path.Name) - $($_.Exception.Message)"
                Write-Host "  ✗ $errorMsg" -ForegroundColor Red
                $results.Errors += $errorMsg
            }
        }
    }
    
    if ($outlookRegistered) {
        $results.OutlookAppRegistered = $true
        Write-Host "  ✓ Schritt 2 erfolgreich abgeschlossen" -ForegroundColor Green
    } else {
        $warningMsg = "Keine Outlook-App gefunden oder registriert"
        Write-Host "  ⚠ $warningMsg" -ForegroundColor Yellow
        $results.Warnings += $warningMsg
    }
    
} catch {
    $errorMsg = "Fehler bei App-Registrierung: $($_.Exception.Message)"
    Write-Host "  ✗ $errorMsg" -ForegroundColor Red
    $results.Errors += $errorMsg
    Write-Host "  → Versuche alternative Reparaturmethode..." -ForegroundColor Yellow
    
    # Alternative: Services neu starten
    Write-Host "`n[3/3] " -NoNewline -ForegroundColor Cyan
    Write-Host "Starte Windows-Services neu..." -ForegroundColor White
    try {
        $services = @("AppXSvc", "StateRepository", "TokenBroker")
        $servicesRestarted = 0
        
        foreach ($service in $services) {
            $svc = Get-Service $service -ErrorAction SilentlyContinue
            if ($svc) {
                if ($svc.Status -eq "Running") {
                    Restart-Service $service -Force -ErrorAction Stop
                    Write-Host "  ✓ Service $service neu gestartet" -ForegroundColor Green
                    $servicesRestarted++
                } else {
                    Write-Host "  ℹ Service $service läuft nicht" -ForegroundColor Gray
                }
            } else {
                Write-Host "  ⚠ Service $service nicht gefunden" -ForegroundColor Yellow
            }
        }
        
        if ($servicesRestarted -gt 0) {
            $results.ServicesRestarted = $true
            Write-Host "  ✓ $servicesRestarted Service(s) erfolgreich neu gestartet" -ForegroundColor Green
        }
    } catch {
        $errorMsg = "Service-Neustart fehlgeschlagen: $($_.Exception.Message)"
        Write-Host "  ✗ $errorMsg" -ForegroundColor Red
        $results.Errors += $errorMsg
    }
}

# Summary and results
Write-Host "`n" + "="*70 -ForegroundColor Cyan
Write-Host "ZUSAMMENFASSUNG" -ForegroundColor Cyan
Write-Host "="*70 -ForegroundColor Cyan

$successCount = 0
$totalSteps = 4

Write-Host "`nStatus der Reparaturschritte:" -ForegroundColor White
Write-Host "  " -NoNewline
if ($results.ProcessesStopped) { Write-Host "✓" -NoNewline -ForegroundColor Green; $successCount++ } else { Write-Host "✗" -NoNewline -ForegroundColor Red }
Write-Host " Outlook-Prozesse beendet"

Write-Host "  " -NoNewline
if ($results.AADBrokerRegistered) { Write-Host "✓" -NoNewline -ForegroundColor Green; $successCount++ } else { Write-Host "✗" -NoNewline -ForegroundColor Red }
Write-Host " AAD BrokerPlugin registriert"

Write-Host "  " -NoNewline
if ($results.OutlookAppRegistered) { Write-Host "✓" -NoNewline -ForegroundColor Green; $successCount++ } else { Write-Host "✗" -NoNewline -ForegroundColor Red }
Write-Host " Outlook-App registriert"

Write-Host "  " -NoNewline
if ($results.ServicesRestarted) { Write-Host "✓" -NoNewline -ForegroundColor Green; $successCount++ } else { Write-Host "✗" -NoNewline -ForegroundColor Red }
Write-Host " Services neu gestartet"

# Overall status
Write-Host "`nGesamtergebnis: " -NoNewline
if ($successCount -eq $totalSteps) {
    Write-Host "ERFOLGREICH" -ForegroundColor Green
    Write-Host "Alle Schritte wurden erfolgreich abgeschlossen." -ForegroundColor Green
} elseif ($successCount -ge 2) {
    Write-Host "TEILWEISE ERFOLGREICH" -ForegroundColor Yellow
    Write-Host "$successCount von $totalSteps Schritten erfolgreich abgeschlossen." -ForegroundColor Yellow
} else {
    Write-Host "FEHLGESCHLAGEN" -ForegroundColor Red
    Write-Host "Nur $successCount von $totalSteps Schritten erfolgreich abgeschlossen." -ForegroundColor Red
}

# Display warnings if any
if ($results.Warnings.Count -gt 0) {
    Write-Host "`nWarnungen:" -ForegroundColor Yellow
    foreach ($warning in $results.Warnings) {
        Write-Host "  ⚠ $warning" -ForegroundColor Yellow
    }
}

# Display errors if any
if ($results.Errors.Count -gt 0) {
    Write-Host "`nFehler:" -ForegroundColor Red
    foreach ($error in $results.Errors) {
        Write-Host "  ✗ $error" -ForegroundColor Red
    }
}

Write-Host "`n" + "="*70 -ForegroundColor Cyan

# Logout prompt
Write-Host "`n"
$shouldLogout = Read-Host "Möchten Sie sich jetzt abmelden? (J/N)"

if ($shouldLogout -match '^[JjYy]') {
    Write-Host "`nAbmeldung wird vorbereitet..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    try {
        # Log out the user
        logoff
        Write-Host "Abmeldung erfolgreich initiiert." -ForegroundColor Green
    } catch {
        Write-Host "Fehler bei der Abmeldung: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Bitte melden Sie sich manuell ab." -ForegroundColor Yellow
    }
} else {
    Write-Host "`nBitte denken Sie daran, sich abzumelden, bevor Sie Outlook erneut öffnen." -ForegroundColor Yellow
}

Write-Host "`nSkript beendet." -ForegroundColor Cyan