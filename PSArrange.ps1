<#
.SYNOPSIS
    Organisiert Dateien und Ordner umfassend basierend auf einer detaillierten JSON-Konfiguration
    oder automatisch erkannten "nicht-risikanten" Verzeichnissen, mit Unterstützung für globale Ausschlüsse
    und erweiterte Dateiregeln inklusive Duplicate Handling.
    Bietet flexible Regeln für Dateitypen, Ordnerstruktur (inkl. Subfolder), Verschieben und
    Umbenennen von Ordnern, Protokollierung und optionale Vorschau.
.DESCRIPTION
    Dieses Skript organisiert Dateien und Ordner in den angegebenen Verzeichnissen gemäß einer
    umfassenden JSON-Konfigurationsdatei. Wenn keine Verzeichnisse in der Konfiguration
    angegeben sind, versucht das Skript, "nicht-riskante" Verzeichnisse automatisch zu erkennen.
    Es werden globale Ausschlusslisten aus der Konfiguration berücksichtigt, um bestimmte
    Dateien und Ordner von der Bearbeitung auszuschließen. Dateiregeln können nun auch
    auf Dateinamenmuster und Alter basieren. Das Skript unterstützt konfigurierbares Verhalten
    bei Dateinamenkonflikten (Duplicate Handling).
.PARAMETER ConfigFile
    Der Pfad zur JSON-Konfigurationsdatei, die alle Organisationsregeln enthält.
    Dieser Parameter ist obligatorisch.
.PARAMETER Directories
    Ein Array von Verzeichnissen, die organisiert werden sollen. Standardmäßig (wenn nicht in der
    Konfigurationsdatei angegeben) versucht das Skript, "nicht-riskante" Verzeichnisse automatisch
    zu erkennen. Dieser Parameter überschreibt die automatische Erkennung, wenn er direkt
    beim Skriptaufruf verwendet wird.
.PARAMETER DeleteEmptyFolders
    Ein Switch-Parameter, der angibt, ob nach der Organisation leere Ordner
    in den bearbeiteten Verzeichnissen gelöscht werden sollen.
.PARAMETER Preview
    Ein Switch-Parameter, der angibt, ob die Aktionen nur als Vorschau im Terminal
    angezeigt und nicht tatsächlich durchgeführt werden sollen.
.EXAMPLE
    .\organisieren.ps1 -ConfigFile ".\organisation_config.json"
    Verwendet die Konfigurationsdatei zur Organisation der dort gelisteten Verzeichnisse,
    unter Berücksichtigung globaler Ausschlüsse, erweiterter Dateiregeln und Duplicate Handling.
.EXAMPLE
    .\organisieren.ps1 -ConfigFile ".\meine_regeln.json" -DeleteEmptyFolders -Preview
    Zeigt eine Vorschau der Organisation und des Löschens leerer Ordner an, basierend auf
    den Verzeichnissen in 'meine_regeln.json' und globalen Einstellungen.
.EXAMPLE
    .\organisieren.ps1 -ConfigFile ".\config_automatisch.json"
    Verwendet die Konfiguration, um "nicht-riskante" Verzeichnisse automatisch zu erkennen
    und zu organisieren (vorausgesetzt, "Directories" ist in der JSON leer oder fehlt),
    unter Berücksichtigung globaler Einstellungen.
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$ConfigFile,

    [string[]]$Directories = @(), # Standardmäßig leer, automatische Erkennung, wenn nicht in JSON

    [switch]$DeleteEmptyFolders,

    [switch]$Preview
)

# --- Initialisierung ---
# Sicherstellen, dass das Skriptverzeichnis der aktuelle Ort ist (für relative Pfade)
# Push-Location $PSScriptRoot

# Pfad zur Protokolldatei
$logFile = Join-Path (Split-Path -Path $MyInvocation.MyCommand.Path -Parent) "organisation.log"

# Globale Einstellungen (werden aus Konfigurationsdatei geladen)
$globalFileExclusionPatterns = @()
$globalFolderExclusionPatterns = @()
$globalDuplicateHandling = "Skip" # Standardwert, falls nicht in Konfig geladen

# --- Kernfunktionen ---

# Funktion zum Schreiben in die Protokolldatei
function Write-Log ($message) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    try {
        # Versuche, mit UTF8 zu schreiben, was für die meisten Fälle gut ist.
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8 -ErrorAction Stop
    } catch {
        try {
            # Fallback auf Default-Encoding, falls UTF8 fehlschlägt (z.B. wegen Berechtigungen, die nichts mit Encoding zu tun haben)
            Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
        } catch {
            Write-Warning "FEHLER beim Schreiben in die Logdatei '$logFile': $($_.Exception.Message)"
        }
    }
    # Gib die Nachricht auch auf der Konsole aus
    Write-Host $logEntry
}

# Funktion zum Überprüfen, ob ein Pfad ausgeschlossen werden soll
function Should-ExcludePath ($path, [switch]$IsFolder) {
    $patternsToUse = if ($IsFolder.IsPresent) { $globalFolderExclusionPatterns } else { $globalFileExclusionPatterns }

    if (-not $patternsToUse -or $patternsToUse.Count -eq 0) {
        return $false # Keine Ausschlussmuster definiert oder die Liste ist leer
    }

    # Handle potentielle Null- oder Leer-Pfade
    if ([string]::IsNullOrWhiteSpace($path)) { return $false }

    # Normalisiere den Pfad für konsistente Vergleiche (z.B. Backslashes, Kleinschreibung)
    try {
        $normalizedPath = $path.Replace('/', '\').ToLower()
    } catch {
        Write-Log "WARNUNG: Konnte Pfad '$path' für Ausschlussprüfung nicht normalisieren. Fehler: $($_.Exception.Message)"
        return $false # Im Zweifel nicht ausschließen
    }

    foreach ($pattern in $patternsToUse) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        try {
            # Normalisiere das Muster ebenfalls
            $normalizedPattern = $pattern.Replace('/', '\').ToLower()
            if ($normalizedPath -like $normalizedPattern) {
                return $true # Pfad matcht ein Ausschlussmuster
            }
        } catch {
            Write-Log "WARNUNG: Fehler beim Verarbeiten des Ausschlussmusters '$pattern' für Pfad '$path'. Fehler: $($_.Exception.Message)"
            # Fahre mit dem nächsten Muster fort
        }
    }
    return $false # Kein Ausschlussmuster hat gematcht
}

# Funktion zum Erstellen von Ordnern (mit Protokollierung und Vorschau)
function Create-FolderIfNotExists ($folderPath) {
    # Prüfe auf ungültige Pfade
    if ([string]::IsNullOrWhiteSpace($folderPath)) {
        Write-Log "FEHLER: Ungültiger (leerer) Ordnerpfad für Create-FolderIfNotExists angegeben."
        return $false
    }

    if (Should-ExcludePath $folderPath -IsFolder) {
        Write-Log "INFO: Überspringe Erstellung des ausgeschlossenen Ordners: '$folderPath'"
        return $false # Gib false zurück, wenn nicht erstellt, da ausgeschlossen
    }

    if (-not (Test-Path -Path $folderPath -PathType Container)) {
        $message = "PREVIEW: Erstelle Ordner: '$folderPath'"
        if (-not $Preview) {
            $message = "AKTION: Erstelle Ordner: '$folderPath'"
        }
        Write-Log $message

        if (-not $Preview) {
            try {
                New-Item -Path $folderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                return $true # Gib true zurück, wenn erfolgreich erstellt
            } catch {
                Write-Warning "Fehler beim Erstellen des Ordners '$folderPath': $($_.Exception.Message)"
                Write-Log "FEHLER: Fehler beim Erstellen des Ordners '$folderPath': $($_.Exception.Message)"
                return $false # Gib false zurück, wenn Erstellung fehlschlägt
            }
        } else {
            return $true # Im Preview-Modus geben wir true zurück, da die Aktion simuliert wird
        }
    }
    return $true # Gib true zurück, wenn der Ordner bereits existiert
}

# Funktion zum Verschieben von Dateien (mit Protokollierung, Preview, Duplicate Handling und Ausschlussprüfung)
function Move-File ($sourcePath, $destinationPath, $duplicateHandlingStrategy) {
    # Prüfe auf ungültige Pfade
    if ([string]::IsNullOrWhiteSpace($sourcePath) -or [string]::IsNullOrWhiteSpace($destinationPath)) {
         Write-Log "FEHLER: Ungültiger Quell- ('$sourcePath') oder Zielpfad ('$destinationPath') für Move-File angegeben."
         return
    }

    # Prüfe Quell-Datei auf Ausschluss
    if (Should-ExcludePath $sourcePath -IsFolder:$false) {
        Write-Log "INFO: Überspringe Verschieben der ausgeschlossenen Quelldatei: '$sourcePath'"
        return
    }

     # Überprüfe Zielordner und -pfad auf Ausschluss
    $destinationFolder = Split-Path -Path $destinationPath -Parent
    if ([string]::IsNullOrWhiteSpace($destinationFolder)) {
         Write-Log "FEHLER: Konnte Zielordner aus '$destinationPath' nicht extrahieren."
         return
    }
     if (Should-ExcludePath $destinationFolder -IsFolder) {
        Write-Log "WARNUNG: Zielordner '$destinationFolder' für Dateiverschiebung ist ausgeschlossen. Überspringe Verschieben der Datei '$sourcePath'."
        return
    }
     if (Should-ExcludePath $destinationPath -IsFolder:$false) {
        Write-Log "WARNUNG: Vollständiger Zielpfad '$destinationPath' für Dateiverschiebung ist ausgeschlossen. Überspringe Verschieben der Datei '$sourcePath'."
        return
    }

    # Behandle Dateikonflikte
    $destinationFileExists = Test-Path -Path $destinationPath -PathType Leaf # Prüft, ob Zieldatei existiert

    if ($destinationFileExists) {
        $strategy = if ($duplicateHandlingStrategy) { $duplicateHandlingStrategy.ToLower() } else { "skip" }

        switch ($strategy) {
            "skip" {
                Write-Log "WARNUNG: Zieldatei '$destinationPath' existiert bereits. Überspringe '$sourcePath' (Duplicate Handling: Skip)."
            }
            "renamewithtimestamp" {
                $timestamp = Get-Date -Format "yyyyMMddHHmmssfff" # Millisekunden hinzugefügt für höhere Eindeutigkeit
                $sourceFileItem = Get-Item -Path $sourcePath -ErrorAction SilentlyContinue
                if (-not $sourceFileItem) {
                    Write-Log "FEHLER: Quelldatei '$sourcePath' nicht gefunden. Kann nicht umbenannt und verschoben werden."
                    break
                }
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sourceFileItem.Name)
                $extension = $sourceFileItem.Extension # Beinhaltet den Punkt
                $newFileName = "${baseName}_${timestamp}${extension}"
                $newDestinationPath = Join-Path -Path $destinationFolder -ChildPath $newFileName

                # Prüfe neuen Pfad auf Ausschluss
                if (Should-ExcludePath $newDestinationPath -IsFolder:$false) {
                     Write-Log "WARNUNG: Generierter Zieldateiname mit Zeitstempel '$newDestinationPath' ist global ausgeschlossen. Überspringe '$sourcePath'."
                     break
                }

                # Prüfe, ob der *neue* Name auch schon existiert (extrem unwahrscheinlich, aber sicher ist sicher)
                if (Test-Path -Path $newDestinationPath -PathType Leaf) {
                    Write-Log "WARNUNG: Der Zieldateiname mit Zeitstempel '$newDestinationPath' existiert ebenfalls bereits. Überspringe '$sourcePath'."
                } else {
                    $message = "PREVIEW: Verschiebe '$sourcePath' nach '$newDestinationPath' (Duplicate Handling: RenameWithTimestamp)"
                    if (-not $Preview) {
                         $message = "AKTION: Verschiebe '$sourcePath' nach '$newDestinationPath' (Duplicate Handling: RenameWithTimestamp)"
                         try {
                            Move-Item -Path $sourcePath -Destination $newDestinationPath -Force -ErrorAction Stop
                         } catch {
                             $message = "FEHLER: Fehler beim Verschieben von '$sourcePath' nach '$newDestinationPath': $($_.Exception.Message)"
                             Write-Warning $message
                         }
                    }
                    Write-Log $message
                }
            }
            "overwrite" {
                $message = "PREVIEW: Überschreibe '$destinationPath' mit '$sourcePath' (Duplicate Handling: Overwrite)"
                if (-not $Preview) {
                    $message = "AKTION: Überschreibe '$destinationPath' mit '$sourcePath' (Duplicate Handling: Overwrite)"
                     try {
                        Move-Item -Path $sourcePath -Destination $destinationPath -Force -ErrorAction Stop
                     } catch {
                         $message = "FEHLER: Fehler beim Überschreiben von '$destinationPath' mit '$sourcePath': $($_.Exception.Message)"
                         Write-Warning $message
                     }
                }
                Write-Log $message
            }
            "ask" {
                 Write-Log "WARNUNG: Duplicate Handling Strategie 'Ask' ist im nicht-interaktiven Modus nicht implementiert. Überspringe Datei '$sourcePath'."
                 # Behandle wie "skip"
            }
            default {
                 Write-Log "WARNUNG: Unbekannte Duplicate Handling Strategie '$strategy' für Datei '$sourcePath'. Verwende 'Skip'."
                 # Behandle wie "skip"
            }
        } # End switch
    } else {
        # Zieldatei existiert nicht, normal verschieben
        $message = "PREVIEW: Verschiebe '$sourcePath' nach '$destinationPath'"
        if (-not $Preview) {
            $message = "AKTION: Verschiebe '$sourcePath' nach '$destinationPath'"
            try {
                Move-Item -Path $sourcePath -Destination $destinationPath -Force -ErrorAction Stop
            } catch {
                $message = "FEHLER: Fehler beim Verschieben von '$sourcePath' nach '$destinationPath': $($_.Exception.Message)"
                Write-Warning $message
            }
        }
        Write-Log $message
    }
}

# Funktion zum Umbenennen von Ordnern (mit Protokollierung und Vorschau)
function Rename-Folder ($oldPath, $newPath) {
    if ([string]::IsNullOrWhiteSpace($oldPath) -or [string]::IsNullOrWhiteSpace($newPath)) {
         Write-Log "FEHLER: Ungültiger alter ('$oldPath') oder neuer Pfad ('$newPath') für Rename-Folder angegeben."
         return
    }
    if ($oldPath -eq $newPath) { return } # Nichts zu tun

    if (Should-ExcludePath $oldPath -IsFolder) {
        Write-Log "INFO: Überspringe Umbenennen des ausgeschlossenen Ordners: '$oldPath'"
        return
    }
    if (Should-ExcludePath $newPath -IsFolder) {
        Write-Log "WARNUNG: Zielname für Umbenennung '$newPath' führt zu ausgeschlossenem Pfad. Überspringe Umbenennen von '$oldPath'."
        return
    }

    $message = "PREVIEW: Benenne Ordner '$oldPath' in '$newPath' um"
    if (-not $Preview) {
        $message = "AKTION: Benenne Ordner '$oldPath' in '$newPath' um"
        try {
            # Extrahiere nur den neuen Namensteil für -NewName
            $newName = Split-Path -Path $newPath -Leaf
            if ([string]::IsNullOrWhiteSpace($newName)) { throw "Konnte neuen Namen nicht aus '$newPath' extrahieren." }
            Rename-Item -Path $oldPath -NewName $newName -Force -ErrorAction Stop
        } catch {
            $message = "FEHLER: Fehler beim Umbenennen von '$oldPath' nach '$newPath': $($_.Exception.Message)"
            Write-Warning $message
        }
    }
    Write-Log $message
}

# Funktion zum Verschieben von Ordnern (mit Protokollierung und Vorschau)
function Move-Folder ($oldPath, $newPath) {
    if ([string]::IsNullOrWhiteSpace($oldPath) -or [string]::IsNullOrWhiteSpace($newPath)) {
         Write-Log "FEHLER: Ungültiger alter ('$oldPath') oder neuer Pfad ('$newPath') für Move-Folder angegeben."
         return
    }
     if ($oldPath -eq $newPath) { return } # Nichts zu tun

    if (Should-ExcludePath $oldPath -IsFolder) {
        Write-Log "INFO: Überspringe Verschieben des ausgeschlossenen Quellordners: '$oldPath'"
        return
    }
    if (Should-ExcludePath $newPath -IsFolder) {
        Write-Log "WARNUNG: Zielpfad '$newPath' für Ordnerverschiebung ist ausgeschlossen. Überspringe Verschieben des Ordners '$oldPath'."
        return
    }

    # Sicherstellen, dass der Elternordner des Ziels existiert
    $newParentFolder = Split-Path -Path $newPath -Parent
    if (-not (Test-Path -Path $newParentFolder -PathType Container)) {
        Write-Log "INFO: Erstelle Elternordner '$newParentFolder' für Ordnerverschiebung."
        if (-not (Create-FolderIfNotExists $newParentFolder)) {
            Write-Log "FEHLER: Konnte notwendigen Elternordner '$newParentFolder' nicht erstellen. Überspringe Verschieben von '$oldPath'."
            return
        }
    }

    $message = "PREVIEW: Verschiebe Ordner '$oldPath' nach '$newPath'"
    if (-not $Preview) {
        $message = "AKTION: Verschiebe Ordner '$oldPath' nach '$newPath'"
        try {
            # -Force verwenden, aber beachten, dass dies bei Ordnern anders wirkt als bei Dateien.
            # Es hilft hauptsächlich, wenn das Ziel ein leerer Ordner ist oder um schreibgeschützte Attribute zu überwinden.
            # Es überschreibt NICHT standardmäßig den Inhalt eines vorhandenen, nicht leeren Zielordners.
            Move-Item -Path $oldPath -Destination $newPath -Force -ErrorAction Stop
        } catch {
            $message = "FEHLER: Fehler beim Verschieben von '$oldPath' nach '$newPath': $($_.Exception.Message)"
            Write-Warning $message
        }
    }
    Write-Log $message
}

# Funktion zum Löschen leerer Ordner (mit Protokollierung und Vorschau)
function Remove-EmptyFolders ($directory) {
     if ([string]::IsNullOrWhiteSpace($directory) -or -not (Test-Path -Path $directory -PathType Container)) {
         return # Ungültiges oder nicht existierendes Verzeichnis
     }
     if (Should-ExcludePath $directory -IsFolder) {
        # Write-Log "INFO: Überspringe Prüfung auf leere Ordner in ausgeschlossenem Pfad: '$directory'" # Optional: Weniger verbose
        return
    }

    # Hole Unterordner, die NICHT ausgeschlossen sind
    $subFolders = Get-ChildItem -Path $directory -Directory -ErrorAction SilentlyContinue | Where-Object { -not (Should-ExcludePath $_.FullName -IsFolder) }

    # Rekursiver Aufruf für jeden nicht ausgeschlossenen Unterordner
    foreach ($subFolder in $subFolders) {
        Remove-EmptyFolders $subFolder.FullName
    }

    # Nach der Rekursion: Prüfe, ob der *aktuelle* Ordner ($directory) jetzt leer ist
    # Er muss leer sein (keine Dateien, keine verbleibenden Unterordner) UND darf nicht ausgeschlossen sein.
    if (-not (Should-ExcludePath $directory -IsFolder)) {
        try {
            # Prüfe, ob der Ordner wirklich leer ist (-Force berücksichtigt versteckte/Systemelemente)
            if ((Get-ChildItem -Path $directory -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                $message = "PREVIEW: Lösche leeren Ordner: '$directory'"
                if (-not $Preview) {
                    $message = "AKTION: Lösche leeren Ordner: '$directory'"
                    # Zusätzliche Sicherheitsprüfung: Nicht das Stammverzeichnis eines Laufwerks löschen
                    if ($directory -match '^[a-zA-Z]:\\?$') {
                         Write-Log "WARNUNG: Sicherheitsprüfung verhindert Löschen des Laufwerksstamm '$directory'."
                    } else {
                        try {
                            Remove-Item -Path $directory -Force -ErrorAction Stop
                        } catch {
                             $message = "FEHLER: Fehler beim Löschen des leeren Ordners '$directory': $($_.Exception.Message)"
                             Write-Warning $message
                        }
                    }
                }
                Write-Log $message
            }
        } catch {
            # Fehler beim Prüfen des Ordnerinhalts
             Write-Log "FEHLER: Fehler beim Prüfen des Inhalts von '$directory' zum Löschen leerer Ordner: $($_.Exception.Message)"
        }
    }
}


# Funktion zum Organisieren eines Verzeichnisses
function Organize-Directory ($directory, $config, $dirsToOrganize) {
    # Überprüfen, ob das aktuelle Verzeichnis selbst ausgeschlossen ist oder nicht existiert
    if (-not (Test-Path -Path $directory -PathType Container)) {
         Write-Log "WARNUNG: Zu organisierendes Verzeichnis '$directory' nicht gefunden. Überspringe."
         return
    }
    if (Should-ExcludePath $directory -IsFolder) {
        Write-Log "INFO: Überspringe Organisation in ausgeschlossenem Verzeichnis: '$directory'"
        return
    }

    Write-Log "---- Starte Organisation im Verzeichnis: '$directory' ----"

    # --- Datei-Organisation ---
    if ($null -ne $config.FileRules -and $config.FileRules -is [System.Array] -and $config.FileRules.Count -gt 0) {
        Write-Log "--- Verarbeite Dateien in '$directory' ---"
        # Hole alle Dateien im aktuellen Verzeichnis, die nicht ausgeschlossen sind
        $filesToProcess = Get-ChildItem -Path $directory -File -ErrorAction SilentlyContinue | Where-Object { -not (Should-ExcludePath $_.FullName -IsFolder:$false) }

        foreach ($file in $filesToProcess) {
            $fileName = $file.Name
            $fileExtension = $file.Extension # Behält den Punkt, z.B. ".txt"
            $sourcePath = $file.FullName
            $matchingRule = $null
            # Standard-Duplicate Handling (global oder default "Skip")
            $effectiveDuplicateHandling = $globalDuplicateHandling

            # Finde die erste passende Regel
            foreach ($rule in $config.FileRules) {
                # Überspringe ungültige oder leere Regeln
                 if ($null -eq $rule -or `
                    (($null -eq $rule.Extensions -or $rule.Extensions.Count -eq 0) -and `
                     ($null -eq $rule.NamePatterns -or $rule.NamePatterns.Count -eq 0) -and `
                     ($null -eq $rule.OlderThanDays -or $rule.OlderThanDays -le 0) -and `
                     ($null -eq $rule.NewerThanDays -or $rule.NewerThanDays -le 0))) {
                     continue
                 }

                $isMatch = $true # Annahme: Regel passt, widerlege es bei Bedarf

                # 1. Prüfe Extensions
                if ($isMatch -and $rule.Extensions -is [System.Array] -and $rule.Extensions.Count -gt 0) {
                    $ruleExtensionsLower = $rule.Extensions | ForEach-Object { $_.ToLower() }
                    # Prüfe, ob die aktuelle Dateiendung (klein, mit Punkt) NICHT in der Liste der Regel-Endungen (klein) ist
                    if ($ruleExtensionsLower -notcontains $fileExtension.ToLower()) {
                        $isMatch = !($true) # Workaround für PowerShell $false Erkennungsproblem
                    }
                }

                # 2. Prüfe NamePatterns (nur wenn noch Match)
                if ($isMatch -and $rule.NamePatterns -is [System.Array] -and $rule.NamePatterns.Count -gt 0) {
                    $namePatternMatch = !($true) # Workaround für PowerShell $false Erkennungsproblem
                    foreach ($pattern in $rule.NamePatterns) {
                         if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
                        try {
                            if ($fileName -like $pattern) {
                                $namePatternMatch = $true
                                break
                            }
                        } catch {
                            Write-Log "WARNUNG: Fehler beim Anwenden des Namensmusters '$pattern' auf '$fileName': $($_.Exception.Message)"
                        }
                    }
                    if (-not $namePatternMatch) {
                        $isMatch = !($true) # Workaround für PowerShell $false Erkennungsproblem
                    }
                }

                # 3. Prüfe OlderThanDays (nur wenn noch Match)
                if ($isMatch -and $null -ne $rule.OlderThanDays -and $rule.OlderThanDays -gt 0) {
                    $fileDate = $null
                    if ($file.LastWriteTime) { $fileDate = $file.LastWriteTime }
                    elseif ($file.CreationTime) { $fileDate = $file.CreationTime }

                    if (-not $fileDate -or $fileDate -ge (Get-Date).AddDays(-$rule.OlderThanDays)) {
                        $isMatch = !($true) # Workaround für PowerShell $false Erkennungsproblem
                    }
                }

                # 4. Prüfe NewerThanDays (nur wenn noch Match)
                if ($isMatch -and $null -ne $rule.NewerThanDays -and $rule.NewerThanDays -gt 0) {
                     $fileDate = $null
                     if ($file.LastWriteTime) { $fileDate = $file.LastWriteTime }
                     elseif ($file.CreationTime) { $fileDate = $file.CreationTime }

                    if (-not $fileDate -or $fileDate -le (Get-Date).AddDays(-$rule.NewerThanDays)) {
                        $isMatch = !($true) # Workaround für PowerShell $false Erkennungsproblem
                    }
                }

                # 5. Prüfe Action (nur wenn noch Match)
                # Aktuell wird nur "Move" unterstützt. Andere Actions führen dazu, dass die Regel nicht angewendet wird.
                 $action = if ($rule.Action) { $rule.Action.ToLower() } else { "move" } # Standardaktion ist Move
                 if ($action -ne "move") {
                      if (-not [string]::IsNullOrWhiteSpace($rule.Action)) { # Nur loggen, wenn explizit was anderes angegeben wurde
                           Write-Log "INFO: Regel für '$fileName' matcht, aber Aktion '$($rule.Action)' ist nicht 'Move'. Regel wird nicht angewendet."
                      }
                      $isMatch = !($true) # Workaround für PowerShell $false Erkennungsproblem
                 }


                # Wenn alle Kriterien passen, Regel anwenden und Schleife verlassen
                if ($isMatch) {
                    $matchingRule = $rule
                    # Regel-spezifisches Duplicate Handling überschreibt globales
                    if ($rule.DuplicateHandling) {
                         $effectiveDuplicateHandling = $rule.DuplicateHandling # Behalte Original-Case für spätere Validierung in Move-File
                    }
                    Write-Log "Regel matcht für Datei '$fileName' (Beschreibung: '$($rule.Description)'). Effektives Duplicate Handling: '$effectiveDuplicateHandling'."
                    break # Erste passende Regel gefunden
                }
            } # Ende foreach ($rule in $config.FileRules)

            # --- Aktion basierend auf Regel-Match ausführen ---
            if ($matchingRule) {
                # --- Korrekte Pfad-Erstellung (wie oben korrigiert) ---
                $currentDestinationFolder = $directory
                if (-not [string]::IsNullOrWhiteSpace($matchingRule.TargetFolder)) {
                    $currentDestinationFolder = Join-Path -Path $currentDestinationFolder -ChildPath $matchingRule.TargetFolder
                }
                if (-not [string]::IsNullOrWhiteSpace($matchingRule.SubFolder)) {
                    $currentDestinationFolder = Join-Path -Path $currentDestinationFolder -ChildPath $matchingRule.SubFolder
                }
                if ($matchingRule.OrganizeByDate) {
                    $fileDate = $null
                    if ($file.LastWriteTime) { $fileDate = $file.LastWriteTime }
                    elseif ($file.CreationTime) { $fileDate = $file.CreationTime }
                     if ($fileDate) {
                        $yearFolder = $fileDate.ToString("yyyy")
                        $monthFolder = $fileDate.ToString("yyyy-MM")
                        $currentDestinationFolder = Join-Path -Path $currentDestinationFolder -ChildPath $yearFolder
                        $currentDestinationFolder = Join-Path -Path $currentDestinationFolder -ChildPath $monthFolder
                     } else {
                         Write-Log "WARNUNG: Datum für Datei '$fileName' konnte nicht ermittelt werden für Datumsordner."
                     }
                }
                $destinationFolder = $currentDestinationFolder
                # --- Ende Korrektur Pfad-Erstellung ---

                # Versuche, Zielordner zu erstellen (prüft intern auf Ausschluss)
                if (Create-FolderIfNotExists $destinationFolder) {
                    $finalDestinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName
                    Move-File $sourcePath $finalDestinationPath $effectiveDuplicateHandling
                } else {
                     Write-Log "WARNUNG: Zielordner '$destinationFolder' für '$fileName' konnte nicht erstellt werden oder ist ausgeschlossen. Überspringe."
                }
            } else {
                # Keine Regel hat gepasst -> Standard-Zielordner
                $defaultFolderName = if ($config.DefaultTargetFolder) { $config.DefaultTargetFolder } else { "Sonstiges" }
                $defaultTarget = Join-Path -Path $directory -ChildPath $defaultFolderName

                 if (Create-FolderIfNotExists $defaultTarget) {
                    $finalDestinationPath = Join-Path -Path $defaultTarget -ChildPath $fileName
                    Write-Log "Keine Regel matcht für Datei '$fileName'. Verschiebe zum Standardordner '$defaultTarget'. Effektives Duplicate Handling: '$effectiveDuplicateHandling'."
                    Move-File $sourcePath $finalDestinationPath $effectiveDuplicateHandling
                 } else {
                      Write-Log "WARNUNG: Standardzielordner '$defaultTarget' für '$fileName' konnte nicht erstellt werden oder ist ausgeschlossen. Überspringe."
                 }
            }
        } # Ende foreach ($file in $filesToProcess)
    } else {
         Write-Log "INFO: Keine gültigen 'FileRules' in der Konfiguration gefunden. Datei-Organisation wird übersprungen für '$directory'."
    }

    # --- Ordner-Management ---
    if ($null -ne $config.FolderRules -and $config.FolderRules -is [System.Array] -and $config.FolderRules.Count -gt 0) {
        Write-Log "--- Verarbeite Ordnerregeln rekursiv in '$directory' ---"
        # Hole alle Unterordner rekursiv, die NICHT ausgeschlossen sind
        $foldersToProcess = Get-ChildItem -Path $directory -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { -not (Should-ExcludePath $_.FullName -IsFolder) }

        foreach ($currentFolder in $foldersToProcess) {
            # Verhindere, dass die Haupt-Organisationsverzeichnisse selbst verschoben/umbenannt werden
            $isMainOrganizeDirectory = $false
            foreach ($mainDir in $dirsToOrganize) {
                 # Normalisiere beide Pfade vor dem Vergleich
                 try {
                     if ((Resolve-Path $mainDir).Path -eq (Resolve-Path $currentFolder.FullName).Path) {
                        $isMainOrganizeDirectory = $true
                        break
                     }
                 } catch {
                     Write-Log "WARNUNG: Konnte Pfade beim Vergleich für Hauptverzeichnisschutz nicht auflösen: '$mainDir' vs '$($currentFolder.FullName)'"
                 }
            }
            if ($isMainOrganizeDirectory) {
                 continue # Überspringe dieses Hauptverzeichnis für Ordnerregeln
            }

            # Wende die erste passende Ordnerregel an
            foreach ($folderRule in $config.FolderRules) {
                 # --- Regelanwendung initialisieren ---
                 $appliesRenameRule = !($true) # Workaround für PowerShell $false Erkennungsproblem
                 $appliesMoveRule = !($true)   # Workaround für PowerShell $false Erkennungsproblem
                 $ruleMatchedAndActionTaken = !($true) # Verhindert, dass mehrere Regeln auf denselben Ordner wirken

                 # --- Umbenennen-Regel prüfen ---
                 if (-not [string]::IsNullOrWhiteSpace($folderRule.RenamePattern) -and $currentFolder.Name -like $folderRule.RenamePattern) {
                     $renameCondition = $true
                     if ($null -ne $folderRule.RenameOlderThanDays -and $folderRule.RenameOlderThanDays -gt 0) {
                         if ($currentFolder.LastWriteTime -ge (Get-Date).AddDays(-$folderRule.RenameOlderThanDays)) {
                             $renameCondition = !($true) # Workaround
                             Write-Log "INFO: Ordner '$($currentFolder.FullName)' matcht RenamePattern '$($folderRule.RenamePattern)', aber RenameOlderThanDays nicht erfüllt."
                         }
                     }

                     if ($renameCondition) {
                         $newNameTemplate = if ($folderRule.NewNameTemplate) {$folderRule.NewNameTemplate} else {"{OriginalName}"}
                         $folderDate = $null
                         if ($currentFolder.LastWriteTime) { $folderDate = $currentFolder.LastWriteTime }
                         elseif ($currentFolder.CreationTime) { $folderDate = $currentFolder.CreationTime }

                         $newName = $newNameTemplate -replace '{OriginalName}', $currentFolder.Name
                         if ($folderDate) {
                             $newName = $newName -replace '\{JJJJ-MM\}', ($folderDate.ToString("yyyy-MM")) `
                                                -replace '\{JJJJ\}', ($folderDate.ToString("yyyy")) `
                                                -replace '\{MM\}', ($folderDate.ToString("MM")) `
                                                -replace '\{TT\}', ($folderDate.ToString("dd"))
                         } else {
                             $newName = $newName -replace '\{JJJJ-MM\}', "" -replace '\{JJJJ\}', "" -replace '\{MM\}', "" -replace '\{TT\}', ""
                         }

                         if ($currentFolder.Name -ne $newName) {
                             $newPath = Join-Path -Path $currentFolder.Parent.FullName -ChildPath $newName
                             Write-Log "Regel 'Umbenennen' (Beschreibung: '$($folderRule.Description)') matcht für Ordner '$($currentFolder.FullName)'."
                             Rename-Folder $currentFolder.FullName $newPath
                             $ruleMatchedAndActionTaken = $true
                             # WICHTIG: Nach erfolgreicher Umbenennung muss der Pfad aktualisiert werden, falls danach noch eine Move-Regel passen könnte!
                             # Dies wird hier vereinfacht ignoriert (break nach erster Aktion).
                         }
                     }
                 }

                 # --- Verschieben-Regel prüfen (nur wenn noch keine Aktion erfolgte) ---
                 if (-not $ruleMatchedAndActionTaken -and `
                     -not [string]::IsNullOrWhiteSpace($folderRule.MovePattern) -and `
                     $currentFolder.Name -like $folderRule.MovePattern -and `
                     -not [string]::IsNullOrWhiteSpace($folderRule.TargetFolder))
                 {
                     $moveCondition = $true
                     if ($null -ne $folderRule.MoveOlderThanDays -and $folderRule.MoveOlderThanDays -gt 0) {
                          if ($currentFolder.LastWriteTime -ge (Get-Date).AddDays(-$folderRule.MoveOlderThanDays)) {
                             $moveCondition = !($true) # Workaround
                             Write-Log "INFO: Ordner '$($currentFolder.FullName)' matcht MovePattern '$($folderRule.MovePattern)', aber MoveOlderThanDays nicht erfüllt."
                         }
                     }

                     if ($moveCondition) {
                         # Zielordner ist relativ zum *Haupt*-Organisationsverzeichnis ($directory)
                         $targetMoveBaseFolder = Join-Path -Path $directory -ChildPath $folderRule.TargetFolder
                         $newFolderPath = Join-Path -Path $targetMoveBaseFolder -ChildPath $currentFolder.Name

                         # Nur verschieben, wenn Pfade unterschiedlich sind und Ziel nicht existiert
                         if ($currentFolder.FullName.ToLower() -ne $newFolderPath.ToLower()) {
                            if (-not (Test-Path -Path $newFolderPath -PathType Container)) {
                                Write-Log "Regel 'Verschieben' (Beschreibung: '$($folderRule.Description)') matcht für Ordner '$($currentFolder.FullName)'."
                                Move-Folder $currentFolder.FullName $newFolderPath
                                $ruleMatchedAndActionTaken = $true
                            } else {
                                Write-Log "WARNUNG: Zielordner '$newFolderPath' für Verschiebung existiert bereits. Überspringe '$($currentFolder.FullName)'."
                            }
                         }
                     }
                 }

                 # Verlasse die Regelschleife für diesen Ordner, wenn eine Aktion durchgeführt wurde
                 if ($ruleMatchedAndActionTaken) { break }

            } # Ende foreach ($folderRule in $config.FolderRules)
        } # Ende foreach ($currentFolder in $foldersToProcess)
    } else {
        Write-Log "INFO: Keine gültigen 'FolderRules' in der Konfiguration gefunden. Ordner-Organisation wird übersprungen für '$directory'."
    }
     Write-Log "---- Organisation im Verzeichnis abgeschlossen: '$directory' ----"
}


# --- Hauptlogik ---

Write-Log "======================================================="
Write-Log "Starte die tiefgreifende Organisation..."
Write-Log "Skriptpfad: $($MyInvocation.MyCommand.Path)"
Write-Log "Konfigurationsdatei: $ConfigFile"
# if ($Preview) { Write-Log "*** VORSCHAU-MODUS AKTIV *** Es werden keine Änderungen durchgeführt." }
Write-Log "======================================================="

# Lade die Konfiguration aus der JSON-Datei
try {
    if (-not (Test-Path -Path $ConfigFile -PathType Leaf)) {
        throw "Konfigurationsdatei '$ConfigFile' nicht gefunden."
    }
    # Lese explizit als UTF8
    $jsonContent = Get-Content -Path $ConfigFile -Raw -Encoding UTF8 -ErrorAction Stop
    $config = $jsonContent | ConvertFrom-Json -ErrorAction Stop
    Write-Log "Konfigurationsdatei '$ConfigFile' erfolgreich geladen und geparst."

     # Globale Einstellungen aus der Konfiguration laden
     # GlobalExclusions
    if ($config.PSObject.Properties.Name -contains 'GlobalExclusions' -and $null -ne $config.GlobalExclusions) {
         if ($config.GlobalExclusions.PSObject.Properties.Name -contains 'FilePatterns' -and $config.GlobalExclusions.FilePatterns -is [System.Array]) {
             $globalFileExclusionPatterns = $config.GlobalExclusions.FilePatterns | Where-Object { -not ([string]::IsNullOrWhiteSpace($_)) }
             if ($globalFileExclusionPatterns.Count -gt 0) {
                Write-Log "Globale Datei-Ausschlussmuster geladen: $($globalFileExclusionPatterns -join ', ')"
             }
         }
         if ($config.GlobalExclusions.PSObject.Properties.Name -contains 'FolderPatterns' -and $config.GlobalExclusions.FolderPatterns -is [System.Array]) {
             $globalFolderExclusionPatterns = $config.GlobalExclusions.FolderPatterns | Where-Object { -not ([string]::IsNullOrWhiteSpace($_)) }
             if ($globalFolderExclusionPatterns.Count -gt 0) {
                Write-Log "Globale Ordner-Ausschlussmuster geladen: $($globalFolderExclusionPatterns -join ', ')"
             }
         }
     } else {
         Write-Log "INFO: Keine oder leere 'GlobalExclusions'-Sektion in der Konfiguration gefunden."
     }

    # GlobalDuplicateHandling (Korrigierte Validierung)
    if ($config.PSObject.Properties.Name -contains 'GlobalDuplicateHandling' -and $config.GlobalDuplicateHandling -is [string] -and -not [string]::IsNullOrWhiteSpace($config.GlobalDuplicateHandling)) {
        $allowedStrategies = @("skip", "renamewithtimestamp", "overwrite", "ask") # Kleinbuchstaben für Vergleich
        $providedStrategy = $config.GlobalDuplicateHandling
        $strategyLower = $providedStrategy.ToLower()

        if ($allowedStrategies -contains $strategyLower) {
            $globalDuplicateHandling = $strategyLower # Verwende validierten Kleinbuchstabenwert intern
            Write-Log "Globale Duplicate Handling Strategie geladen: '$providedStrategy'"
        } else {
            Write-Warning "Ungültige GlobalDuplicateHandling Strategie in Konfiguration: '$providedStrategy'. Verwende Standard: '$globalDuplicateHandling'."
            Write-Log "WARNUNG: Ungültige GlobalDuplicateHandling Strategie in Konfiguration: '$providedStrategy'. Verwende Standard: '$globalDuplicateHandling'."
            # $globalDuplicateHandling behält seinen Standardwert "skip"
        }
    } else {
        Write-Log "INFO: Keine gültige 'GlobalDuplicateHandling' Strategie in der Konfiguration gefunden. Verwende Standard: '$globalDuplicateHandling'"
    }

} catch {
    Write-Error "FATAL: Fehler beim Lesen oder Parsen der Konfigurationsdatei '$ConfigFile': $($_.Exception.Message). Das Skript wird beendet."
    Write-Log "FATAL: Fehler beim Lesen oder Parsen der Konfigurationsdatei '$ConfigFile': $($_.Exception.Message). Das Skript wird beendet."
    # Pop-Location # Falls Push-Location am Anfang verwendet wurde
    exit 1
}

# Bestimme die zu organisierenden Verzeichnisse
$dirsToOrganize = @()
$sourceDescription = ""

# Priorität: Parameter > Konfigurationsdatei > Automatische Erkennung
if ($Directories.Count -gt 0) {
     Write-Log "Verwende Verzeichnisse aus Skriptparametern."
     $sourceDescription = "Parameter"
     $dirsToOrganize = $Directories | ForEach-Object { $_ } # Kopiere Array, um Original nicht zu ändern
} elseif ($config.PSObject.Properties.Name -contains 'Directories' -and $config.Directories -is [System.Array] -and $config.Directories.Count -gt 0) {
     Write-Log "Verwende Verzeichnisse aus Konfigurationsdatei."
     $sourceDescription = "Konfiguration"
     $dirsToOrganize = $config.Directories | ForEach-Object { $_ }
} else {
    # Automatische Erkennung
    Write-Log "Keine Verzeichnisse explizit angegeben. Versuche automatische Erkennung nicht-riskanter Verzeichnisse..."
    $sourceDescription = "Automatisch erkannt"
    $systemDrive = ""
    try {
       $systemDrive = [System.IO.Path]::GetPathRoot((Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Name SystemRoot).SystemRoot)
        Write-Log "Systemlaufwerk ermittelt: '$systemDrive'"
    } catch { Write-Log "WARNUNG: Konnte Systemlaufwerk nicht ermitteln." }

    # Füge Benutzerprofil hinzu (wenn vorhanden und nicht ausgeschlossen)
    if (Test-Path -Path $env:USERPROFILE -PathType Container -ErrorAction SilentlyContinue) {
        if (-not (Should-ExcludePath $env:USERPROFILE -IsFolder)) {
             $dirsToOrganize += $env:USERPROFILE
             Write-Log "Automatisch hinzugefügt: Benutzerprofil '$env:USERPROFILE'"
        } else { Write-Log "INFO: Benutzerprofilpfad '$env:USERPROFILE' ist ausgeschlossen." }
    } else { Write-Log "WARNUNG: Benutzerprofilpfad '$env:USERPROFILE' nicht gefunden." }

    # Füge Wurzeln von Nicht-System-Festplatten hinzu (wenn vorhanden und nicht ausgeschlossen)
    Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Where-Object {$_.DriveType -eq "Fixed"} | ForEach-Object {
        $isSystemDriveRoot = $false
        if ($systemDrive -and ($_.Root -eq $systemDrive -or $_.Root -eq ($systemDrive -replace '\\$', ''))) {
            $isSystemDriveRoot = $true
        }
        if (-not $isSystemDriveRoot) {
            if (-not (Should-ExcludePath $_.Root -IsFolder)) {
                 $dirsToOrganize += $_.Root
                 Write-Log "Automatisch hinzugefügt: Laufwerksstamm '$($_.Root)'"
            } else { Write-Log "INFO: Laufwerksstamm '$($_.Root)' ist ausgeschlossen." }
        } elseif ($systemDrive) { Write-Log "INFO: Systemlaufwerksstamm '$($_.Root)' wird nicht automatisch hinzugefügt." }
    }
}

# Bereinige und filtere die Liste der zu organisierenden Verzeichnisse
$originalDirsCount = $dirsToOrganize.Count
$dirsToOrganize = $dirsToOrganize | Where-Object { -not ([string]::IsNullOrWhiteSpace($_)) } | Select-Object -Unique
$uniqueDirsCount = $dirsToOrganize.Count
$dirsToOrganize = $dirsToOrganize | Where-Object { Test-Path -Path $_ -PathType Container -ErrorAction SilentlyContinue }
$existingDirsCount = $dirsToOrganize.Count
$dirsToOrganize = $dirsToOrganize | Where-Object { -not (Should-ExcludePath $_ -IsFolder) }
$finalDirsCount = $dirsToOrganize.Count

Write-Log "Verzeichnisliste bereinigt (Quelle: $sourceDescription): $originalDirsCount -> $uniqueDirsCount (Unique) -> $existingDirsCount (Existiert) -> $finalDirsCount (Nicht ausgeschlossen)"

if ($dirsToOrganize.Count -eq 0) {
    Write-Error "FATAL: Keine gültigen, existierenden und nicht ausgeschlossenen Verzeichnisse zum Organisieren gefunden. Das Skript wird beendet."
    Write-Log "FATAL: Keine gültigen, existierenden und nicht ausgeschlossenen Verzeichnisse zum Organisieren gefunden. Das Skript wird beendet."
    # Pop-Location
    exit 1
}

Write-Log "ENDGÜLTIGE Liste der zu organisierenden Hauptverzeichnisse: $($dirsToOrganize -join ', ')"

# ---- Hauptverarbeitungsschleife ----
foreach ($dir in $dirsToOrganize) {
    # Die Prüfung auf Existenz und Ausschluss erfolgte bereits bei der Listenerstellung.
    # Direkter Aufruf der Organisationsfunktion.
    try {
        Organize-Directory -directory $dir -config $config -dirsToOrganize $dirsToOrganize
    } catch {
         Write-Error "FATAL: Unerwarteter Fehler bei der Organisation von '$dir': $($_.Exception.ToString()). Bearbeitung wird gestoppt."
         Write-Log "FATAL: Unerwarteter Fehler bei der Organisation von '$dir': $($_.Exception.ToString()). Bearbeitung wird gestoppt."
         # Pop-Location
         exit 1
    }
}

# ---- Aufräumen: Leere Ordner löschen (optional) ----
if ($DeleteEmptyFolders) {
    Write-Log "======================================================="
    Write-Log "Starte das Löschen leerer Ordner..."
    if ($Preview) { Write-Log "*** VORSCHAU-MODUS AKTIV *** Es werden keine Ordner gelöscht." }
    # Durchlaufe die *finalen* Hauptverzeichnisse
    foreach ($dir in $dirsToOrganize) {
        Write-Log "Prüfe auf leere Ordner in '$dir'..."
        try {
             Remove-EmptyFolders $dir # Rekursive Funktion
             Write-Log "Prüfung auf leere Ordner in '$dir' abgeschlossen."
        } catch {
             Write-Error "FEHLER beim Löschen leerer Ordner in '$dir': $($_.Exception.Message)"
             Write-Log "FEHLER beim Löschen leerer Ordner in '$dir': $($_.Exception.Message)"
        }
    }
    Write-Log "Löschen leerer Ordner abgeschlossen."
}

Write-Log "======================================================="
Write-Log "Tiefgreifende Organisation abgeschlossen."
# Write-Log "======================================================="

# Pop-Location # Falls Push-Location am Anfang verwendet wurde