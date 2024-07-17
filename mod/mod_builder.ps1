
function Show-CustomPopup {
	
    param (
        [string]$Title = "SOLID",
        [string]$Message = "logged in as:",
		[string]$username
    )
	
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()
	

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(1200, 800)
    $form.StartPosition = "CenterScreen"
	$form.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0)  

	# Hintergrundbild einfügen
    $backgroundImagePath = "SOLID.jpeg"  # Ändern Sie dies entsprechend Ihrem tatsächlichen Pfad
    $backgroundImage = [System.Drawing.Image]::FromFile($backgroundImagePath)
	
	$transparency = 0.2  # Hier kannst du den Transparenzwert anpassen (0.0 = vollständig transparent, 1.0 = undurchsichtig)
	$transparent1 = 0.0
	
# Bitmap mit der gewünschten Transparenz erstellen
$bitmap = New-Object System.Drawing.Bitmap $backgroundImage.Width, $backgroundImage.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$colorMatrix = New-Object Drawing.Imaging.ColorMatrix
$colorMatrix.Matrix33 = $transparency
$imageAttributes = New-Object Drawing.Imaging.ImageAttributes
$imageAttributes.SetColorMatrix($colorMatrix, [Drawing.Imaging.ColorMatrixFlag]::SkipGrays)

# Bild auf das Bitmap zeichnen
$graphics.DrawImage($backgroundImage, [System.Drawing.Rectangle]::new(0, 0, $bitmap.Width, $bitmap.Height), 0, 0, $bitmap.Width, $bitmap.Height, [System.Drawing.GraphicsUnit]::Pixel, $imageAttributes)
$graphics.Dispose()

# Hintergrundbild auf dem Formular setzen
$form.BackgroundImage = $bitmap
$form.BackgroundImageLayout = "Stretch"  # Oder "Zoom", je nach Bedarf
	
	# Festlegen der Größe des Formulars und Deaktivieren der Größenänderung
    $form.FormBorderStyle = "Fixed3D"  # Oder "FixedDialog" je nach Bedarf
    $form.MaximizeBox = $false
	
    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Location = New-Object Drawing.Point @(10, 20)
	$label.Size = New-Object Drawing.Size @(95, 20)
	$label.ForeColor = [System.Drawing.Color]::White
	$label.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 70)
	$label.Font = New-Object Drawing.Font("Arial", 11)
    $form.Controls.Add($label)
	
    $labelUser = New-Object Windows.Forms.Label
    $labelUser.Text = $username
    $labelUser.Location = New-Object Drawing.Point @(105, 17)
    $labelUser.Size = New-Object Drawing.Size @(95, 20)
    $labelUser.ForeColor = [System.Drawing.Color]::FromArgb(62, 219, 0)
	$labelUser.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 70)
    $labelUser.Font = New-Object Drawing.Font("Arial", 15, [System.Drawing.FontStyle]::Bold)
	
    $form.Controls.Add($labelUser)
	

    # Anzeige des Online-Status unten links
    $labelOnline = New-Object Windows.Forms.Label
    $labelOnline.Text = "Online Users:"
    $labelOnline.Location = New-Object Drawing.Point @(10, 80)
    $labelOnline.Size = New-Object Drawing.Size @(150, 20)
    $labelOnline.ForeColor = [System.Drawing.Color]::White
    $labelOnline.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 70)
    $labelOnline.Font = New-Object Drawing.Font("Arial", 11)
    $form.Controls.Add($labelOnline)

	$labelStat = New-Object Windows.Forms.Label
	$labelStat.Location = New-Object Drawing.Point @(10, 90)
	$labelStat.Size = New-Object Drawing.Size @(100, 100)
	$labelStat.Name = "labelStat"
	#$labelStat.Text = GET-Content stat.txt -Raw
	$labelstat.ForeColor = [System.Drawing.Color]::Orange
	$labelStat.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 70)
	$labelstat.Font = New-Object Drawing.Font("Arial", 11)
	$form.Controls.Add($labelStat)
	
	# Funktion zum Aktualisieren des Label-Inhalts
function Update-LabelStat {
    $labelStat.Text = Get-Content -Path "stat.txt" -Raw
}
	
	# Timer erstellen
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 10000  # Interval in Millisekunden (10 Sekunden)
$timer.Add_Tick({
    Update-LabelStat
})
# Timer starten
$timer.Start()

# Überwachung der Datei auf Änderungen
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = (Get-Item "stat.txt").DirectoryName
$watcher.Filter = (Get-Item "stat.txt").Name
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
$watcher.IncludeSubdirectories = $false

# Ereignis für Änderungen hinzufügen
    $onChangeAction = {
        # Hier den Code ausführen, wenn die Datei geändert wurde
        Update-LabelStat
    }
Register-ObjectEvent -InputObject $watcher -EventName Changed -Action $onChangeAction | Out-Null
# Starten der Überwachung
$watcher.EnableRaisingEvents = $true
	
    $buttonWidth = 230
    $buttonHeight = 70

    $button1 = New-Object Windows.Forms.Button
    $button1.Text = "Station Suchen"
    $button1.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button1.Location = New-Object Drawing.Point @(490, 200)
    $button1.BackColor = [System.Drawing.Color]::Green  # Ändern Sie hier die Farbe des Buttons
    $button1.ForeColor = [System.Drawing.Color]::White  # Ändern Sie hier die Textfarbe des Buttons
    $button1.Add_Click({ mod_R_C_create })
	$button1.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	# MouseEnter Ereignis für Hover-Effekt hinzufügen
    $button1.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::LightGreen
    })
    # MouseLeave Ereignis für den ursprünglichen Zustand hinzufügen
    $button1.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::Green
    })
    $form.Controls.Add($button1)
	
    $button2 = New-Object Windows.Forms.Button
    $button2.Text = "Station Anlegen"
    $button2.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button2.Location = New-Object Drawing.Point @(490, 400)
    $button2.BackColor = [System.Drawing.Color]::Blue  # Ändern Sie hier die Farbe des Buttons
    $button2.ForeColor = [System.Drawing.Color]::White  # Ändern Sie hier die Textfarbe des Buttons
    $button2.Add_Click({ Show-StationCreationPopup -username $username })
	$button2.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	# MouseEnter Ereignis für Hover-Effekt hinzufügen
    $button2.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::LightBlue
    })
    # MouseLeave Ereignis für den ursprünglichen Zustand hinzufügen
    $button2.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::Blue
    })
    $form.Controls.Add($button2)

    $button3 = New-Object Windows.Forms.Button
    $button3.Text = "Dokumentation"
    $button3.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button3.Location = New-Object Drawing.Point @(110, 300)
    $button3.BackColor = [System.Drawing.Color]::FromArgb(235, 117, 0)  # Ändern Sie hier die Farbe des Buttons
    $button3.ForeColor = [System.Drawing.Color]::White  # Ändern Sie hier die Textfarbe des Buttons
    $button3.Add_Click({ DokumentationFunction })
	$button3.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	# MouseEnter Ereignis für Hover-Effekt hinzufügen
    $button3.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(255, 177, 77)
    })
    # MouseLeave Ereignis für den ursprünglichen Zustand hinzufügen
    $button3.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(235, 117, 0)
    })
    $form.Controls.Add($button3)

    $button4 = New-Object Windows.Forms.Button
    $button4.Text = "Ticket"
    $button4.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button4.Location = New-Object Drawing.Point @(870, 300)
    $button4.Add_Click({ Show-TicketFunction -username $username })
	$button4.BackColor = [System.Drawing.Color]::FromArgb(0, 153, 156)
	$button4.ForeColor = [System.Drawing.Color]::White
	$button4.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	# MouseEnter Ereignis für Hover-Effekt hinzufügen
    $button4.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(0, 173, 176)
    })
    # MouseLeave Ereignis für den ursprünglichen Zustand hinzufügen
    $button4.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(0, 153, 156)
    })
    $form.Controls.Add($button4)

    $button5 = New-Object Windows.Forms.Button
    $button5.Text = "logs"
    $button5.Size = New-Object Drawing.Size @(150, 40)
    $button5.Location = New-Object Drawing.Point @(530, 650)
	$button5.BackColor = [System.Drawing.Color]::FromArgb(171, 171, 171)
    $button5.Add_Click({ Show-LogsFunction })
	$button5.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	# MouseEnter Ereignis für Hover-Effekt hinzufügen
    $button5.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(191, 191, 191)
    })
    # MouseLeave Ereignis für den ursprünglichen Zustand hinzufügen
    $button5.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(171, 171, 171)
    })
    $form.Controls.Add($button5)

    $form.add_Closing({
        # Hier den Code ausführen, bevor das Fenster geschlossen wird
        # Zum Beispiel: Save-DataFunction
        Save-DataFunction
		    # Timer stoppen
    $timer.Stop()
    $timer.Dispose()
		
		        # Stoppen der Überwachung
        $watcher.EnableRaisingEvents = $false

        # Das Fenster schließen
        $_.Cancel = $false  # Setzen Sie auf $true, wenn das Schließen verhindert werden soll
    })
	Update-LabelStat  # Initialen Online-Status setzen
	Update-FormEncoding

    $form.Topmost = $false
    $form.ShowDialog()
}
function Save-DataFunction {
    $statFilePath = "stat.txt"
    
    # Überprüfen, ob die Datei existiert
    if (Test-Path $statFilePath) {
        # Benutzername zum Vergleichen
        $usernameToCompare = $username  # Hier kann eine andere Variable verwendet werden, falls erforderlich

        # Inhalt der Datei "stat.txt" lesen
        $content = Get-Content $statFilePath -Raw

        # Muster für den Abschnitt erstellen
        $pattern = "(?ms)^\s*$usernameToCompare.*?^\s*(?:\n|\Z)"

        # Überprüfen, ob der Abschnitt vorhanden ist
        if ($content -match $pattern) {
            # Den Abschnitt aus dem Inhalt entfernen
            $content = $content -replace $pattern

            # Den aktualisierten Inhalt zurück in die Datei schreiben
            $content | Set-Content $statFilePath -Force
            Write-Host "Abschnitt für Benutzer '$usernameToCompare' wurde aus 'stat.txt' entfernt."
        } else {
            Write-Host "Abschnitt für Benutzer '$usernameToCompare' wurde nicht gefunden in 'stat.txt'."
        }
    } else {
        Write-Host "Datei 'stat.txt' wurde nicht gefunden."
    }
}

# Funktionen für die verschiedenen Aktionen
