
# Funktion für das Popup-Fenster zum Anlegen einer Station
function Show-StationCreationPopup {
    param (
        [string]$Title = "Station Anlegen",
        [string]$Message = "Neue Station anlegen:"
    )

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(600, 300)  # Größe des Fensters erhöht
    $form.StartPosition = "CenterScreen"

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Location = New-Object Drawing.Point @(10, 20)
    $form.Controls.Add($label)

    $labelStationNumber = New-Object Windows.Forms.Label
    $labelStationNumber.Text = "Stationsnummer:"
    $labelStationNumber.Location = New-Object Drawing.Point @(10, 50)
    $form.Controls.Add($labelStationNumber)

    $textBoxStationNumber = New-Object Windows.Forms.TextBox
    $textBoxStationNumber.Location = New-Object Drawing.Point @(150, 50)
    $textBoxStationNumber.Size = New-Object Drawing.Size @(250, 20)
    $textBoxStationNumber.Text = "Geben Sie die Stationsnummer ein..."
    $textBoxStationNumber.Add_Enter({
        if ($textBoxStationNumber.Text -eq "Geben Sie die Stationsnummer ein...") {
            $textBoxStationNumber.Text = ""
            $textBoxStationNumber.ForeColor = [System.Drawing.Color]::Black
        }
    })
    $textBoxStationNumber.Add_Leave({
        if ($textBoxStationNumber.Text -eq "") {
            $textBoxStationNumber.Text = "Geben Sie die Stationsnummer ein..."
            $textBoxStationNumber.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    $form.Controls.Add($textBoxStationNumber)

    $labelStationName = New-Object Windows.Forms.Label
    $labelStationName.Text = "Stationsname:"
    $labelStationName.Location = New-Object Drawing.Point @(10, 80)
    $form.Controls.Add($labelStationName)

    $textBoxStationName = New-Object Windows.Forms.TextBox
    $textBoxStationName.Location = New-Object Drawing.Point @(150, 80)
    $textBoxStationName.Size = New-Object Drawing.Size @(250, 20)
    $textBoxStationName.Text = "Geben Sie den Stationsnamen ein..."
    $textBoxStationName.Add_Enter({
        if ($textBoxStationName.Text -eq "Geben Sie den Stationsnamen ein...") {
            $textBoxStationName.Text = ""
            $textBoxStationName.ForeColor = [System.Drawing.Color]::Black
        }
    })
    $textBoxStationName.Add_Leave({
        if ($textBoxStationName.Text -eq "") {
            $textBoxStationName.Text = "Geben Sie den Stationsnamen ein..."
            $textBoxStationName.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    $form.Controls.Add($textBoxStationName)

    $labelOptions = New-Object Windows.Forms.Label
    $labelOptions.Text = "Typ:"
    $labelOptions.Location = New-Object Drawing.Point @(10, 110)
    $form.Controls.Add($labelOptions)

    $comboBoxOptions = New-Object Windows.Forms.ComboBox
    $comboBoxOptions.AutoCompleteMode = "SuggestAppend"
    $comboBoxOptions.AutoCompleteSource = "CustomSource"
    $comboBoxOptions.Location = New-Object Drawing.Point @(150, 110)
    $comboBoxOptions.Size = New-Object Drawing.Size @(250, 20)
    $comboBoxOptions.Items.AddRange(@("Typ 1 RRT", "Typ 2 RRL", "Typ 3 RRT + Absorber", "Typ 4 RRL + Absorber", "Typ 5 RRTT", "Typ 6 RRRT"))
    $form.Controls.Add($comboBoxOptions)
	
    $progressBar = New-Object Windows.Forms.ProgressBar
    $progressBar.Style = "Continuous"  # Animationsstil
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0  # Anfangswert
	$progressBar.ForeColor = "Green"  # Farbe des Fortschritts
    $progressBar.Location = New-Object Drawing.Point @(10, 190)
    $progressBar.Size = New-Object Drawing.Size @(560, 20)
    $form.Controls.Add($progressBar)
	
	$loaderBar = New-Object Windows.Forms.ProgressBar
    $loaderBar.Style = "Continuous"  # Animationsstil
    $LoaderBar.Minimum = 0
    $loaderBar.Maximum = 100
    $loaderBar.Value = 0  # Anfangswert
	$loaderBar.ForeColor = "Green"  # Farbe des Fortschritts
    $loaderBar.Location = New-Object Drawing.Point @(10, 190)
    $loaderBar.Size = New-Object Drawing.Size @(560, 20)
    $form.Controls.Add($loaderBar)

    $buttonCreate = New-Object Windows.Forms.Button
    $buttonCreate.Text = "Anlegen"
    $buttonCreate.Size = New-Object Drawing.Size @(80, 30)
    $buttonCreate.Location = New-Object Drawing.Point @(10, 140)
    $buttonCreate.Add_Click({
        $stationNumber = $textBoxStationNumber.Text
        $stationName = $textBoxStationName.Text
        $selectedOption = $comboBoxOptions.SelectedItem

        try {
            # Überprüfen, ob alle Felder ausgefüllt sind
            if ($stationNumber -eq "" -or $stationName -eq "" -or $selectedOption -eq $null) {
                throw "Bitte füllen Sie alle Felder aus."
            }

            # Überprüfen, ob die Stationsnummer aus Zahlen besteht und genau 4-stellig ist
            if ($stationNumber -notmatch '^\d{4}$') {
                throw "Ungültige Stationsnummer. Geben Sie eine 4-stellige Zahl ein."
            }

            # Zielverzeichnis erstellen und prüfen, ob es bereits existiert
            $destinationDirectory = "U:\Daten\Netzmanagement\SUW\00 Doku"
            $newFolderName = "$stationNumber $stationName"
            $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $newFolderName

            if (-not (Test-Path -Path $destinationPath -PathType Container)) {
                New-Item -Path $destinationPath -ItemType Directory | Out-Null
            } else {
                $newFolderName += " - Solid"
                $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $newFolderName
                New-Item -Path $destinationPath -ItemType Directory | Out-Null
            }

            # Kopieren des Ordners basierend auf der ausgewählten Option
			$ue = [char]0xDC
			$uedit = "bergeordnet"
            $sourceDirectory = "U:\Daten\Netzmanagement\SUW\00 Doku\0000 $ue$uedit\50 Netzstationen\$selectedOption"

            if (-not (Test-Path -Path $sourceDirectory -PathType Container)) {
                throw "Quellverzeichnis '$sourceDirectory' existiert nicht."
            }
			
			
            # Kopieren der Dateien
            Copy-Item -Path $sourceDirectory\* -Destination $destinationPath -Recurse -ErrorAction Stop
			
			$files = Get-ChildItem -Path $sourceDirectory -File
			$totalFiles = $files.Count
			$progress = 0
			
            foreach ($file in $files) {
                Copy-Item -Path $file.FullName -Destination $destinationPath -Force
                $progress++
                $percentComplete = [math]::Round(($progress / $totalFiles) * 100)
                Update-ProgressBar -Value $percentComplete
                $form.Refresh()
            }
			
            # Log-Eintrag hinzufügen
            $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Der Benutzer $env:USERNAME hat die Station $stationName angelegt."
            Add-Content -Path "log.txt" -Value $logEntry

            # Den Log-Eintrag in der Konsole in roter Schrift anzeigen
            Write-Host -ForegroundColor Red $logEntry
			

			# Durchsuchen und Ersetzen der Dateinamen im neu erstellten Ordner
			$fwtfolder = "\Fernsteuerung"
            $files = Get-ChildItem -Path $destinationPath$fwtfolder -File
            foreach ($file in $files) {
                $newFileName = $file.Name -replace '^X{1,4}', $stationNumber
                $newFilePath = Join-Path -Path $destinationPath -ChildPath $newFileName
                Rename-Item -Path $file.FullName -NewName $newFileName -Force
				
            }
			# Erfolgsmeldung anzeigen
            [System.Windows.Forms.MessageBox]::Show("Die neue Station wurde erfolgreich angelegt.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			
        } catch {
            Write-Host "Fehler beim Anlegen der Station: $_"
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Anlegen der Station. $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        } 
    })
    $form.Controls.Add($buttonCreate)

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Text = "Abbrechen"
    $buttonCancel.Size = New-Object Drawing.Size @(80, 30)
    $buttonCancel.Location = New-Object Drawing.Point @(100, 140)
    $buttonCancel.Add_Click({
        # Fügen Sie hier den Code hinzu, der auf den Abbrechen-Button klicken soll
        $form.Close()
    })
    $form.Controls.Add($buttonCancel)

    $form.Topmost = $true
    $form.ShowDialog()
}

# Beispielaufruf für das Anlegen einer Station
$buttonStationAnlegen.Add_Click({
    Show-StationCreationPopup
})