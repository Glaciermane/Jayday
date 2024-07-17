
# Funktion für das Popup-Fenster mit Eingabezeile, Dropdown-Menü und Checkbox

function Show-StationSearchPopup {
    param (
        [string]$Title = "Station Suchen",
        [string]$Message = "Station:"
    )

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(400, 150)  # Ändern Sie hier die Größe des Fensters
    $form.StartPosition = "CenterScreen"

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Location = New-Object Drawing.Point @(10, 20)
    $form.Controls.Add($label)

    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point @(10, 50)
    $textBox.Size = New-Object Drawing.Size @(250, 20)
    $textBox.AutoCompleteMode = "SuggestAppend"
    $textBox.AutoCompleteSource = "CustomSource"

    # Lese die Namen der Unterordner aus dem angegebenen Pfad und füge sie zur Autovervollständigungsquelle hinzu
    $folderPath = "\Daten\Netzmanagement\SUW\00 Doku"  # Ändern Sie dies entsprechend Ihrem tatsächlichen Pfad
    $subfolders = Get-ChildItem -Path $folderPath -Directory | Select-Object -ExpandProperty Name

    $autoComplete = New-Object Windows.Forms.AutoCompleteStringCollection
    $autoComplete.AddRange($subfolders)
    $textBox.AutoCompleteCustomSource = $autoComplete

    $form.Controls.Add($textBox)

    $buttonSearch = New-Object Windows.Forms.Button
    $buttonSearch.Text = "Suchen"
    $buttonSearch.Size = New-Object Drawing.Size @(80, 30)
    $buttonSearch.Location = New-Object Drawing.Point @(10, 80)
    $buttonSearch.Add_Click({
        $enteredText = $textBox.Text

        # Fügen Sie hier den Code hinzu, um die Informationen zur ausgewählten Station abzurufen und anzuzeigen
        if ($enteredText -ne "") {
            $form.Hide()
            Show-StationInfo -stationName $enteredText
            $form.Show()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen gültigen Stationsnamen ein.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
		$form.Close()
        
    })
    $form.Controls.Add($buttonSearch)

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Text = "Abbrechen"
    $buttonCancel.Size = New-Object Drawing.Size @(80, 30)
    $buttonCancel.Location = New-Object Drawing.Point @(100, 80)
    $buttonCancel.Add_Click({
        # Fügen Sie hier den Code hinzu, der auf den Abbrechen-Button klicken soll
        $form.Close()
    })
    $form.Controls.Add($buttonCancel)

    $form.Topmost = $true
    $form.ShowDialog()
}

function Show-StationInfo {
    param (
        [string]$stationName
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = "Stationsinformationen - $stationName"
    $form.Size = New-Object Drawing.Size @(500, 400)
    $form.StartPosition = "CenterScreen"

    $labelName = New-Object Windows.Forms.Label
    $labelName.Text = "Name: $stationName"
    $labelName.Location = New-Object Drawing.Point @(10, 20)
    $labelName.Size = New-Object Drawing.Size @(400, 20)
    $form.Controls.Add($labelName)

    $labelSchaltanlagen = New-Object Windows.Forms.Label
    $schaltanlagenFile = Join-Path $stationName "Schaltanlage.txt"
    $labelSchaltanlagen.Text = "Schaltanlage: $(Get-Content $schaltanlagenFile -ErrorAction SilentlyContinue)"
    $labelSchaltanlagen.Location = New-Object Drawing.Point @(10, 60)
    $labelSchaltanlagen.Size = New-Object Drawing.Size @(400, 20)
    $form.Controls.Add($labelSchaltanlagen)

    $labelDatenpunkte = New-Object Windows.Forms.Label
    $datenpunkteFile = Join-Path $stationName "Datenpunktliste.txt"
    $labelDatenpunkte.Text = "Datenpunktliste: $(Get-Content $datenpunkteFile -ErrorAction SilentlyContinue)"
    $labelDatenpunkte.Location = New-Object Drawing.Point @(10, 90)
    $labelDatenpunkte.Size = New-Object Drawing.Size @(400, 20)
    $form.Controls.Add($labelDatenpunkte)

    $labelAusbaustand = New-Object Windows.Forms.Label
    $ausbaustandFile = Join-Path $stationName "14 Fernwirktechnik\Anlagendokumentation\Datenpunktliste.xls"
    $labelAusbaustand.Text = "Aktueller Ausbaustand: $(Get-Content $ausbaustandFile -ErrorAction SilentlyContinue)"
    $labelAusbaustand.Location = New-Object Drawing.Point @(10, 120)
    $labelAusbaustand.Size = New-Object Drawing.Size @(400, 20)
    $form.Controls.Add($labelAusbaustand)

    $buttonAufruf = New-Object Windows.Forms.Button
    $buttonAufruf.Text = "Aufruf"
    $buttonAufruf.Size = New-Object Drawing.Size @(80, 30)
    $buttonAufruf.Location = New-Object Drawing.Point @(10, 300)
    $buttonAufruf.Add_Click({
        # Hier können Sie den Code für den Aufruf der Schaltanlagenpläne oder anderer Informationen hinzufügen.
        Write-Host "Aufruf durchgeführt"
    })
    $form.Controls.Add($buttonAufruf)

    $buttonSchliessen = New-Object Windows.Forms.Button
    $buttonSchliessen.Text = "Schließen"
    $buttonSchliessen.Size = New-Object Drawing.Size @(80, 30)
    $buttonSchliessen.Location = New-Object Drawing.Point @(100, 300)
    $buttonSchliessen.Add_Click({
        $form.Close()
    })

    $form.Controls.Add($buttonSchliessen)
    Update-FormEncoding
    $form.Topmost = $true
    $form.ShowDialog()
}