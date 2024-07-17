
function Show-TicketFunction {
    param (
        [string]$username
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = "Ticket erstellen"
    $form.Size = New-Object Drawing.Size @(500, 400)
    $form.StartPosition = "CenterScreen"

    # Initialisieren der Ticketnummer aus einer Datei oder einem anderen Speicherort
	$ticketNumberFilePath = "\Daten\Netzmanagement\SUW\00 Doku\FWT\X Datenbank\Ticketsystem\tickno.txt"
    $ticketNumber = 1

    if (Test-Path $ticketNumberFilePath) {
        $ticketNumber = Get-Content $ticketNumberFilePath
        $ticketNumber = [int]$ticketNumber
    }

    $labelType = New-Object Windows.Forms.Label
    $labelType.Text = "Tickettyp:"
    $labelType.Location = New-Object Drawing.Point @(10, 20)
	$labelType.Forecolor = [System.Drawing.Color]::White
    $form.Controls.Add($labelType)
	
	$form.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0) 
	$form.BackgroundImage = $bitmap
	$form.BackgroundImageLayout = "Stretch"  # Oder "Zoom", je nach Bedarf

    $comboBoxType = New-Object Windows.Forms.ComboBox
    $comboBoxType.AutoCompleteMode = "SuggestAppend"
    $comboBoxType.AutoCompleteSource = "CustomSource"
    $comboBoxType.Location = New-Object Drawing.Point @(120, 20)
    $comboBoxType.Size = New-Object Drawing.Size @(250, 20)
    $comboBoxType.Items.AddRange(@("Solid - Bug Report", "Solid - Verbesserung", "EPLAN - Plananfrage", "EPLAN - Korrekturen"))
    $form.Controls.Add($comboBoxType)

    $labelUser = New-Object Windows.Forms.Label
    $labelUser.Text = "Ersteller: "
    $labelUser.Location = New-Object Drawing.Point @(10, 60)
	$labelUser.ForeColor = [System.Drawing.Color]::White  # Ändern Sie hier die Textfarbe des Buttons
	$labelUser.BackColor = [System.Drawing.Color]::FromArgb(2,7,26)
    $form.Controls.Add($labelUser)
	    $labelUserUser = New-Object Windows.Forms.Label
		$labelUserUser.Text = "$username"
		$labelUserUser.Location = New-Object Drawing.Point @(120, 60)
		$labelUserUser.ForeColor = [System.Drawing.Color]::Orange  # Ändern Sie hier die Textfarbe des Buttons
		$labelUserUser.BackColor = $transparent1
		$form.Controls.Add($labelUserUser)

    $textBoxDescription = New-Object Windows.Forms.TextBox
	$textBoxDescription.Multiline = $true
	$textBoxDescription.ScrollBars = 'Both'
	$textBoxDescription.Location = New-Object Drawing.Point @(10, 90)
	$textBoxDescription.Size = New-Object Drawing.Size @(400, 200)
	$textBoxDescription.Text = "Beschreibung eingeben..."  # Platzhaltertext
	$textBoxDescription.Add_Enter({
		if ($textBoxDescription.Text -eq "Beschreibung eingeben...") {
			$textBoxDescription.Text = ""
			$textBoxDescription.ForeColor = [System.Drawing.Color]::Black
		}
	})
	$textBoxDescription.Add_Leave({
		if ($textBoxDescription.Text -eq "") {
			$textBoxDescription.Text = "Beschreibung eingeben..."
			$textBoxDescription.ForeColor = [System.Drawing.Color]::Gray
		}
	})
	$form.Controls.Add($textBoxDescription)

    $buttonSend = New-Object Windows.Forms.Button
    $buttonSend.Text = "Senden"
    $buttonSend.Size = New-Object Drawing.Size @(80, 30)
    $buttonSend.Location = New-Object Drawing.Point @(10, 300)
	$buttonSend.BackColor = [System.Drawing.Color]::White  # Ändern Sie hier die Textfarbe des Buttons
	$buttonSend.ForeColor = [System.Drawing.Color]::Black
    $buttonSend.Add_Click({
        $selectedType = $comboBoxType.SelectedItem
        $description = $textBoxDescription.Text

        # Ticketnummer erhöhen und in die Datei schreiben
        $ticketNumber++
        $ticketNumber | Set-Content $ticketNumberFilePath

        # Betreff mit der aktualisierten Ticketnummer erstellen
        $subject = "Ticket #$ticketNumber - SOLID"

        # Outlook-Anwendung erstellen
        $outlook = New-Object -ComObject Outlook.Application

        # Neue Mail erstellen
        $mail = $outlook.CreateItem(0)
        $mail.Subject = $subject
        $mail.Body = "Ein neues Ticket wurde erstellt.`r`nTickettyp: $selectedType`r`nBenutzer: $username`r`nBeschreibung: $description"

        # Empfänger hinzufügen
        $mail.Recipients.Add("alper.frink@sw-netz.de")

        # Mail senden
        $mail.Send()

        # Ändern Sie die Farbe des Buttons nach erfolgreichem Senden
        $buttonSend.BackColor = [System.Drawing.Color]::Green

        # Entfernen des ursprünglichen Senden-Buttons und Abbrechen-Buttons
        $buttonSend.Dispose()
        $buttonCancel.Dispose()

        # Hinzufügen des "Fertig"-Buttons
        $buttonFinish = New-Object Windows.Forms.Button
        $buttonFinish.Text = "Fertig"
        $buttonFinish.Size = New-Object Drawing.Size @(80, 30)
        $buttonFinish.Location = New-Object Drawing.Point @(10, 300)
		$buttonFinish.BackColor = [System.Drawing.Color]::FromArgb(0, 245, 0)
		$buttonFinish.ForeColor = [System.Drawing.Color]::White
        $buttonFinish.Add_Click({
            $form.Close()
        })
        $form.Controls.Add($buttonFinish)
    })
    $form.Controls.Add($buttonSend)

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Text = "Abbrechen"
    $buttonCancel.Size = New-Object Drawing.Size @(80, 30)
    $buttonCancel.Location = New-Object Drawing.Point @(100, 300)
	$buttonCancel.ForeColor = [System.Drawing.Color]::Black  # Ändern Sie hier die Textfarbe des Buttons
	$buttonCancel.BackColor = [System.Drawing.Color]::White  # Ändern Sie hier die Textfarbe des Buttons
    $buttonCancel.Add_Click({
        $form.Close()
    })
    $form.Controls.Add($buttonCancel)

    $form.Topmost = $true
    $form.ShowDialog()
}