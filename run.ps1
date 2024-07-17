
Install-Module -Name ImportExcel -Force -AllowClobber
$userFilePath = "users.json"
$statFilePath = "stat.json"
. "$PSScriptRoot\mod_R_C_create.ps1"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Windows.Forms.Application]::EnableVisualStyles()

# Funktion zur Überprüfung der Benutzeranmeldeinformationen
function Validate-User {
    param (
        [string]$username,
        [string]$password
    )

	
    # Lade Benutzer aus der JSON-Datei
    $users = Get-Content $userFilePath | ConvertFrom-Json

    # Überprüfen, ob der Benutzer existiert und das Passwort korrekt ist
    $user = $users | Where-Object { $_.Username -eq $username -and $_.Password -eq $password }

        
        return $user
}




	

# ==================================================================== Panel ================================================================================
# ==================================================================== Panel ================================================================================
# ==================================================================== Panel ================================================================================

function Show-CustomPopup {
	
    param (
        [string]$Title = "SOLID",
        [string]$Message = "logged in as:",
		[string]$username
    )
	
	$tag = ""
    if ($username -eq "4860") {
        $tag = "AF"
    } elseif ($username -eq "5206") {
        $tag = "MB"
    } elseif ($username -eq "9583") {
        $tag = "MD"
	} elseif ($username -eq "5859") {
        $tag = "SK"
	} elseif ($username -eq "8338") {
        $tag = "SR"
	} elseif ($username -eq "7069") {
        $tag = "AK"
	} 
	
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()
	
	
	
    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(1200, 800)
    $form.StartPosition = "CenterScreen"
	$form.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0)  

	
    $backgroundImagePath = "SOLID.jpeg"  
    $backgroundImage = [System.Drawing.Image]::FromFile($backgroundImagePath)
	
	$transparency = 0.2  
	$transparent1 = 0.0
	

$bitmap = New-Object System.Drawing.Bitmap $backgroundImage.Width, $backgroundImage.Height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$colorMatrix = New-Object Drawing.Imaging.ColorMatrix
$colorMatrix.Matrix33 = $transparency
$imageAttributes = New-Object Drawing.Imaging.ImageAttributes
$imageAttributes.SetColorMatrix($colorMatrix, [Drawing.Imaging.ColorMatrixFlag]::SkipGrays)


$graphics.DrawImage($backgroundImage, [System.Drawing.Rectangle]::new(0, 0, $bitmap.Width, $bitmap.Height), 0, 0, $bitmap.Width, $bitmap.Height, [System.Drawing.GraphicsUnit]::Pixel, $imageAttributes)
$graphics.Dispose()


$form.BackgroundImage = $bitmap
$form.BackgroundImageLayout = "Stretch" 
	
	
    $form.FormBorderStyle = "Fixed3D"  
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
    $labelUser.Size = New-Object Drawing.Size @(100, 20)
    $labelUser.ForeColor = [System.Drawing.Color]::FromArgb(62, 219, 0)
	$labelUser.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 70)
    $labelUser.Font = New-Object Drawing.Font("Arial", 15, [System.Drawing.FontStyle]::Bold)
	
    $form.Controls.Add($labelUser)
	

    
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
	
	$labelstat.ForeColor = [System.Drawing.Color]::Orange
	$labelStat.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 70)
	$labelstat.Font = New-Object Drawing.Font("Arial", 11)
	$form.Controls.Add($labelStat)
	
	# Funktion zum Aktualisieren des Label-Inhalts
function Update-LabelStat {
    $labelStat.Text = Get-Content -Path "qCON.go" -Raw
}
	
	
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 3000  # Interval in Millisekunden (3 Sekunden)
$timer.Add_Tick({
    Update-LabelStat
})

$timer.Start()

# Überwachung der Datei auf Änderungen
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = (Get-Item "qCON.go").DirectoryName
$watcher.Filter = (Get-Item "qCON.go").Name
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
$watcher.IncludeSubdirectories = $false


    $onChangeAction = {
        
        Update-LabelStat
    }
Register-ObjectEvent -InputObject $watcher -EventName Changed -Action $onChangeAction | Out-Null

$watcher.EnableRaisingEvents = $true
	
    $buttonWidth = 230
    $buttonHeight = 70

    $button1 = New-Object Windows.Forms.Button
    $button1.Text = "Station Suchen"
    $button1.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button1.Location = New-Object Drawing.Point @(490, 200)
    $button1.BackColor = [System.Drawing.Color]::Green  
    $button1.ForeColor = [System.Drawing.Color]::White   
    $button1.Add_Click({ Show-StationSearchPopup })
	$button1.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	
    $button1.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::LightGreen
    })
    
    $button1.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::Green
    })
    $form.Controls.Add($button1)
	
    $button2 = New-Object Windows.Forms.Button
    $button2.Text = "Station Anlegen"
    $button2.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button2.Location = New-Object Drawing.Point @(490, 400)
    $button2.BackColor = [System.Drawing.Color]::Blue 
    $button2.ForeColor = [System.Drawing.Color]::White   
    $button2.Add_Click({ Show-StationCreationPopup -username $username })
	$button2.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	
    $button2.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::LightBlue
    })
    
    $button2.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::Blue
    })
    $form.Controls.Add($button2)

    $button3 = New-Object Windows.Forms.Button
    $button3.Text = "Dokumentation"
    $button3.Size = New-Object Drawing.Size @($buttonWidth, $buttonHeight)
    $button3.Location = New-Object Drawing.Point @(110, 300)
    $button3.BackColor = [System.Drawing.Color]::FromArgb(235, 117, 0)  
    $button3.ForeColor = [System.Drawing.Color]::White  
    $button3.Add_Click({ DokumentationFunction })
	$button3.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	
    $button3.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(255, 177, 77)
    })
    
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
	
    $button4.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(0, 173, 176)
    })
    
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
	# hover effect button
    $button5.Add_MouseEnter({
        $this.BackColor = [System.Drawing.Color]::FromArgb(191, 191, 191)
    })
   
    $button5.Add_MouseLeave({
        $this.BackColor = [System.Drawing.Color]::FromArgb(171, 171, 171)
    })
    $form.Controls.Add($button5)

    $form.add_Closing({
       
        Save-DataFunction
		   
    $timer.Stop()
    $timer.Dispose()
		
		        # Stoppen der Überwachung
        $watcher.EnableRaisingEvents = $false

        
        $_.Cancel = $false  
    })
	Update-LabelStat  # Initialen Online-Status setzen
	Update-FormEncoding

    $form.Topmost = $false
    $form.ShowDialog()
}
function Save-DataFunction {
    $statFilePath = "qCON.go"
    
   
    if (Test-Path $statFilePath) {
        # Benutzername zum Vergleichen
        $usernameToCompare = $username 

        # Inhalt der Datei "stat.txt" lesen
        $content = Get-Content $statFilePath -Raw

        # Muster für den Abschnitt erstellen
        $pattern = "(?ms)^\s*$usernameToCompare.*?^\s*(?:\n|\Z)"

        
        if ($content -match $pattern) {
            
            $content = $content -replace $pattern

            
            $content | Set-Content $statFilePath -Force
            Write-Host "Abschnitt für Benutzer '$usernameToCompare' wurde aus 'qCON.go' entfernt."
        } else {
            Write-Host "Abschnitt für Benutzer '$usernameToCompare' wurde nicht gefunden in 'qCON.go'."
        }
    } else {
        Write-Host "Datei 'qCON.go' wurde nicht gefunden."
    }
}





# ==================================================================== Dokumentation ================================================================================
# ==================================================================== Dokumentation ================================================================================
# ==================================================================== Dokumentation ================================================================================

Import-Module ImportExcel
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

function DokumentationFunction {
    param (
        [string]$stationName
    )
	[System.Windows.Forms.MessageBox]::Show("Wegen Wartungsarbeiten ist dieses Modul gerade nicht aufrufbar.", "Wartung", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
			
}

function dokumaintenance {
	
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = "Stationsinformationen"
    $form.Size = New-Object Drawing.Size @(800, 600)
    $form.StartPosition = "CenterScreen"

    $labelName = New-Object Windows.Forms.Label
    $labelName.Text = "Name: $stationName"
    $labelName.Location = New-Object Drawing.Point @(10, 20)
    $labelName.Size = New-Object Drawing.Size @(400, 20)
    $form.Controls.Add($labelName)

    
    $dataGridView = New-Object Windows.Forms.DataGridView
    $dataGridView.Location = New-Object Drawing.Point @(10, 50)
    $dataGridView.Size = New-Object Drawing.Size @(780, 500)
    $form.Controls.Add($dataGridView)

    
$excelFilePath = "Modul.xlsx"

    try {
        
        $excelApp = New-Object -ComObject Excel.Application
        $workbook = $excelApp.Workbooks.Open($excelFilePath)
        $worksheet = $workbook.Worksheets.Item(1)

        # Lesen Sie die Daten aus der Excel-Tabelle
        $rows = $worksheet.UsedRange.Rows.Count
        $columns = $worksheet.UsedRange.Columns.Count

        # Füllen Sie die DataGridView mit den Daten
        for ($rowIndex = 1; $rowIndex -le $rows; $rowIndex++) {
            $row = New-Object PSObject
            for ($colIndex = 1; $colIndex -le $columns; $colIndex++) {
                $columnName = $worksheet.Cells.Item(1, $colIndex).Text
                $row | Add-Member -MemberType NoteProperty -Name $columnName -Value $worksheet.Cells.Item($rowIndex, $colIndex).Text
            }
            $dataGridView.Rows.Add($row)
        }
    } catch {
        Write-Host "Fehler beim Verarbeiten der Excel-Datei: $_"
    } finally {
        
        $workbook.Close()
        $excelApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    $form.ShowDialog()
}



# ==================================================================== logs ================================================================================
# ==================================================================== logs ================================================================================
# ==================================================================== logs ================================================================================





function Show-LogsFunction {
    try {
        $logsPath = "log.txt"  
        $logsContent = Get-Content -Path $logsPath -Raw
        Show-TextPopup -Title "Logs" -Message $logsContent
    } catch {
        Write-Host "Fehler beim Anzeigen der Logs: $_"
    }
}


function Show-TextPopup {
    param (
        [string]$Title,
        [string]$Message
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(600, 900)  
    $form.StartPosition = "CenterScreen"
	

    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Text = $Message
    $textBox.Size = New-Object Drawing.Size @(560, 830)  
    $textBox.Location = New-Object Drawing.Point @(10, 10)
	$textBox.Font = New-Object Drawing.Font("Arial", 12)
    $form.Controls.Add($textBox)

    $buttonClose = New-Object Windows.Forms.Button
    $buttonClose.Text = "exit"
    $buttonClose.Size = New-Object Drawing.Size @(80, 30)
    $buttonClose.Location = New-Object Drawing.Point @(10, 220)
    $buttonClose.Add_Click({
        $form.Close()
    })
    $form.Controls.Add($buttonClose)

    $form.Topmost = $true
    $form.Add_Shown({
        $textBox.SelectionStart = $textBox.Text.Length
        $textBox.ScrollToCaret()
    })

    $form.ShowDialog()
}

# ==================================================================== Station Suche ================================================================================
# ==================================================================== Station Suche ================================================================================
# ==================================================================== Station Suche ================================================================================




function Show-StationSearchPopup {
    param (
        [string]$Title = "Station Suchen",
        [string]$Message = "Station:"
    )

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(400, 150)  
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

    
    $folderPath = "\Daten\Netzmanagement\SUW\00 Doku"  
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

       
        if ($enteredText -ne "") {
            $form.Hide()
            Show-StationInfo -stationName $enteredText
            $form.Show()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Stationsnamen ein.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
		$form.Close()
        
    })
    $form.Controls.Add($buttonSearch)

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Text = "Abbrechen"
    $buttonCancel.Size = New-Object Drawing.Size @(80, 30)
    $buttonCancel.Location = New-Object Drawing.Point @(100, 80)
    $buttonCancel.Add_Click({
        
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
# ==================================================================== Station Anlegen ================================================================================
# ==================================================================== Station Anlegen ================================================================================
# ==================================================================== Station Anlegen ================================================================================

# Funktion für das Popup-Fenster zum Anlegen einer Station
function Show-StationCreationPopup {
    param (
        [string]$Title = "Station Anlegen",
        [string]$Message = "Neue Station"
    )
	
	$UE = [char]0xDC
	$OE = [char]0xD6
	$AE = [char]0xC4
	$ue = [char]0xFC
	$oe = [char]0xF6
	$ae = [char]0xE4
	
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(800, 650)  # Größe des Fensters erhöht
    $form.StartPosition = "CenterScreen"

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Location = New-Object Drawing.Point @(10, 20)
    $form.Controls.Add($label)

    $labelStationNumber = New-Object Windows.Forms.Label
    $labelStationNumber.Text = "Stationsnummer:"
    $labelStationNumber.Location = New-Object Drawing.Point @(10, 50)
    $form.Controls.Add($labelStationNumber)
	
    $Star = New-Object Windows.Forms.Label
    $Star.Text = "*"
    $Star.ForeColor = [System.Drawing.Color]::Red
	$Star.Font = New-Object Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
	$Star.Location = New-Object Drawing.Point @(600, 80)

    $textBoxStationNumber = New-Object Windows.Forms.TextBox
    $textBoxStationNumber.Location = New-Object Drawing.Point @(150, 50)
    $textBoxStationNumber.Size = New-Object Drawing.Size @(250, 20)
    $textBoxStationNumber.Text = "Stationsnummer..."
    $textBoxStationNumber.Add_Enter({
        if ($textBoxStationNumber.Text -eq "Stationsnummer...") {
            $textBoxStationNumber.Text = ""
            $textBoxStationNumber.ForeColor = [System.Drawing.Color]::black
        }
    })
    $textBoxStationNumber.Add_Leave({
        if ($textBoxStationNumber.Text -eq "") {
            $textBoxStationNumber.Text = "Stationsnummer..."
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
    $textBoxStationName.Text = "Stationsnamen..."
    $textBoxStationName.Add_Enter({
        if ($textBoxStationName.Text -eq "Stationsnamen...") {
            $textBoxStationName.Text = ""
            $textBoxStationName.ForeColor = [System.Drawing.Color]::black
        }
    })
    $textBoxStationName.Add_Leave({
        if ($textBoxStationName.Text -eq "") {
            $textBoxStationName.Text = "Stationsnamen..."
            $textBoxStationName.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    $form.Controls.Add($textBoxStationName)

    $labelTyp = New-Object Windows.Forms.Label
    $labelTyp.Text = "Typ:"
    $labelTyp.Location = New-Object Drawing.Point @(10, 110)
    $form.Controls.Add($labelTyp)

    $comboBoxTyp = New-Object Windows.Forms.ComboBox
    $comboBoxTyp.AutoCompleteMode = "SuggestAppend"
    $comboBoxTyp.AutoCompleteSource = "CustomSource"
    $comboBoxTyp.Location = New-Object Drawing.Point @(150, 110)
    $comboBoxTyp.Size = New-Object Drawing.Size @(250, 20)
    $comboBoxTyp.Items.AddRange(@("Typ 1 RRT", "Typ 2 RRL", "Typ 3 RRT + Absorber", "Typ 4 RRL + Absorber", "Typ 5 RRTT", "Typ 6 RRRT"))
    $form.Controls.Add($comboBoxTyp)
	
	$labelSys = New-Object Windows.Forms.Label
    $labelSys.Text = "Fernwirksystem"
    $labelSys.Location = New-Object Drawing.Point @(10, 140)
    $form.Controls.Add($labelSys)

    $comboBoxSys = New-Object Windows.Forms.ComboBox
    $comboBoxSys.AutoCompleteMode = "SuggestAppend"
    $comboBoxSys.AutoCompleteSource = "CustomSource"
    $comboBoxSys.Location = New-Object Drawing.Point @(150, 140)
    $comboBoxSys.Size = New-Object Drawing.Size @(250, 20)
    $comboBoxSys.Items.AddRange(@("Siemens A8000", "Sprecher T3", "TM", "Pheonix SMS"))
    $form.Controls.Add($comboBoxSys)
	
	
	
    function Add-CheckBox {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [string]$id
    )

    $checkBox = New-Object Windows.Forms.CheckBox
    $checkBox.Text = $text
    $checkBox.Name = "checkBox$id"  
    $checkBox.Location = New-Object Drawing.Point @($x, $y)
    $checkBox.Size = New-Object Drawing.Size @(250, 20)
    $checkBox.Font = New-Object Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
    $form.Controls.Add($checkBox)
    $checkBox.Add_CheckedChanged({
        if ($this.Checked) {
            
            $this.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)  
            $this.Font = New-Object Drawing.Font($this.Font, [System.Drawing.FontStyle]::Bold)
        } else {
            
            $this.ForeColor = [System.Drawing.Color]::Black
            $this.Font = New-Object Drawing.Font($this.Font, [System.Drawing.FontStyle]::Regular)
        }
		})
	}

	
	Add-CheckBox -text "Datenpunktliste Erstellt" -x 10 -y 200 -id 1
	Add-CheckBox -text "Witcom Antrag bearbeitet" -x 10 -y 220 -id 2
	Add-CheckBox -text "Koordinierung Vorabtest" -x 10 -y 240 -id 3
	Add-CheckBox -text "IP Festgelegt" -x 10 -y 260 -id 4
	Add-CheckBox -text "Doku Netzbau bereitgestellt" -x 10 -y 280 -id 4

	
	

    $progressBar = New-Object Windows.Forms.ProgressBar
    $progressBar.Style = "Continuous"  
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0  
	$progressBar.ForeColor = "Green"  
    $progressBar.Location = New-Object Drawing.Point @(100, 550)
    $progressBar.Size = New-Object Drawing.Size @(560, 20)
    $form.Controls.Add($progressBar)
	
	$loaderBar = New-Object Windows.Forms.ProgressBar
    $loaderBar.Style = "Continuous" 
    $LoaderBar.Minimum = 0
    $loaderBar.Maximum = 100
    $loaderBar.Value = 0  
	$loaderBar.ForeColor = "Green"  
    $loaderBar.Location = New-Object Drawing.Point @(100, 550)
    $loaderBar.Size = New-Object Drawing.Size @(560, 20)
    $form.Controls.Add($loaderBar)

    $buttonCreate = New-Object Windows.Forms.Button
    $buttonCreate.Text = "Anlegen"
    $buttonCreate.Size = New-Object Drawing.Size @(80, 30)
    $buttonCreate.Location = New-Object Drawing.Point @(300, 500)
	$buttonCreate.Add_Click({
    $stationNumber = $textBoxStationNumber.Text
    $stationName = $textBoxStationName.Text
    $selectedOption = $comboBoxTyp.SelectedItem

    try {
        
        if ($stationNumber -eq "" -or $stationName -eq "" -or $selectedOption -eq $null) {
            throw "Bitte füllen Sie alle Felder aus."
        }
        $ltige = "ltige"
        
        if ($stationNumber -notmatch '^\d{4}$') {
            throw "Ung$ue$ltige Stationsnummer. Geben Sie eine 4-stellige Zahl ein."
        }

        $destinationPath = $null  

        
        foreach ($drive in [System.IO.DriveInfo]::GetDrives()) {
            $destinationDirectory = "$($drive.Name)\Daten\Netzmanagement\SUW\00 Doku"
            $newFolderName = "$stationNumber $stationName"
            $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $newFolderName

            if (-not (Test-Path -Path $destinationDirectory -PathType Container)) {
                continue  
            }

            
            if (-not (Test-Path -Path $destinationPath -PathType Container)) {
                New-Item -Path $destinationPath -ItemType Directory | Out-Null
                break  
            } else {
                $newFolderName += " - Solid"
                $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $newFolderName
                New-Item -Path $destinationPath -ItemType Directory | Out-Null
                break  
            }
        }

        
        if ($destinationPath -eq $null) {
            throw "Zielverzeichnis wurde nicht gefunden."
        }
		$bertragungstechnik = "bertragungstechnik"
        
        $requiredDirectories = @(
            "02 Batterieanlagen-USV",
            "03 Bilder",
            "09 Kosten",
            "10 Schaltanlagen",
            "12 Schutz- Leittechnik",
            "14 Fernwirktechnik",
            "15 $UE$bertragungstechnik",
            "99 Planung"
        )

        foreach ($dir in $requiredDirectories) {
            $dirPath = Join-Path -Path $destinationPath -ChildPath $dir
            New-Item -Path $dirPath -ItemType Directory | Out-Null
        }

        # Kopieren des Ordners basierend auf der ausgewählten Option
        $uedit = "bergeordnet"
        $sourceDirectory = "$($drive.Name)\Daten\Netzmanagement\SUW\00 Doku\0000 $UE$uedit\50 Netzstationen\$selectedOption"

        if (-not (Test-Path -Path $sourceDirectory -PathType Container)) {
            throw "Quellverzeichnis '$sourceDirectory' existiert nicht."
        }

        # Kopieren der Dateien in das Verzeichnis "14 Fernwirktechnik"
        $sourceFernsteuerung = Join-Path -Path $sourceDirectory -ChildPath "Fernsteuerung"
        $destinationFernsteuerung = Join-Path -Path $destinationPath -ChildPath "14 Fernwirktechnik"
        Copy-Item -Path $sourceFernsteuerung\* -Destination $destinationFernsteuerung -Recurse -Force

        # Durchsuchen und Ersetzen der Dateinamen im neu erstellten Ordner
        $filesFernsteuerung = Get-ChildItem -Path $destinationFernsteuerung -File
        $totalFiles = $filesFernsteuerung.Count
        $progress = 0
        
        foreach ($file in $filesFernsteuerung) {
            $newFileName = $file.Name -replace '^X{1,4}', $stationNumber
            $newFilePath = Join-Path -Path $destinationFernsteuerung -ChildPath $newFileName
            Rename-Item -Path $file.FullName -NewName $newFileName -Force
            
            $progress++
            $percentComplete = [math]::Round(($progress / $totalFiles) * 100)
            Update-ProgressBar -Value $percentComplete
            $form.Refresh()
            
            $datastat++
        }

        # Kopieren der Dateien in das Verzeichnis "10 Schaltanlagen"
        $sourceSchaltanlage = Join-Path -Path $sourceDirectory -ChildPath "Schaltanlage"
        $destinationSchaltanlage = Join-Path -Path $destinationPath -ChildPath "10 Schaltanlagen"
        Copy-Item -Path $sourceSchaltanlage\* -Destination $destinationSchaltanlage -Recurse -Force

        # Erfolgsmeldung 
        [System.Windows.Forms.MessageBox]::Show("Die neue Station wurde erfolgreich angelegt.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } catch {
        Write-Host "Fehler beim Anlegen der Station: $_"
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Anlegen der Station. $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } 
})


function Update-ProgressBar {
    param (
        [int]$Value
    )
    $progressBar.Value = $Value
}
    $form.Controls.Add($buttonCreate)

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Text = "Abbrechen"
    $buttonCancel.Size = New-Object Drawing.Size @(80, 30)
    $buttonCancel.Location = New-Object Drawing.Point @(400, 500)
    $buttonCancel.Add_Click({
        
        $form.Close()
    })
    $form.Controls.Add($buttonCancel)

    $form.Topmost = $true
    $form.ShowDialog()
}


$buttonStationAnlegen.Add_Click({
    Show-StationCreationPopup
})




# ==================================================================== Ticket ================================================================================
# ==================================================================== Ticket ================================================================================
# ==================================================================== Ticket ================================================================================

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
	$form.BackgroundImageLayout = "Stretch"  

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
	$labelUser.ForeColor = [System.Drawing.Color]::White  
	$labelUser.BackColor = [System.Drawing.Color]::FromArgb(2,7,26)
    $form.Controls.Add($labelUser)
	    $labelUserUser = New-Object Windows.Forms.Label
		$labelUserUser.Text = "$username"
		$labelUserUser.Location = New-Object Drawing.Point @(120, 60)
		$labelUserUser.ForeColor = [System.Drawing.Color]::Orange  
		$labelUserUser.BackColor = $transparent1
		$form.Controls.Add($labelUserUser)

    $textBoxDescription = New-Object Windows.Forms.TextBox
	$textBoxDescription.Multiline = $true
	$textBoxDescription.ScrollBars = 'Both'
	$textBoxDescription.Location = New-Object Drawing.Point @(10, 90)
	$textBoxDescription.Size = New-Object Drawing.Size @(400, 200)
	$textBoxDescription.Text = "Beschreibung eingeben..."  
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
	$buttonSend.BackColor = [System.Drawing.Color]::White  
	$buttonSend.ForeColor = [System.Drawing.Color]::Black
    $buttonSend.Add_Click({
        $selectedType = $comboBoxType.SelectedItem
        $description = $textBoxDescription.Text

        # Ticketnummer erhöhen 
        $ticketNumber++
        $ticketNumber | Set-Content $ticketNumberFilePath

        # Betreff 
        $subject = "Ticket #$ticketNumber - SOLID"

        # Outlook-Anwendung erstellen
        $outlook = New-Object -ComObject Outlook.Application

        # Neue Mail 
        $mail = $outlook.CreateItem(0)
        $mail.Subject = $subject
        $mail.Body = "Ein neues Ticket wurde erstellt.`r`nTickettyp: $selectedType`r`nBenutzer: $username`r`nBeschreibung: $description"

        # Empfänger 
        $mail.Recipients.Add("alper.frink@sw-netz.de")

        # Mail senden
        $mail.Send()

        
        $buttonSend.BackColor = [System.Drawing.Color]::Green

        
        $buttonSend.Dispose()
        $buttonCancel.Dispose()

        
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
	$buttonCancel.ForeColor = [System.Drawing.Color]::Black 
	$buttonCancel.BackColor = [System.Drawing.Color]::White 
    $buttonCancel.Add_Click({
        $form.Close()
    })
    $form.Controls.Add($buttonCancel)

    $form.Topmost = $true
    $form.ShowDialog()
}


# ==================================================================== LOGIN ================================================================================
# ==================================================================== LOGIN ================================================================================
# ==================================================================== LOGIN ================================================================================

Add-Type -AssemblyName System.Windows.Forms

# Encoding für die Steuerelemente im Formular 
function Update-FormEncoding {
    $utf8 = [System.Text.Encoding]::UTF8

    $form.Controls | ForEach-Object {
        if ($_ -is [System.Windows.Forms.TextBox] -or $_ -is [System.Windows.Forms.Label] -or $_ -is [System.Windows.Forms.Button]) {
            $_.Text = $utf8.GetString($utf8.GetBytes($_.Text))
        }
		$form.Topmost = $true
    }
}

try {
    Clear-Host

    @"
   APP Running... 
"@
Add-Type -AssemblyName System.Windows.Forms

    # Benutzereingabe 
    do {
        $credential = Get-Credential -Message "Benutzername und Passwort"
        $username = $credential.UserName
        $password = $credential.GetNetworkCredential().Password

        if (-not (Validate-User -username $username -password $password)) {
            Write-Host "Falscher Benutzername oder Passwort. Versuchen Sie es erneut."
			[System.Windows.Forms.MessageBox]::Show("Falscher Benutzername oder Passwort. Versuchen Sie es erneut.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, [System.Windows.Forms.MessageBoxOptions]::DefaultDesktopOnly)
        }
    } while (-not (Validate-User -username $username -password $password))

    # log
    Write-Host "Benutzername: $username"
    Write-Host "Passwort: ****"  

    # Protokolliere den Login 
    $logPath = "log.txt"
    $logEntry = "[$(Get-Date)] login: $username"
    Add-Content -Path $logPath -Value $logEntry -Encoding UTF8

	$onlinepath = "qCON.go"
	Add-Content -Path $onlinepath -Value $username -Encoding UTF8
    
    Show-CustomPopup -Title "SOLID" -username $username

} catch {
    Write-Host "Fehler beim Ausführen des Skripts: $_"
}
