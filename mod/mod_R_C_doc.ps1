Import-Module ImportExcel
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

function DokumentationFunction {
    param (
        [string]$stationName
    )

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

    # Erstellen Sie eine DataGridView
    $dataGridView = New-Object Windows.Forms.DataGridView
    $dataGridView.Location = New-Object Drawing.Point @(10, 50)
    $dataGridView.Size = New-Object Drawing.Size @(780, 500)
    $form.Controls.Add($dataGridView)

    # Pfad zur Excel-Datei
$excelFilePath = "Modul.xlsx"

    try {
        # Erstellen Sie eine Excel-Anwendung und öffnen Sie die Arbeitsmappe
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
        # Schließen Sie die Excel-Anwendung und die Arbeitsmappe
        $workbook.Close()
        $excelApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    $form.ShowDialog()
}

