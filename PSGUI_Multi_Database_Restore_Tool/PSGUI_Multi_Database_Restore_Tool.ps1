Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

# === Form ===
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Multi Database Restore Tool'
$form.Size = New-Object System.Drawing.Size(1200, 650)
$form.StartPosition = 'CenterScreen'

# === Server Controls ===
$lblServer = New-Object System.Windows.Forms.Label -Property @{ Text = "SQL Server Name:"; Location = '20,20'; Size = '120,20' }
$txtServer = New-Object System.Windows.Forms.TextBox -Property @{ Location = '150,20'; Size = '250,20'; Text = '(local)' }
$btnCheckConnection = New-Object System.Windows.Forms.Button -Property @{ Text = "Check Connectivity"; Location = '420,18'; Size = '130,25' }
$lblStatus = New-Object System.Windows.Forms.Label -Property @{ Text = "Status: Not Connected"; Location = '20,50'; Size = '700,20'; ForeColor = 'Red' }
$form.Controls.AddRange(@($lblServer, $txtServer, $btnCheckConnection, $lblStatus))

# === Folder Inputs ===
$lblBackup = New-Object System.Windows.Forms.Label -Property @{ Text = 'Backup Location:'; Location = '20,80'; Size = '120,20' }
$txtBackup = New-Object System.Windows.Forms.TextBox -Property @{ Location = '150,80'; Size = '400,20' }
$btnBrowseBackup = New-Object System.Windows.Forms.Button -Property @{ Text = 'Browse'; Location = '560,78'; Size = '80,24' }

$lblMDF = New-Object System.Windows.Forms.Label -Property @{ Text = 'MDF File Location:'; Location = '20,120'; Size = '120,20' }
$txtMDF = New-Object System.Windows.Forms.TextBox -Property @{ Location = '150,120'; Size = '400,20' }
$btnBrowseMDF = New-Object System.Windows.Forms.Button -Property @{ Text = 'Browse'; Location = '560,118'; Size = '80,24' }

$lblLDF = New-Object System.Windows.Forms.Label -Property @{ Text = 'LDF File Location:'; Location = '20,160'; Size = '120,20' }
$txtLDF = New-Object System.Windows.Forms.TextBox -Property @{ Location = '150,160'; Size = '400,20' }
$btnBrowseLDF = New-Object System.Windows.Forms.Button -Property @{ Text = 'Browse'; Location = '560,158'; Size = '80,24' }

$form.Controls.AddRange(@($lblBackup, $txtBackup, $btnBrowseBackup, $lblMDF, $txtMDF, $btnBrowseMDF, $lblLDF, $txtLDF, $btnBrowseLDF))

# === DB Grid ===
$lblDBList = New-Object System.Windows.Forms.Label -Property @{ Text = 'Databases Found:'; Location = '20,200'; Size = '200,20' }

$gridDBList = New-Object System.Windows.Forms.DataGridView
$gridDBList.Location = New-Object System.Drawing.Point(20,230)
$gridDBList.Size = New-Object System.Drawing.Size(740,200)
$gridDBList.AllowUserToAddRows = $false
$gridDBList.AllowUserToDeleteRows = $false
$gridDBList.RowHeadersVisible = $false
$gridDBList.SelectionMode = 'FullRowSelect'
$gridDBList.AutoSizeColumnsMode = 'Fill'

$colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelect.HeaderText = "Select"
$colSelect.Width = 50
$gridDBList.Columns.Add($colSelect)

$colOriginal = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOriginal.HeaderText = "Original DB Name"
$colOriginal.ReadOnly = $true
$gridDBList.Columns.Add($colOriginal)

$colNewName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNewName.HeaderText = "New DB Name (Optional)"
$gridDBList.Columns.Add($colNewName)

$form.Controls.AddRange(@($lblDBList, $gridDBList))

# === Buttons ===
$btnSelectAll = New-Object System.Windows.Forms.Button -Property @{ Text = 'Select All'; Location = '20,450'; Size = '100,30' }
$btnDeselectAll = New-Object System.Windows.Forms.Button -Property @{ Text = 'Deselect All'; Location = '130,450'; Size = '120,30' }
$btnRestore = New-Object System.Windows.Forms.Button -Property @{ Text = 'Restore Databases'; Location = '270,450'; Size = '150,30' }
$btnShowDBNames = New-Object System.Windows.Forms.Button -Property @{ Text = 'Show DB Names'; Location = '440,450'; Size = '150,30' }
$btnDropDBs = New-Object System.Windows.Forms.Button -Property @{ Text = 'Drop Selected DBs'; Location = '600,450'; Size = '150,30' }
$form.Controls.AddRange(@($btnSelectAll, $btnDeselectAll, $btnRestore, $btnShowDBNames, $btnDropDBs))

# === Progress Bar ===
$progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{
    Location = '20,490'; Size = '740,15'; Minimum = 0; Maximum = 100; Value = 0
}
$form.Controls.Add($progressBar)

# === Status Log ===
$txtStatus = New-Object System.Windows.Forms.TextBox -Property @{
    Location = '800,20'; Size = '470,600'; Multiline = $true; ScrollBars = 'Vertical'; ReadOnly = $true
}
$form.Controls.Add($txtStatus)

# === Events ===

$btnCheckConnection.Add_Click({
    try {
        $connStr = "Server=$($txtServer.Text.Trim());Integrated Security=True;TrustServerCertificate=True;"
        $conn = New-Object System.Data.SqlClient.SqlConnection $connStr
        $conn.Open()
        $lblStatus.Text = "✅ Connected to: " + $conn.DataSource
        $lblStatus.ForeColor = 'Green'
        $txtStatus.AppendText("$(Get-Date -Format 'HH:mm:ss') - Connected to $($txtServer.Text)`r`n")
        $conn.Close()
    } catch {
        $lblStatus.Text = "❌ Connection Failed: $($_.Exception.Message)"
        $lblStatus.ForeColor = 'Red'
        $txtStatus.AppendText("$(Get-Date -Format 'HH:mm:ss') - Connection failed: $($_.Exception.Message)`r`n")
    }
})

$btnBrowseBackup.Add_Click({ $f = New-Object System.Windows.Forms.FolderBrowserDialog; if ($f.ShowDialog() -eq "OK") { $txtBackup.Text = $f.SelectedPath } })
$btnBrowseMDF.Add_Click({ $f = New-Object System.Windows.Forms.FolderBrowserDialog; if ($f.ShowDialog() -eq "OK") { $txtMDF.Text = $f.SelectedPath } })
$btnBrowseLDF.Add_Click({ $f = New-Object System.Windows.Forms.FolderBrowserDialog; if ($f.ShowDialog() -eq "OK") { $txtLDF.Text = $f.SelectedPath } })

$btnShowDBNames.Add_Click({
    $gridDBList.Rows.Clear()
    $backupFolder = $txtBackup.Text.Trim()
    if (-not (Test-Path $backupFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid Backup Location path.", "Error", "OK", "Error")
        return
    }

    $connStr = "Server=$($txtServer.Text.Trim());Integrated Security=True;TrustServerCertificate=True;"
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStr

    try {
        $sqlConnection.Open()
        $backupFiles = Get-ChildItem -Path $backupFolder -Filter *.bak
        if ($backupFiles.Count -eq 0) {
            $txtStatus.AppendText("No .bak files found in $backupFolder`r`n")
            return
        }

        foreach ($file in $backupFiles) {
            $query = "RESTORE HEADERONLY FROM DISK = N'$($file.FullName.Replace("'", "''"))';"
            $cmd = $sqlConnection.CreateCommand()
            $cmd.CommandText = $query
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
            $dt = New-Object System.Data.DataTable
            $adapter.Fill($dt) | Out-Null

            foreach ($row in $dt.Rows) {
                $dbName = $row["DatabaseName"]
                if (-not ($gridDBList.Rows | Where-Object { $_.Cells[1].Value -eq $dbName })) {
                    $gridDBList.Rows.Add($true, $dbName, "")
                    $txtStatus.AppendText("Found database: $dbName from $($file.Name)`r`n")
                }
            }
        }
        $sqlConnection.Close()
    } catch {
        $txtStatus.AppendText("Error reading backup: $($_.Exception.Message)`r`n")
        $sqlConnection.Close()
    }
})

$btnSelectAll.Add_Click({ foreach ($row in $gridDBList.Rows) { $row.Cells[0].Value = $true } })
$btnDeselectAll.Add_Click({ foreach ($row in $gridDBList.Rows) { $row.Cells[0].Value = $false } })

$btnRestore.Add_Click({
    $backupFolder = $txtBackup.Text.Trim()
    $mdfFolder = $txtMDF.Text.Trim()
    $ldfFolder = $txtLDF.Text.Trim()

    if (-not (Test-Path $backupFolder) -or -not (Test-Path $mdfFolder) -or -not (Test-Path $ldfFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Ensure all folder paths are valid.", "Path Error", "OK", "Error")
        return
    }

    $selectedRows = $gridDBList.Rows | Where-Object { $_.Cells[0].Value -eq $true }
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one database to restore.", "No Selection", "OK", "Warning")
        return
    }

    $connStr = "Server=$($txtServer.Text.Trim());Integrated Security=True;TrustServerCertificate=True;"
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStr

    try {
        $sqlConnection.Open()
        $total = $selectedRows.Count
        $current = 0
        $progressBar.Value = 0

        foreach ($row in $selectedRows) {
            $originalName = $row.Cells[1].Value
            $newName = if ([string]::IsNullOrWhiteSpace($row.Cells[2].Value)) { $originalName } else { $row.Cells[2].Value }

            $current++
            $progressBar.Value = [math]::Round(($current / $total) * 100)
            $progressBar.Refresh()

            $bakFile = Get-ChildItem -Path $backupFolder -Filter *.bak | Where-Object {
                $filePath = $_.FullName.Replace("'", "''")
                $query = "RESTORE HEADERONLY FROM DISK = N'$filePath';"
                $cmd = $sqlConnection.CreateCommand()
                $cmd.CommandText = $query
                $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
                $dt = New-Object System.Data.DataTable
                $adapter.Fill($dt) | Out-Null
                $dt.Rows | Where-Object { $_["DatabaseName"] -eq $originalName }
            } | Select-Object -First 1

            if ($null -eq $bakFile) {
                $txtStatus.AppendText("❌ No backup found for ${originalName}`r`n")
                continue
            }

            $mdfPath = Join-Path $mdfFolder "$newName.mdf"
            $ldfPath = Join-Path $ldfFolder "$newName.ldf"

            $restoreQuery = @"
RESTORE DATABASE [${newName}]
FROM DISK = N'$($bakFile.FullName.Replace("'", "''"))'
WITH MOVE '${originalName}' TO N'$mdfPath',
     MOVE '${originalName}_log' TO N'$ldfPath',
     REPLACE
"@

            try {
                $cmd = $sqlConnection.CreateCommand()
                $cmd.CommandText = $restoreQuery
                $cmd.ExecuteNonQuery() | Out-Null
                $txtStatus.AppendText("✅ Restored ${originalName} as ${newName}`r`n")
            } catch {
                $txtStatus.AppendText("❌ Failed to restore ${originalName}: $($_.Exception.Message)`r`n")
            }
        }

        $progressBar.Value = 100
        $sqlConnection.Close()
    } catch {
        $txtStatus.AppendText("❌ Error: $($_.Exception.Message)`r`n")
        $progressBar.Value = 0
        $sqlConnection.Close()
    }
})

$btnDropDBs.Add_Click({
    $selectedRows = $gridDBList.Rows | Where-Object { $_.Cells[0].Value -eq $true }
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one database to drop.", "No Selection", "OK", "Warning")
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to drop the selected databases? This action is irreversible!", "Confirm Drop", "YesNo", "Warning")
    if ($confirm -ne 'Yes') { return }

    $connStr = "Server=$($txtServer.Text.Trim());Integrated Security=True;TrustServerCertificate=True;"
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $connStr

    try {
        $sqlConnection.Open()
        foreach ($row in $selectedRows) {
            $dbName = $row.Cells[1].Value
            try {
                $dropQuery = "ALTER DATABASE [$dbName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [$dbName];"
                $cmd = $sqlConnection.CreateCommand()
                $cmd.CommandText = $dropQuery
                $cmd.ExecuteNonQuery() | Out-Null
                $txtStatus.AppendText("🗑️ Dropped database: ${dbName}`r`n")
            } catch {
                $txtStatus.AppendText("❌ Failed to drop ${dbName}: $($_.Exception.Message)`r`n")
            }
        }
        $sqlConnection.Close()
    } catch {
        $txtStatus.AppendText("❌ Connection or drop error: $($_.Exception.Message)`r`n")
    }
})

# === Show Form ===
$form.ShowDialog()
