Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to check if the script is running as administrator
function Test-Admin {
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Restart the script with elevated privileges if not running as administrator
if (-not (Test-Admin)) {
    $newProcess = Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -PassThru
    $newProcess.WaitForExit()
    exit
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "File Permissions Auditor"
$form.Size = New-Object System.Drawing.Size(450, 700)
$form.StartPosition = "CenterScreen"

# Tooltips for helper text
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500
$toolTip.ShowAlways = $true

# Create label and textbox for directory input
$labelDirectory = New-Object System.Windows.Forms.Label
$labelDirectory.Text = "Directory Path:"
$labelDirectory.Location = New-Object System.Drawing.Point(10, 20)
$labelDirectory.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelDirectory)

$textBoxDirectory = New-Object System.Windows.Forms.TextBox
$textBoxDirectory.Location = New-Object System.Drawing.Point(140, 20)
$textBoxDirectory.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBoxDirectory)
$toolTip.SetToolTip($textBoxDirectory, "Enter the path of the directory to audit")

$buttonBrowseDirectory = New-Object System.Windows.Forms.Button
$buttonBrowseDirectory.Text = "Browse"
$buttonBrowseDirectory.Location = New-Object System.Drawing.Point(350, 18)
$buttonBrowseDirectory.Size = New-Object System.Drawing.Size(75, 23)
$form.Controls.Add($buttonBrowseDirectory)
$toolTip.SetToolTip($buttonBrowseDirectory, "Browse to select the directory")

# Create checkbox for "Save as report"
$checkBoxSaveReport = New-Object System.Windows.Forms.CheckBox
$checkBoxSaveReport.Text = "Save as report"
$checkBoxSaveReport.Location = New-Object System.Drawing.Point(10, 60)
$checkBoxSaveReport.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($checkBoxSaveReport)
$toolTip.SetToolTip($checkBoxSaveReport, "Check to save the results as a CSV report")

# Create label and textbox for output file input
$labelOutput = New-Object System.Windows.Forms.Label
$labelOutput.Text = "Output CSV Path:"
$labelOutput.Location = New-Object System.Drawing.Point(10, 90)
$labelOutput.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($labelOutput)

$textBoxOutput = New-Object System.Windows.Forms.TextBox
$textBoxOutput.Location = New-Object System.Drawing.Point(140, 90)
$textBoxOutput.Size = New-Object System.Drawing.Size(200, 20)
$textBoxOutput.Enabled = $false
$textBoxOutput.BackColor = [System.Drawing.SystemColors]::Control
$form.Controls.Add($textBoxOutput)
$toolTip.SetToolTip($textBoxOutput, "Enter the path where the output CSV file should be saved")

$buttonBrowseOutput = New-Object System.Windows.Forms.Button
$buttonBrowseOutput.Text = "Browse"
$buttonBrowseOutput.Location = New-Object System.Drawing.Point(350, 88)
$buttonBrowseOutput.Size = New-Object System.Drawing.Size(75, 23)
$buttonBrowseOutput.Enabled = $false
$form.Controls.Add($buttonBrowseOutput)
$toolTip.SetToolTip($buttonBrowseOutput, "Browse to select the output file path")

# Create checkbox for recursive option
$checkBoxRecursive = New-Object System.Windows.Forms.CheckBox
$checkBoxRecursive.Text = "Include Folders Recursively"
$checkBoxRecursive.Location = New-Object System.Drawing.Point(10, 120)
$checkBoxRecursive.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($checkBoxRecursive)
$toolTip.SetToolTip($checkBoxRecursive, "Include all folders and subfolders recursively")

# Create checkbox for all files option
$checkBoxAllFiles = New-Object System.Windows.Forms.CheckBox
$checkBoxAllFiles.Text = "Include All Files"
$checkBoxAllFiles.Location = New-Object System.Drawing.Point(10, 150)
$checkBoxAllFiles.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($checkBoxAllFiles)
$toolTip.SetToolTip($checkBoxAllFiles, "Include all files in the directory")

# Create labels to display the number of files and folders
$labelFilesCount = New-Object System.Windows.Forms.Label
$labelFilesCount.Text = "Files: 0"
$labelFilesCount.Location = New-Object System.Drawing.Point(10, 180)
$labelFilesCount.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelFilesCount)

$labelFoldersCount = New-Object System.Windows.Forms.Label
$labelFoldersCount.Text = "Folders: 0"
$labelFoldersCount.Location = New-Object System.Drawing.Point(120, 180)
$labelFoldersCount.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($labelFoldersCount)

# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 210)
$progressBar.Size = New-Object System.Drawing.Size(415, 20)
$form.Controls.Add($progressBar)

# Create button to run audit
$buttonRunAudit = New-Object System.Windows.Forms.Button
$buttonRunAudit.Text = "Run Audit"
$buttonRunAudit.Location = New-Object System.Drawing.Point(10, 240)
$buttonRunAudit.Size = New-Object System.Drawing.Size(200, 30)
$form.Controls.Add($buttonRunAudit)
$toolTip.SetToolTip($buttonRunAudit, "Run the audit and display the results")

# Create button to save audit
$buttonSaveAudit = New-Object System.Windows.Forms.Button
$buttonSaveAudit.Text = "Save Audit"
$buttonSaveAudit.Location = New-Object System.Drawing.Point(230, 240)
$buttonSaveAudit.Size = New-Object System.Drawing.Size(100, 30)
$buttonSaveAudit.Enabled = $false
$form.Controls.Add($buttonSaveAudit)
$toolTip.SetToolTip($buttonSaveAudit, "Save the audit results to a CSV file")

# Create a text box to display the status
$textBoxStatus = New-Object System.Windows.Forms.TextBox
$textBoxStatus.Location = New-Object System.Drawing.Point(10, 280)
$textBoxStatus.Size = New-Object System.Drawing.Size(415, 360)
$textBoxStatus.Multiline = $true
$textBoxStatus.ScrollBars = "Vertical"
$form.Controls.Add($textBoxStatus)
$toolTip.SetToolTip($textBoxStatus, "Displays the status of the audit process")

# Log file path
$logFilePath = "AuditLog.txt"

# Function to write to log file
function Write-Log {
    param (
        [string]$message,
        [bool]$isComplete = $false,
        [bool]$isHighlighted = $false
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage

    if ($isComplete) {
        $originalColor = $textBoxStatus.ForeColor
        $textBoxStatus.ForeColor = [System.Drawing.Color]::Green
        $textBoxStatus.AppendText("$logMessage`r`n")
        $textBoxStatus.ForeColor = $originalColor
    } elseif ($isHighlighted) {
        $originalColor = $textBoxStatus.ForeColor
        $textBoxStatus.ForeColor = [System.Drawing.Color]::Red
        $textBoxStatus.AppendText("$logMessage`r`n")
        $textBoxStatus.ForeColor = $originalColor
    } else {
        $textBoxStatus.AppendText("$logMessage`r`n")
    }
}

# Browse for directory
$buttonBrowseDirectory.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxDirectory.Text = $folderBrowser.SelectedPath
    }
})

# Browse for output file
$buttonBrowseOutput.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "csv"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxOutput.Text = $saveFileDialog.FileName
    }
})

# Function to enable or disable the Save Audit button based on input fields
function UpdateAuditButtonState {
    $buttonSaveAudit.Enabled = $checkBoxSaveReport.Checked -and -not [string]::IsNullOrEmpty($textBoxDirectory.Text) -and -not [string]::IsNullOrEmpty($textBoxOutput.Text)
    if ($checkBoxSaveReport.Checked) {
        $textBoxOutput.Enabled = $true
        $buttonBrowseOutput.Enabled = $true
        $textBoxOutput.BackColor = [System.Drawing.SystemColors]::Window
    } else {
        $textBoxOutput.Enabled = $false
        $buttonBrowseOutput.Enabled = $false
        $textBoxOutput.BackColor = [System.Drawing.SystemColors]::Control
    }
}

# Add event handlers to update the button state when text changes
$textBoxDirectory.Add_TextChanged({ UpdateAuditButtonState })
$textBoxOutput.Add_TextChanged({ UpdateAuditButtonState })
$checkBoxSaveReport.Add_CheckedChanged({ UpdateAuditButtonState })

# Function to audit file permissions
function Audit-FilePermissions {
    param (
        [string]$path,
        [string]$outputCsvPath,
        [bool]$recursive,
        [bool]$includeAllFiles,
        [bool]$saveReport
    )

    $items = @()
    $filesCount = 0
    $foldersCount = 0

    Write-Log "Starting audit for path: $path"

    if ($recursive) {
        $items += Get-ChildItem -Path $path -Recurse -Directory -ErrorAction SilentlyContinue
    } else {
        $items += Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue
    }
    
    if ($includeAllFiles) {
        $items += Get-ChildItem -Path $path -Recurse:$recursive -File -ErrorAction SilentlyContinue
    }

    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            $foldersCount++
        } else {
            $filesCount++
        }
    }

    $labelFilesCount.Text = "Files: $filesCount"
    $labelFoldersCount.Text = "Folders: $foldersCount"
    $totalItems = $items.Count
    $progressBar.Maximum = $totalItems
    $progressBar.Value = 0

    $results = @()
    $itemCount = 0
    foreach ($item in $items) {
        try {
            $acl = Get-Acl -Path $item.FullName
            $permissions = $acl.Access
            $owner = $acl.Owner

            foreach ($access in $permissions) {
                if ($access.IdentityReference -notmatch "^(NT AUTHORITY|NT SERVICE|BUILTIN)" -and $access.IdentityReference -ne "NT AUTHORITY\SYSTEM") {
                    $isHighlighted = $false
                    if ($access.IdentityReference -eq "Everyone" -and $access.FileSystemRights -eq "FullControl") {
                        $isHighlighted = $true
                    }
                    $results += [PSCustomObject]@{
                        Path              = $item.FullName
                        Owner             = $owner
                        Identity          = $access.IdentityReference
                        FileSystemRights  = $access.FileSystemRights
                        AccessControlType = $access.AccessControlType
                        InheritanceFlags  = $access.InheritanceFlags
                        PropagationFlags  = $access.PropagationFlags
                        IsInherited       = $access.IsInherited
                        AuditDate         = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    Write-Log "Processed: $($item.FullName) - Identity: $($access.IdentityReference) - Rights: $($access.FileSystemRights)" $false $isHighlighted
                }
            }
        } catch {
            Write-Log "Error processing file: $($item.FullName). Error: $_"
        }
        $itemCount++
        $progressBar.Value = $itemCount
    }

    if ($saveReport -and -not [string]::IsNullOrEmpty($outputCsvPath)) {
        $results | Export-Csv -Path $outputCsvPath -NoTypeInformation
        Write-Log "Audit completed. Results saved to $outputCsvPath" $true
    } else {
        Write-Log "Audit completed for path: $path" $true
    }

    return $results
}

# Function to display audit results in a custom DataGridView with filter functionality
function Show-Results {
    param (
        [System.Collections.ArrayList]$results
    )

    # Create a new form
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Audit Results"
    $resultsForm.Size = New-Object System.Drawing.Size(800, 600)
    $resultsForm.StartPosition = "CenterScreen"

    # Create a TextBox for filtering
    $filterTextBox = New-Object System.Windows.Forms.TextBox
    $filterTextBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $resultsForm.Controls.Add($filterTextBox)

    # Create a DataGridView
    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dataGridView.ReadOnly = $true
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false

    # Add columns to the DataGridView
    $columns = @("Path", "Owner", "Identity", "FileSystemRights", "AccessControlType", "InheritanceFlags", "PropagationFlags", "IsInherited", "AuditDate")
    foreach ($column in $columns) {
        $dataGridView.Columns.Add($column, $column)
    }

    # Add rows to the DataGridView
    $dataTable = New-Object System.Data.DataTable
    foreach ($column in $columns) {
        $dataTable.Columns.Add($column)
    }

    foreach ($result in $results) {
        $row = $dataTable.NewRow()
        $row["Path"] = $result.Path
        $row["Owner"] = $result.Owner
        $row["Identity"] = $result.Identity
        $row["FileSystemRights"] = $result.FileSystemRights
        $row["AccessControlType"] = $result.AccessControlType
        $row["InheritanceFlags"] = $result.InheritanceFlags
        $row["PropagationFlags"] = $result.PropagationFlags
        $row["IsInherited"] = $result.IsInherited
        $row["AuditDate"] = $result.AuditDate
        $dataTable.Rows.Add($row)
    }

    $dataGridView.DataSource = $dataTable

    # Highlight rows where Identity is Everyone and FileSystemRights is FullControl
    $dataGridView.DataBindingComplete.Add({
        foreach ($row in $dataGridView.Rows) {
            if ($row.Cells["Identity"].Value -eq "Everyone" -and $row.Cells["FileSystemRights"].Value -eq "FullControl") {
                foreach ($cell in $row.Cells) {
                    $cell.Style.BackColor = [System.Drawing.Color]::Red
                }
            }
        }
    })

    # Add filtering functionality
    $filterTextBox.Add_TextChanged({
        $dataTable.DefaultView.RowFilter = "Path LIKE '%$($filterTextBox.Text)%' OR Owner LIKE '%$($filterTextBox.Text)%' OR Identity LIKE '%$($filterTextBox.Text)%' OR FileSystemRights LIKE '%$($filterTextBox.Text)%' OR AccessControlType LIKE '%$($filterTextBox.Text)%' OR InheritanceFlags LIKE '%$($filterTextBox.Text)%' OR PropagationFlags LIKE '%$($filterTextBox.Text)%' OR IsInherited LIKE '%$($filterTextBox.Text)%' OR AuditDate LIKE '%$($filterTextBox.Text)%'"
    })

    # Add the DataGridView to the form
    $resultsForm.Controls.Add($dataGridView)

    # Show the form
    $resultsForm.ShowDialog()
}

# Add click event to the run audit button
$results = $null
$buttonRunAudit.Add_Click({
    $directoryPath = $textBoxDirectory.Text
    $recursive = $checkBoxRecursive.Checked
    $includeAllFiles = $checkBoxAllFiles.Checked

    if ([string]::IsNullOrEmpty($directoryPath)) {
        [System.Windows.Forms.MessageBox]::Show("Please provide the directory path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        $textBoxStatus.Clear()
        Write-Log "Starting audit..."
        $results = Audit-FilePermissions -path $directoryPath -outputCsvPath $null -recursive $recursive -includeAllFiles $includeAllFiles -saveReport $false
        if ($results -and $results.Count -gt 0) {
            Show-Results -results $results
        } else {
            [System.Windows.Forms.MessageBox]::Show("No results to display.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Add click event to the save audit button
$buttonSaveAudit.Add_Click({
    $directoryPath = $textBoxDirectory.Text
    $outputCsvPath = $textBoxOutput.Text
    $recursive = $checkBoxRecursive.Checked
    $includeAllFiles = $checkBoxAllFiles.Checked
    $saveReport = $checkBoxSaveReport.Checked

    if ([string]::IsNullOrEmpty($directoryPath) -or ($saveReport -and [string]::IsNullOrEmpty($outputCsvPath))) {
        [System.Windows.Forms.MessageBox]::Show("Please provide the necessary paths.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        $textBoxStatus.Clear()
        Write-Log "Starting audit and saving report..."
        $results = Audit-FilePermissions -path $directoryPath -outputCsvPath $outputCsvPath -recursive $recursive -includeAllFiles $includeAllFiles -saveReport $saveReport
        $buttonSaveAudit.Text = "Rerun Audit and Save"
    }
})

# Show the form
$form.ShowDialog()
