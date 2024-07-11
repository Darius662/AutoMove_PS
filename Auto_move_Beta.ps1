# Define the function to handle the button click event
function Start-MoveFiles {
    # Get user-selected folder
    $parentFolder = $folderTextBox.Text

    # Get user-selected start and end date/time for moving files
    $moveStartDate = $startDatePicker.Value.Date
    $moveStartTime = $startTimePicker.Value.TimeOfDay
    $moveStartTimestamp = $moveStartDate.Add($moveStartTime)

    $moveEndDate = $endDatePicker.Value.Date
    $moveEndTime = $endTimePicker.Value.TimeOfDay
    $moveEndTimestamp = $moveEndDate.Add($moveEndTime)

    # Set default CSV file name
    $timestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
    $csvFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "PA_MOVE_$timestamp.csv")

    # Initialize an array to store the results
    $results = @()

    # Get all the '_fehler' folders within the parent folder and its subfolders
    $errorFolders = Get-ChildItem -Path $parentFolder -Filter "_fehler" -Recurse -Directory

    foreach ($errorFolder in $errorFolders) {
        # Get a list of files in the error folder
        $files = Get-ChildItem -Path $errorFolder.FullName -File

        foreach ($file in $files) {
            # Check if the file is within the user-selected timestamp range
            if ($file.LastWriteTime -gt $moveStartTimestamp -and $file.LastWriteTime -lt $moveEndTimestamp) {
                # Move the file to the parent folder
                Move-Item -Path $file.FullName -Destination $errorFolder.Parent.FullName

                # Add the result to the results array
                $results += [PSCustomObject]@{
                    FileName       = $file.Name
                    OriginalFolder = $errorFolder.FullName
                    Destination    = $errorFolder.Parent.FullName
                    MovedTime      = (Get-Date)
                }
            }
        }
    }

    # Export the results to a CSV file
    $results | Export-Csv -Path $csvFilePath -NoTypeInformation

    # Show a message box indicating the operation is complete
    [System.Windows.Forms.MessageBox]::Show("Files moved and results exported to $csvFilePath", "Operation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Define the function to handle the button click event for 'Help'
function Show-Help {
    [System.Windows.Forms.MessageBox]::Show(
        "This app allows you to perform the following operations:`n`n" +
        "1. Move files within the specified timestamp range from 'error' folders to their parent folder.`n" +
        "2. Check files in '_fehler' folders and export information to a CSV without moving them.`n`n" +
        "To get started, follow these steps:`n`n" +
        "1. Select the folder containing 'error' folders.`n" +
        "2. Choose the desired start and end date/time for the interval.`n" +
        "3. Select the location to save the CSV file.`n" +
        "4. Click 'Start' to begin moving files or 'Check' to export information without moving.`n`n" +
        "For additional assistance, contact Jeleru Darius, (Darius.Jeleru@partner.bmw.de) (qxz3m5t).",
        "Help",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# Define the function to handle Check Button
function Check-Files {
    # Get user-selected folder
    $parentFolder = $folderTextBox.Text
    # Get user-selected CSV file path
    $csvFilePath = $csvPathTextBox.Text
    # Initialize an array to store the results
    $results = @()
    # Get all the 'error' folders within the parent folder and its subfolders
    $errorFolders = Get-ChildItem -Path $parentFolder -Filter "_fehler" -Recurse -Directory

    foreach ($errorFolder in $errorFolders) {
        # Get a list of files in the error folder
        $files = Get-ChildItem -Path $errorFolder.FullName -File
        foreach ($file in $files) {
            # Add the file information to the results array
            $results += [PSCustomObject]@{
                FileName       = $file.Name
                OriginalFolder = $errorFolder.FullName
                LastWriteTime  = $file.LastWriteTime
            }
        }
    }
    # Export the results to a CSV file
    $results | Export-Csv -Path $csvFilePath -NoTypeInformation
    # Show a message box indicating the operation is complete
    [System.Windows.Forms.MessageBox]::Show("File information exported to $csvFilePath", "Operation Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Create the main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "(Beta v0.2)Move _Fehler Dateien"
$mainForm.Size = New-Object System.Drawing.Size(460, 320)
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$mainForm.MaximizeBox = $false
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# Create a label and textbox for selecting the folder
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "Select Folder:"
$folderLabel.Location = New-Object System.Drawing.Point(10, 20)
$mainForm.Controls.Add($folderLabel)

$folderTextBox = New-Object System.Windows.Forms.TextBox
$folderTextBox.Location = New-Object System.Drawing.Point(120, 20)
$folderTextBox.Size = New-Object System.Drawing.Size(200, 50)
$mainForm.Controls.Add($folderTextBox)

$folderBrowseButton = New-Object System.Windows.Forms.Button
$folderBrowseButton.Text = "Browse..."
$folderBrowseButton.Location = New-Object System.Drawing.Point(330, 20)
$folderBrowseButton.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $folderBrowserDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $folderTextBox.Text = $folderBrowserDialog.SelectedPath
    }
})
$mainForm.Controls.Add($folderBrowseButton)

# Create date and time pickers for selecting the start date/time
$startDateLabel = New-Object System.Windows.Forms.Label
$startDateLabel.Text = "Select Start Date:"
$startDateLabel.Location = New-Object System.Drawing.Point(10, 60)
$mainForm.Controls.Add($startDateLabel)

$startDatePicker = New-Object System.Windows.Forms.DateTimePicker
$startDatePicker.Location = New-Object System.Drawing.Point(120, 60)
$startDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$mainForm.Controls.Add($startDatePicker)

$startTimeLabel = New-Object System.Windows.Forms.Label
$startTimeLabel.Text = "Select Start Time:"
$startTimeLabel.Location = New-Object System.Drawing.Point(10, 100)
$mainForm.Controls.Add($startTimeLabel)

$startTimePicker = New-Object System.Windows.Forms.DateTimePicker
$startTimePicker.Location = New-Object System.Drawing.Point(120, 100)
$startTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$startTimePicker.ShowUpDown = $true
$mainForm.Controls.Add($startTimePicker)

# Create date and time pickers for selecting the end date/time
$endDateLabel = New-Object System.Windows.Forms.Label
$endDateLabel.Text = "Select End Date:"
$endDateLabel.Location = New-Object System.Drawing.Point(10, 140)
$mainForm.Controls.Add($endDateLabel)

$endDatePicker = New-Object System.Windows.Forms.DateTimePicker
$endDatePicker.Location = New-Object System.Drawing.Point(120, 140)
$endDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$mainForm.Controls.Add($endDatePicker)

$endTimeLabel = New-Object System.Windows.Forms.Label
$endTimeLabel.Text = "Select End Time:"
$endTimeLabel.Location = New-Object System.Drawing.Point(10, 180)
$mainForm.Controls.Add($endTimeLabel)

$endTimePicker = New-Object System.Windows.Forms.DateTimePicker
$endTimePicker.Location = New-Object System.Drawing.Point(120, 180)
$endTimePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$endTimePicker.ShowUpDown = $true
$mainForm.Controls.Add($endTimePicker)

# Create a label and textbox for entering the CSV file path
$csvPathLabel = New-Object System.Windows.Forms.Label
$csvPathLabel.Text = "CSV File Path:"
$csvPathLabel.Location = New-Object System.Drawing.Point(10, 220)
$mainForm.Controls.Add($csvPathLabel)

# Add this line to get the timestamp
$timestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"

# Add this line to set the default CSV file path
$csvFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "PA_MOVE_$timestamp.csv")

$csvPathTextBox = New-Object System.Windows.Forms.TextBox
$csvPathTextBox.Location = New-Object System.Drawing.Point(120, 220)
$csvPathTextBox.Size = New-Object System.Drawing.Size(200, 50)
# Set the text of the textbox to the generated path
$csvPathTextBox.Text = $csvFilePath
$mainForm.Controls.Add($csvPathTextBox)

$csvPathBrowseButton = New-Object System.Windows.Forms.Button
$csvPathBrowseButton.Text = "Browse..."
$csvPathBrowseButton.Location = New-Object System.Drawing.Point(330, 220)
$csvPathBrowseButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.FileName = "PA_Move.csv"
    $result = $saveFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPathTextBox.Text = $saveFileDialog.FileName
    }
})
$mainForm.Controls.Add($csvPathBrowseButton)

# Create the Start, Cancel, and Close buttons
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start"
$startButton.Location = New-Object System.Drawing.Point(10, 260)
$startButton.Add_Click({ Start-MoveFiles })
$mainForm.Controls.Add($startButton)

# Create the Help button
$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Text = "Help"
$helpButton.Location = New-Object System.Drawing.Point(10, 1)
$helpButton.Add_Click({ Show-Help })
$mainForm.Controls.Add($helpButton)

# Create the Close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Location = New-Object System.Drawing.Point(190, 260)
$closeButton.Add_Click({ $mainForm.Close() })
$mainForm.Controls.Add($closeButton)

# Create the Check button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Text = "Check"
$checkButton.Location = New-Object System.Drawing.Point(100, 260)
$checkButton.Add_Click({ Check-Files })
$mainForm.Controls.Add($checkButton)

# Show the form
$mainForm.ShowDialog() | Out-Null
