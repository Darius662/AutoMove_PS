# Define the function to handle the button click event
function Start-MoveFiles {
    # Get user-selected folder
    $parentFolder = $folderTextBox.Text

    # Get user-selected time interval for moving files
    $timeInterval = $timePicker.SelectedItem

    if ($timeInterval -eq "Custom") {
        # Get user-selected date and time for moving files
        $moveDate = $datePicker.Value.Date
        $moveTime = $timePickerCustom.Value.TimeOfDay
        $moveTimestamp = $moveDate.Add($moveTime)
    } else {
        # Convert the selected time interval to minutes
        $minutes = @{
            "5 Minutes"  = 5
            "10 Minutes" = 10
            "15 Minutes" = 15
            "20 Minutes" = 20
            "30 Minutes" = 30
        }[$timeInterval]

        # Calculate the timestamp based on the selected time interval
        $moveTimestamp = (Get-Date).AddMinutes(-$minutes)
    }
    # Define the function to handle the button click event for 'Check'

    # Get user-selected CSV file path
    # Set default CSV file name
    $csvFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "PA_MOVE_$timestamp.csv")

    # Initialize an array to store the results
    $results = @()

    # Get all the '_fehler' folders within the parent folder and its subfolders
    $errorFolders = Get-ChildItem -Path $parentFolder -Filter "_fehler" -Recurse -Directory

    foreach ($errorFolder in $errorFolders) {
        # Get a list of files in the error folder
        $files = Get-ChildItem -Path $errorFolder.FullName -File

        foreach ($file in $files) {
            # Check if the file is newer than the user-selected timestamp
            if ($file.LastWriteTime -gt $moveTimestamp) {
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
        "1. Move files newer than the specified timestamp from 'error' folders to their parent folder.`n" +
        "2. Check files in '_fehler' folders and export information to a CSV without moving them.`n`n" +
        "To get started, follow these steps:`n`n" +
        "1. Select the folder containing 'error' folders.`n" +
        "2. Choose the desired timestamp or select 'Custom' to specify your own date and time.`n" +
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
    $errorFolders = Get-ChildItem -Path $parentFolder -Filter "error" -Recurse -Directory

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
$mainForm.Size = New-Object System.Drawing.Size(460, 260)
$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
#$mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$mainForm.MaximizeBox = $false
$mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# Create a label and textbox for selecting the folder
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Text = "Select Folder:"
$folderLabel.Location = New-Object System.Drawing.Point(10, 40)
$mainForm.Controls.Add($folderLabel)

$folderTextBox = New-Object System.Windows.Forms.TextBox
$folderTextBox.Location = New-Object System.Drawing.Point(120, 40)
$folderTextBox.Size = New-Object System.Drawing.Size(200, 50)
$mainForm.Controls.Add($folderTextBox)

$folderBrowseButton = New-Object System.Windows.Forms.Button
$folderBrowseButton.Text = "Browse..."
$folderBrowseButton.Location = New-Object System.Drawing.Point(330, 40)
$folderBrowseButton.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $folderBrowserDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $folderTextBox.Text = $folderBrowserDialog.SelectedPath
    }
})
$mainForm.Controls.Add($folderBrowseButton)

# Create a dropdown list for selecting the time interval
$timeLabel = New-Object System.Windows.Forms.Label
$timeLabel.Text = "Select Time Interval:"
$timeLabel.Location = New-Object System.Drawing.Point(10, 70)
$mainForm.Controls.Add($timeLabel)

$timePicker = New-Object System.Windows.Forms.ComboBox
$timePicker.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$timePicker.Items.Add("5 Minutes")
$timePicker.Items.Add("10 Minutes")
$timePicker.Items.Add("15 Minutes")
$timePicker.Items.Add("20 Minutes")
$timePicker.Items.Add("30 Minutes")
$timePicker.Items.Add("Custom")
$timePicker.Location = New-Object System.Drawing.Point(150, 70)
$timePicker.Add_SelectedIndexChanged({
    if ($timePicker.SelectedItem -eq "Custom") {
        $datePicker.Enabled = $true
        $timePickerCustom.Enabled = $true
    } else {
        $datePicker.Enabled = $false
        $timePickerCustom.Enabled = $false
    }
})
$mainForm.Controls.Add($timePicker)

# Create a date picker for selecting the date (initially disabled)
$dateLabel = New-Object System.Windows.Forms.Label
$dateLabel.Text = "Select Date:"
$dateLabel.Location = New-Object System.Drawing.Point(10, 100)
$mainForm.Controls.Add($dateLabel)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(150, 100)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$datePicker.Enabled = $false
$mainForm.Controls.Add($datePicker)


# Create a label for the time picker
$timePickerLabel = New-Object System.Windows.Forms.Label
$timePickerLabel.Text = "Select Time Interval:"
$timePickerLabel.Location = New-Object System.Drawing.Point(10, 100)
$mainForm.Controls.Add($timePickerLabel)

# Create a time picker for selecting the time (initially disabled)
$timePickerCustom = New-Object System.Windows.Forms.DateTimePicker
$timePickerCustom.Location = New-Object System.Drawing.Point(150, 130)
$timePickerCustom.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
$timePickerCustom.ShowUpDown = $true
$timePickerCustom.Enabled = $false
$mainForm.Controls.Add($timePickerCustom)

# Create a label and textbox for entering the CSV file path
$csvPathLabel = New-Object System.Windows.Forms.Label
$csvPathLabel.Text = "CSV File Path:"
$csvPathLabel.Location = New-Object System.Drawing.Point(10, 160)
$mainForm.Controls.Add($csvPathLabel)
# Add this line to get the timestamp
$timestamp = Get-Date -Format "yyyy.MM.dd_HH.mm.ss" 
# Add this line to set the default CSV file path
$csvFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), "PA_MOVE_$timestamp.csv") 
$csvPathTextBox = New-Object System.Windows.Forms.TextBox
$csvPathTextBox.Location = New-Object System.Drawing.Point(150, 160)
$csvPathTextBox.Size = New-Object System.Drawing.Size(200, 50)
# Set the text of the textbox to the generated path
$csvPathTextBox.Text = $csvFilePath
$mainForm.Controls.Add($csvPathTextBox)



$csvPathBrowseButton = New-Object System.Windows.Forms.Button
$csvPathBrowseButton.Text = "Browse..."
$csvPathBrowseButton.Location = New-Object System.Drawing.Point(360, 160)
$csvPathBrowseButton.Add_Click({    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
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
$startButton.Location = New-Object System.Drawing.Point(10, 190)
$startButton.Add_Click({ Start-MoveFiles })
$mainForm.Controls.Add($startButton)

<#$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(100, 160)
$cancelButton.Add_Click({ $mainForm.Close() })
$mainForm.Controls.Add($cancelButton)#>

# Create the Help button
$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Text = "Help"
$helpButton.Location = New-Object System.Drawing.Point(10, 1)
$helpButton.Add_Click({ Show-Help })
$mainForm.Controls.Add($helpButton)

# Create the Close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Location = New-Object System.Drawing.Point(190, 190)
$closeButton.Add_Click({ $mainForm.Close() })
$mainForm.Controls.Add($closeButton)

# Create the Check button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Text = "Check"
$checkButton.Location = New-Object System.Drawing.Point(100, 190)
$checkButton.Add_Click({ Check-Files })
$mainForm.Controls.Add($checkButton)

# Show the form
$mainForm.ShowDialog() | Out-Null


