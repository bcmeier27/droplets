# Collect dropped files/folders from script arguments

##############
# $Paths will be set if files or folders were dropped onto this script
param(
    [Parameter(ValueFromRemainingArguments=$true)]
    [string[]]$Paths
)


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


##############

# Expand folders to individual .doc/.docx files
function Get-WordFiles {
    param([string[]]$Inputs)
    $files = @()
    foreach ($p in $Inputs) {
        if (Test-Path $p) {
            if ((Get-Item $p).PSIsContainer) {
                $files += Get-ChildItem -Path $p -Include *.doc,*.docx,*.docm,*.rtf -File -Recurse
            } else {
                if ($p.ToLower().EndsWith('.doc') -or $p.ToLower().EndsWith('.docx') -or $p.ToLower().EndsWith('.rtf') -or $p.ToLower().EndsWith('.docm')) {
                    $files += Get-Item -LiteralPath $p
                }
            }
        }
    }
    return $files
}

$wordFiles = Get-WordFiles -Inputs $Paths

##############
# This test will cause a file selector dialog to open if nothing was passed to the script, either dropped or command line args
if (-not $wordFiles) {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = "C:\"
    $OpenFileDialog.Filter = "All files (*.*)|*.*"
    $Result = $OpenFileDialog.ShowDialog()
    if ($Result -eq "OK") {
        $SelectedFile = $OpenFileDialog.FileName
        Write-Host "Selected file: $SelectedFile"
        $wordFiles = Get-WordFiles -Inputs $SelectedFile
    } else {
        # [System.Windows.Forms.MessageBox]::Show('No Word documents found to process.', 'Error', 'OK', 'Error') | Out-Null
        exit 1
    }
}

##############
# The dialog needs the count of files to process for the OK button
$itemCount = $wordFiles.Count

# Form setup
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Word Link Droplet'
$form.Size = New-Object System.Drawing.Size(500, 320)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Font
$font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Font = $font

# Label + TextBox helper function
function Add-LabelAndTextbox {
    param (
        [string]$labelText,
        [string]$defaultText,
        [int]$top
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(20, $top)
    $label.Size = New-Object System.Drawing.Size(80, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Text = $defaultText
    $textbox.Location = New-Object System.Drawing.Point(110, ($top - 2))
    $textbox.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($textbox)

    return $textbox
}

# StringA
$txtA = Add-LabelAndTextbox -labelText 'StringA:' -defaultText 'NASB' -top 20

# StringB
$txtB = Add-LabelAndTextbox -labelText 'StringB:' -defaultText 'SCH2000' -top 60

# Checkboxes group
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = 'Options'
$groupBox.Location = New-Object System.Drawing.Point(20, 100)
$groupBox.Size = New-Object System.Drawing.Size(430, 70)
$form.Controls.Add($groupBox)

$chkInsert = New-Object System.Windows.Forms.CheckBox
$chkInsert.Text = 'Insert text from links'
$chkInsert.Checked = $true
$chkInsert.Location = New-Object System.Drawing.Point(10, 20)
$chkInsert.Size = New-Object System.Drawing.Size(150, 20)
$groupBox.Controls.Add($chkInsert)

$chkScreenTip = New-Object System.Windows.Forms.CheckBox
$chkScreenTip.Text = 'Set screen tip'
$chkScreenTip.Checked = $true
$chkScreenTip.Location = New-Object System.Drawing.Point(200, 20)
$chkScreenTip.Size = New-Object System.Drawing.Size(150, 20)
$groupBox.Controls.Add($chkScreenTip)

$chkOpen = New-Object System.Windows.Forms.CheckBox
$chkOpen.Text = 'Open after processing'
$chkOpen.Checked = $true
$chkOpen.Location = New-Object System.Drawing.Point(10, 45)
$chkOpen.Size = New-Object System.Drawing.Size(150, 20)
$groupBox.Controls.Add($chkOpen)

# Folder picker
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = 'Save location:'
$lblFolder.Location = New-Object System.Drawing.Point(20, 185)
$lblFolder.AutoSize = $true
$form.Controls.Add($lblFolder)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(110, 182)
$txtFolder.Size = New-Object System.Drawing.Size(220, 20)
$txtFolder.ReadOnly = $true
$form.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(340, 180)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = [Environment]::GetFolderPath('MyDocuments')
    if ($dlg.ShowDialog() -eq 'OK') { $txtFolder.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnBrowse)

# Conditional logic for enabling/disabling
if ($itemCount -gt 1) {
    $chkOpen.Checked = $false
    $chkOpen.Enabled = $false
}
$btnBrowse.Enabled = -not $chkOpen.Checked
$chkOpen.Add_CheckedChanged({ $btnBrowse.Enabled = -not $chkOpen.Checked })

# Action buttons
$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = "Process $itemCount file" + ($(if ($itemCount -ne 1) { 's' } else { '' }))
$btnProcess.Location = New-Object System.Drawing.Point(250, 230)
$btnProcess.Size = New-Object System.Drawing.Size(100, 30)
$btnProcess.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })
$form.Controls.Add($btnProcess)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(120, 230)
$btnCancel.Size = New-Object System.Drawing.Size(100, 30)
$btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })
$form.Controls.Add($btnCancel)

# Show the form and EXIT if needed
$result = $form.ShowDialog()
if ($result -ne [System.Windows.Forms.DialogResult]::OK) { 
    exit 0 
}

##############
# Initialize Word COM
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

$optionsProcessed = @()

# Gather options into a record
$options = [PSCustomObject]@{
    StringA      = $txtA.Text
    StringB      = $txtB.Text
    InsertText   = $chkInsert.Checked
    OpenAfter    = $chkOpen.Checked
    SetScreenTip = $chkScreenTip.Checked
    SaveLocation = if (-not $chkOpen.Checked) { $txtFolder.Text } else { $null }
}

Write-Host $options

##############
# Start of actual processing here
#
foreach ($file in $wordFiles) {
    $doc = $word.Documents.Open($file.FullName)

    # Get all the hyperlinks
    foreach ($h in $doc.Hyperlinks) {

        ##############
        # Step 1. Replace StringA with StringB within the URL of each matching hyperlink
        $oldUrl = $h.Address
        $newUrl = $oldUrl
        if ($options.StringA -ne "") {
            $newUrl = $oldUrl -replace [regex]::Escape($options.StringA), $options.StringB
            $h.Address = $newUrl
        }

        ##############
        # Step 2. Fetch and insert description
        $desc = ""
        if ($options.InsertText -or $options.SetScreenTip) {
            try {
                $html = Invoke-WebRequest -Uri $newUrl -UseBasicParsing
                if ($html.RawContent -match '<meta[^>]*property="og:description"[^>]*content="([^"]*)') {
                    $desc = $Matches[1]
                }
                if ($options.InsertText) {
                    $r = $h.Range
                    $newRange = $r.Duplicate
                    $newRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
                    $newRange.InsertAfter(" - ""$desc"" ")
                    $newRange.Font.Italic = $true
                }
                if ($options.SetScreenTip) {
                    $h.ScreenTip = $desc
                }
            } catch { }
        }
    }

    ##############
    # Step 3. Leave open or Save 
    if ($options.OpenAfter) {
    #         Leave the document open after processing (can only be allowed for single files) 

        ## DO NOT CALL $doc.Save() AT ALL
        $optionsProcessed += $file.FullName
    } else {
    #         -OR- 
    #         Save as a new file with PROCESSED- prefix

        # SaveLocation might be empty, indicating to leave files in the same directory as original
        if ($options.SaveLocation) { 
            $saveName = Join-Path $options.SaveLocation ("PROCESSED-" + $file.Name)
        } else { 
            $saveName = Join-Path $file.DirectoryName ("PROCESSED-" + $file.Name) 
        }
        if (!(Test-Path -Path (Split-Path $saveName))) {
            New-Item -Path (Split-Path $saveName) -ItemType Directory | Out-Null
        }
        
        # Write-Host "Saving file to new location: $saveName"
        $wdFormatDocumentDefault = 16
        $doc.SaveAs([ref][System.Object]$saveName) # stupidly, the SaveAs method does NOT automatically cast the file name from a string to an object
        $optionsProcessed += $saveName
        $doc.Close()
    }
}

# Finally, keep Word open if requested AND we're only processing a single file, otherwise quit the app
if ($options.OpenAfter) {
    $word.Visible = $true
    # Write-Host ("Word remains open after processing " + $file.Name)
} else {
    $word.Quit()
    # Write-Host ("Word closed after processing " + $file.Name)
    [System.Windows.Forms.MessageBox]::Show("Processed $($optionsProcessed.Count) file(s) saved ... $($options.SaveLocation)", 'Done', 'OK', 'Information') | Out-Null
}

## End of WordLinkDroplet ##
