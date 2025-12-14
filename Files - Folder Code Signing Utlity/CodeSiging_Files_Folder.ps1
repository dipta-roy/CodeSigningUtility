Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- UI Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Code Signing Tool"
$form.Size = New-Object System.Drawing.Size(600, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::White

# --- Styling Helper ---
function Add-Group {
    param($text, $top, $height)
    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = $text
    $grp.Location = New-Object System.Drawing.Point(20, $top)
    $grp.Size = New-Object System.Drawing.Size(545, $height)
    $grp.ForeColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
    $form.Controls.Add($grp)
    return $grp
}

# --- Configuration Group ---
$grpConfig = Add-Group "Configuration" 10 180

$lblSignTool = New-Object System.Windows.Forms.Label
$lblSignTool.Text = "SignTool Path:"
$lblSignTool.Location = New-Object System.Drawing.Point(15, 25)
$lblSignTool.AutoSize = $true
$grpConfig.Controls.Add($lblSignTool)

$txtSignTool = New-Object System.Windows.Forms.TextBox
$txtSignTool.Location = New-Object System.Drawing.Point(15, 45)
$txtSignTool.Size = New-Object System.Drawing.Size(425, 23)
$txtSignTool.Text = "signtool.exe" 
$grpConfig.Controls.Add($txtSignTool)

$btnBrowseSignTool = New-Object System.Windows.Forms.Button
$btnBrowseSignTool.Text = "..."
$btnBrowseSignTool.Location = New-Object System.Drawing.Point(450, 44)
$btnBrowseSignTool.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowseSignTool.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtSignTool.Text = $dlg.FileName }
})
$grpConfig.Controls.Add($btnBrowseSignTool)

$lblPfx = New-Object System.Windows.Forms.Label
$lblPfx.Text = "Certificate (.pfx):"
$lblPfx.Location = New-Object System.Drawing.Point(15, 75)
$lblPfx.AutoSize = $true
$grpConfig.Controls.Add($lblPfx)

$txtPfx = New-Object System.Windows.Forms.TextBox
$txtPfx.Location = New-Object System.Drawing.Point(15, 95)
$txtPfx.Size = New-Object System.Drawing.Size(425, 23)
$grpConfig.Controls.Add($txtPfx)

$btnBrowsePfx = New-Object System.Windows.Forms.Button
$btnBrowsePfx.Text = "..."
$btnBrowsePfx.Location = New-Object System.Drawing.Point(450, 94)
$btnBrowsePfx.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowsePfx.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Certificates (*.pfx;*.p12)|*.pfx;*.p12|All Files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtPfx.Text = $dlg.FileName }
})
$grpConfig.Controls.Add($btnBrowsePfx)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.Location = New-Object System.Drawing.Point(15, 125)
$lblPass.AutoSize = $true
$grpConfig.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(15, 145)
$txtPass.Size = New-Object System.Drawing.Size(515, 23)
$txtPass.PasswordChar = "*"
$grpConfig.Controls.Add($txtPass)


# --- Target Group ---
$grpTarget = Add-Group "Target Selection" 200 100

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "File or Folder to Sign:"
$lblTarget.Location = New-Object System.Drawing.Point(15, 25)
$lblTarget.AutoSize = $true
$grpTarget.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(15, 45)
$txtTarget.Size = New-Object System.Drawing.Size(340, 23)
$grpTarget.Controls.Add($txtTarget)

$btnBrowseFile = New-Object System.Windows.Forms.Button
$btnBrowseFile.Text = "File..."
$btnBrowseFile.Location = New-Object System.Drawing.Point(365, 44)
$btnBrowseFile.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowseFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Executables (*.exe;*.dll)|*.exe;*.dll|All Files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtTarget.Text = $dlg.FileName }
})
$grpTarget.Controls.Add($btnBrowseFile)

$btnBrowseFolder = New-Object System.Windows.Forms.Button
$btnBrowseFolder.Text = "Folder..."
$btnBrowseFolder.Location = New-Object System.Drawing.Point(450, 44)
$btnBrowseFolder.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowseFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select folder containing binaries to sign"
    if ($dlg.ShowDialog() -eq "OK") { $txtTarget.Text = $dlg.SelectedPath }
})
$grpTarget.Controls.Add($btnBrowseFolder)


# --- Execution Zone ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 310)
$progressBar.Size = New-Object System.Drawing.Size(545, 25)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready to start."
$lblStatus.Location = New-Object System.Drawing.Point(20, 340)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = "DimGray"
$form.Controls.Add($lblStatus)

$btnSign = New-Object System.Windows.Forms.Button
$btnSign.Text = "SIGN FILES"
$btnSign.Location = New-Object System.Drawing.Point(20, 370)
$btnSign.Size = New-Object System.Drawing.Size(265, 50)
$btnSign.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnSign.ForeColor = [System.Drawing.Color]::White
$btnSign.FlatStyle = "Flat"
$btnSign.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnSign.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnSign)

$btnVerify = New-Object System.Windows.Forms.Button
$btnVerify.Text = "Check Signature"
$btnVerify.Location = New-Object System.Drawing.Point(300, 370)
$btnVerify.Size = New-Object System.Drawing.Size(265, 50)
$btnVerify.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$btnVerify.FlatStyle = "Flat"
$btnVerify.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$btnVerify.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnVerify)

# --- Logic ---

$btnSign.Add_Click({
    $signTool = $txtSignTool.Text
    $pfx = $txtPfx.Text
    $pass = $txtPass.Text
    $target = $txtTarget.Text

    if (-not $pfx -or -not $pass -or -not $target) {
        [System.Windows.Forms.MessageBox]::Show("Please fill all required fields.", "Missing Info", "OK", "Warning")
        return
    }

    if (-not (Test-Path $target)) {
        [System.Windows.Forms.MessageBox]::Show("Target path does not exist.", "Error", "OK", "Error")
        return
    }

    $isFolder = Test-Path $target -PathType Container
    $filesToSign = @()

    if ($isFolder) {
        $filesToSign = Get-ChildItem -Path $target -Recurse -Include *.exe,*.dll | Select-Object -ExpandProperty FullName
        if ($filesToSign.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No .exe or .dll files found in the selected folder.", "Target Empty", "OK", "Warning")
            return
        }
    } else {
        $filesToSign = @($target)
    }

    # Reset UI
    $progressBar.Value = 0
    $progressBar.Maximum = $filesToSign.Count
    $lblStatus.ForeColor = "Blue"
    $btnSign.Enabled = $false
    $btnVerify.Enabled = $false
    
    $timestampUrl = "http://timestamp.digicert.com"
    $successCount = 0
    $failedFiles = @()

    foreach ($file in $filesToSign) {
        $fileName = [System.IO.Path]::GetFileName($file)
        $lblStatus.Text = "Signing ($($successCount + $failedFiles.Count + 1)/$($filesToSign.Count)): $fileName..."
        
        # Keep UI responsive
        [System.Windows.Forms.Application]::DoEvents()

        $args = @("sign", "/f", "`"$pfx`"", "/p", "`"$pass`"", "/tr", "$timestampUrl", "/td", "sha256", "/fd", "sha256", "`"$file`"")
        
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $signTool
        $pinfo.Arguments = $args -join " "
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError = $true
        $pinfo.UseShellExecute = $false
        $pinfo.CreateNoWindow = $true
        
        try {
            $p = [System.Diagnostics.Process]::Start($pinfo)
            $p.WaitForExit()
            
            if ($p.ExitCode -eq 0) {
                $successCount++
            } else {
                $failedFiles += $file
            }
        } catch {
            $failedFiles += $file
        }
        
        $progressBar.PerformStep()
    }

    # Completion
    $btnSign.Enabled = $true
    $btnVerify.Enabled = $true
    
    if ($failedFiles.Count -eq 0) {
        $lblStatus.Text = "Done! Successfully signed all $successCount file(s)."
        $lblStatus.ForeColor = "Green"
        [System.Windows.Forms.MessageBox]::Show("All files signed successfully!", "Success", "OK", "Information")
    } else {
        $lblStatus.Text = "Completed with errors."
        $lblStatus.ForeColor = "Red"
        $failedList = $failedFiles -join "`n"
        [System.Windows.Forms.MessageBox]::Show("Signed: $successCount`nFailed: $($failedFiles.Count)`n`nFailed files:`n$failedList", "Completed with Errors", "OK", "Warning")
    }
})

$btnVerify.Add_Click({
    $signTool = $txtSignTool.Text
    $target = $txtTarget.Text
    
    if (-not $target) { return }

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $signTool
    $pinfo.Arguments = "verify /pa `"$target`""
    $pinfo.RedirectStandardOutput = $true
    $pinfo.RedirectStandardError = $true
    $pinfo.UseShellExecute = $false
    $pinfo.CreateNoWindow = $true

    $lblStatus.Text = "Verifying signature..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $p = [System.Diagnostics.Process]::Start($pinfo)
        $p.WaitForExit()
        $output = $p.StandardOutput.ReadToEnd()
        
        if ($p.ExitCode -eq 0) {
             [System.Windows.Forms.MessageBox]::Show("File is SIGNED.`n`nOutput: $output", "Verified", "OK", "Information")
             $lblStatus.Text = "Verification check passed."
             $lblStatus.ForeColor = "Green"
        } else {
             [System.Windows.Forms.MessageBox]::Show("File is NOT signed (or verification failed).`n`nOutput: $output", "Not Signed", "OK", "Warning")
             $lblStatus.Text = "Verification check failed."
             $lblStatus.ForeColor = "Red"
        }
    } catch {
         [System.Windows.Forms.MessageBox]::Show("Could not run signtool.", "Error", "OK", "Error")
         $lblStatus.Text = "Verification error."
    }
})

$form.ShowDialog()
