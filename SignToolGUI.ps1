#Requires -RunAsAdministrator

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form with modern styling
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Code Signing Certificate Manager'
$form.Size = New-Object System.Drawing.Size(750, 820)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Create ToolTip object for interactive help
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 1000
$toolTip.ReshowDelay = 500

# ====================================
# SECTION 1: Generate Code Signing Certificate
# ====================================

$groupBoxGenerate = New-Object System.Windows.Forms.GroupBox
$groupBoxGenerate.Location = New-Object System.Drawing.Point(20, 20)
$groupBoxGenerate.Size = New-Object System.Drawing.Size(700, 300)
$groupBoxGenerate.Text = ' Generate Code Signing Certificate '
$groupBoxGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$groupBoxGenerate.ForeColor = [System.Drawing.Color]::FromArgb(30, 70, 130)
$groupBoxGenerate.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($groupBoxGenerate)

# Certificate Destination Path
$labelDestPath = New-Object System.Windows.Forms.Label
$labelDestPath.Location = New-Object System.Drawing.Point(20, 35)
$labelDestPath.Size = New-Object System.Drawing.Size(250, 20)
$labelDestPath.Text = 'Certificate Destination Path:'
$labelDestPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelDestPath.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxGenerate.Controls.Add($labelDestPath)

$textBoxDestPath = New-Object System.Windows.Forms.TextBox
$textBoxDestPath.Location = New-Object System.Drawing.Point(20, 60)
$textBoxDestPath.Size = New-Object System.Drawing.Size(530, 25)
$textBoxDestPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxDestPath.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxDestPath, "Select the folder where your certificate files will be saved")
$groupBoxGenerate.Controls.Add($textBoxDestPath)

$btnBrowseDestPath = New-Object System.Windows.Forms.Button
$btnBrowseDestPath.Location = New-Object System.Drawing.Point(560, 58)
$btnBrowseDestPath.Size = New-Object System.Drawing.Size(110, 29)
$btnBrowseDestPath.Text = 'Browse...'
$btnBrowseDestPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnBrowseDestPath.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$btnBrowseDestPath.FlatStyle = 'Flat'
$btnBrowseDestPath.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnBrowseDestPath.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnBrowseDestPath, "Browse for destination folder")
$groupBoxGenerate.Controls.Add($btnBrowseDestPath)

# Certificate CN Name
$labelCNName = New-Object System.Windows.Forms.Label
$labelCNName.Location = New-Object System.Drawing.Point(20, 100)
$labelCNName.Size = New-Object System.Drawing.Size(300, 20)
$labelCNName.Text = 'Certificate Subject (CN Name):'
$labelCNName.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelCNName.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxGenerate.Controls.Add($labelCNName)

$textBoxCNName = New-Object System.Windows.Forms.TextBox
$textBoxCNName.Location = New-Object System.Drawing.Point(20, 125)
$textBoxCNName.Size = New-Object System.Drawing.Size(650, 25)
$textBoxCNName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxCNName.Text = "My Code Signing Certificate"
$textBoxCNName.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxCNName, "Enter a unique name for your certificate (e.g., 'MyCompany Code Signing')")
$groupBoxGenerate.Controls.Add($textBoxCNName)

# Certificate Password
$labelCertPassword = New-Object System.Windows.Forms.Label
$labelCertPassword.Location = New-Object System.Drawing.Point(20, 165)
$labelCertPassword.Size = New-Object System.Drawing.Size(200, 20)
$labelCertPassword.Text = 'Certificate Password:'
$labelCertPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelCertPassword.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxGenerate.Controls.Add($labelCertPassword)

$textBoxCertPassword = New-Object System.Windows.Forms.TextBox
$textBoxCertPassword.Location = New-Object System.Drawing.Point(20, 190)
$textBoxCertPassword.Size = New-Object System.Drawing.Size(650, 25)
$textBoxCertPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxCertPassword.UseSystemPasswordChar = $true
$textBoxCertPassword.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxCertPassword, "Enter a strong password to protect your private key")
$groupBoxGenerate.Controls.Add($textBoxCertPassword)

# Status Label for Certificate Generation
$labelGenStatus = New-Object System.Windows.Forms.Label
$labelGenStatus.Location = New-Object System.Drawing.Point(20, 225)
$labelGenStatus.Size = New-Object System.Drawing.Size(650, 20)
$labelGenStatus.Text = 'Status: Ready'
$labelGenStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$labelGenStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$groupBoxGenerate.Controls.Add($labelGenStatus)

# Generate Button with gradient effect
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Location = New-Object System.Drawing.Point(20, 250)
$btnGenerate.Size = New-Object System.Drawing.Size(650, 40)
$btnGenerate.Text = 'Generate Certificate'
$btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = 'Flat'
$btnGenerate.FlatAppearance.BorderSize = 0
$btnGenerate.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 105, 190)
$btnGenerate.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 90, 170)
$btnGenerate.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnGenerate, "Click to generate a new code signing certificate")
$groupBoxGenerate.Controls.Add($btnGenerate)

# ====================================
# SECTION 2: Sign Binary
# ====================================

$groupBoxSign = New-Object System.Windows.Forms.GroupBox
$groupBoxSign.Location = New-Object System.Drawing.Point(20, 340)
$groupBoxSign.Size = New-Object System.Drawing.Size(700, 440)
$groupBoxSign.Text = ' Sign Binary '
$groupBoxSign.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$groupBoxSign.ForeColor = [System.Drawing.Color]::FromArgb(30, 100, 50)
$groupBoxSign.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($groupBoxSign)

# Binary Path
$labelBinaryPath = New-Object System.Windows.Forms.Label
$labelBinaryPath.Location = New-Object System.Drawing.Point(20, 35)
$labelBinaryPath.Size = New-Object System.Drawing.Size(200, 20)
$labelBinaryPath.Text = 'Binary File to Sign:'
$labelBinaryPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelBinaryPath.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxSign.Controls.Add($labelBinaryPath)

$textBoxBinaryPath = New-Object System.Windows.Forms.TextBox
$textBoxBinaryPath.Location = New-Object System.Drawing.Point(20, 60)
$textBoxBinaryPath.Size = New-Object System.Drawing.Size(530, 25)
$textBoxBinaryPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxBinaryPath.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxBinaryPath, "Select the executable file (.exe, .dll, .msi) to sign")
$groupBoxSign.Controls.Add($textBoxBinaryPath)

$btnBrowseBinary = New-Object System.Windows.Forms.Button
$btnBrowseBinary.Location = New-Object System.Drawing.Point(560, 58)
$btnBrowseBinary.Size = New-Object System.Drawing.Size(110, 29)
$btnBrowseBinary.Text = 'Browse...'
$btnBrowseBinary.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnBrowseBinary.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$btnBrowseBinary.FlatStyle = 'Flat'
$btnBrowseBinary.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnBrowseBinary.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnBrowseBinary, "Browse for file to sign")
$groupBoxSign.Controls.Add($btnBrowseBinary)

# SignTool Path
$labelSignToolPath = New-Object System.Windows.Forms.Label
$labelSignToolPath.Location = New-Object System.Drawing.Point(20, 95)
$labelSignToolPath.Size = New-Object System.Drawing.Size(200, 20)
$labelSignToolPath.Text = 'SignTool.exe Path:'
$labelSignToolPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelSignToolPath.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxSign.Controls.Add($labelSignToolPath)

$textBoxSignToolPath = New-Object System.Windows.Forms.TextBox
$textBoxSignToolPath.Location = New-Object System.Drawing.Point(20, 120)
$textBoxSignToolPath.Size = New-Object System.Drawing.Size(530, 25)
$textBoxSignToolPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxSignToolPath.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxSignToolPath, "Path to Microsoft SignTool.exe (usually in Windows SDK)")
$groupBoxSign.Controls.Add($textBoxSignToolPath)

$btnBrowseSignTool = New-Object System.Windows.Forms.Button
$btnBrowseSignTool.Location = New-Object System.Drawing.Point(560, 118)
$btnBrowseSignTool.Size = New-Object System.Drawing.Size(110, 29)
$btnBrowseSignTool.Text = 'Browse...'
$btnBrowseSignTool.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnBrowseSignTool.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$btnBrowseSignTool.FlatStyle = 'Flat'
$btnBrowseSignTool.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnBrowseSignTool.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnBrowseSignTool, "Browse for SignTool.exe")
$groupBoxSign.Controls.Add($btnBrowseSignTool)

# Certificate Path
$labelCertPath = New-Object System.Windows.Forms.Label
$labelCertPath.Location = New-Object System.Drawing.Point(20, 155)
$labelCertPath.Size = New-Object System.Drawing.Size(200, 20)
$labelCertPath.Text = 'Certificate (.pfx) Path:'
$labelCertPath.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelCertPath.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxSign.Controls.Add($labelCertPath)

$textBoxCertPath = New-Object System.Windows.Forms.TextBox
$textBoxCertPath.Location = New-Object System.Drawing.Point(20, 180)
$textBoxCertPath.Size = New-Object System.Drawing.Size(530, 25)
$textBoxCertPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxCertPath.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxCertPath, "Select your .pfx certificate file")
$groupBoxSign.Controls.Add($textBoxCertPath)

$btnBrowseCert = New-Object System.Windows.Forms.Button
$btnBrowseCert.Location = New-Object System.Drawing.Point(560, 178)
$btnBrowseCert.Size = New-Object System.Drawing.Size(110, 29)
$btnBrowseCert.Text = 'Browse...'
$btnBrowseCert.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnBrowseCert.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$btnBrowseCert.FlatStyle = 'Flat'
$btnBrowseCert.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$btnBrowseCert.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnBrowseCert, "Browse for certificate file")
$groupBoxSign.Controls.Add($btnBrowseCert)

# Certificate Password for Signing
$labelSignPassword = New-Object System.Windows.Forms.Label
$labelSignPassword.Location = New-Object System.Drawing.Point(20, 215)
$labelSignPassword.Size = New-Object System.Drawing.Size(200, 20)
$labelSignPassword.Text = 'Certificate Password:'
$labelSignPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelSignPassword.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxSign.Controls.Add($labelSignPassword)

$textBoxSignPassword = New-Object System.Windows.Forms.TextBox
$textBoxSignPassword.Location = New-Object System.Drawing.Point(20, 240)
$textBoxSignPassword.Size = New-Object System.Drawing.Size(650, 25)
$textBoxSignPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxSignPassword.UseSystemPasswordChar = $true
$textBoxSignPassword.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxSignPassword, "Enter the password for your certificate")
$groupBoxSign.Controls.Add($textBoxSignPassword)

# Timestamp URL
$labelTimestamp = New-Object System.Windows.Forms.Label
$labelTimestamp.Location = New-Object System.Drawing.Point(20, 275)
$labelTimestamp.Size = New-Object System.Drawing.Size(200, 20)
$labelTimestamp.Text = 'Timestamp URL:'
$labelTimestamp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$labelTimestamp.ForeColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$groupBoxSign.Controls.Add($labelTimestamp)

$textBoxTimestamp = New-Object System.Windows.Forms.TextBox
$textBoxTimestamp.Location = New-Object System.Drawing.Point(20, 300)
$textBoxTimestamp.Size = New-Object System.Drawing.Size(650, 25)
$textBoxTimestamp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxTimestamp.Text = "http://timestamp.sectigo.com"
$textBoxTimestamp.BorderStyle = 'FixedSingle'
$toolTip.SetToolTip($textBoxTimestamp, "Timestamp server URL (ensures signature validity after certificate expiry)")
$groupBoxSign.Controls.Add($textBoxTimestamp)

# *** NEW FEATURE: Checkbox for saving public key to destination ***
$checkBoxSavePublicKey = New-Object System.Windows.Forms.CheckBox
$checkBoxSavePublicKey.Location = New-Object System.Drawing.Point(20, 335)
$checkBoxSavePublicKey.Size = New-Object System.Drawing.Size(650, 25)
$checkBoxSavePublicKey.Text = 'Copy public key (.cer) to binary destination after signing'
$checkBoxSavePublicKey.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$checkBoxSavePublicKey.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
$checkBoxSavePublicKey.Checked = $true
$checkBoxSavePublicKey.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($checkBoxSavePublicKey, "When enabled, the public key will be automatically copied to the same directory as the signed binary")
$groupBoxSign.Controls.Add($checkBoxSavePublicKey)

# Status Label for Signing
$labelSignStatus = New-Object System.Windows.Forms.Label
$labelSignStatus.Location = New-Object System.Drawing.Point(20, 365)
$labelSignStatus.Size = New-Object System.Drawing.Size(650, 20)
$labelSignStatus.Text = 'Status: Ready'
$labelSignStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$groupBoxSign.Controls.Add($labelSignStatus)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 365)
$progressBar.Size = New-Object System.Drawing.Size(650, 8)
$progressBar.Style = 'Continuous'
$progressBar.Visible = $false
$groupBoxSign.Controls.Add($progressBar)

# Sign Button
$btnSign = New-Object System.Windows.Forms.Button
$btnSign.Location = New-Object System.Drawing.Point(20, 385)
$btnSign.Size = New-Object System.Drawing.Size(650, 40)
$btnSign.Text = 'Sign Binary'
$btnSign.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnSign.BackColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
$btnSign.ForeColor = [System.Drawing.Color]::White
$btnSign.FlatStyle = 'Flat'
$btnSign.FlatAppearance.BorderSize = 0
$btnSign.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(14, 120, 54)
$btnSign.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(12, 100, 46)
$btnSign.Cursor = [System.Windows.Forms.Cursors]::Hand
$toolTip.SetToolTip($btnSign, "Click to sign the selected binary file")
$groupBoxSign.Controls.Add($btnSign)

# ====================================
# EVENT HANDLERS
# ====================================

# Browse Destination Path
$btnBrowseDestPath.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select destination folder for certificate"
    $folderBrowser.RootFolder = 'MyComputer'
    
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $textBoxDestPath.Text = $folderBrowser.SelectedPath
        $labelGenStatus.Text = "Status: Destination path selected"
        $labelGenStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    }
})

# Browse Binary File
$btnBrowseBinary.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Executable Files (*.exe;*.dll;*.msi)|*.exe;*.dll;*.msi|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Binary File to Sign"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxBinaryPath.Text = $openFileDialog.FileName
        $labelSignStatus.Text = "Status: Binary file selected - ready to sign"
        $labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
    }
})

# Browse SignTool
$btnBrowseSignTool.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "SignTool (signtool.exe)|signtool.exe|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select SignTool.exe"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxSignToolPath.Text = $openFileDialog.FileName
        $labelSignStatus.Text = "Status: SignTool.exe located"
        $labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
    }
})

# Browse Certificate
$btnBrowseCert.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Certificate Files (*.pfx;*.p12)|*.pfx;*.p12|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Certificate File"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxCertPath.Text = $openFileDialog.FileName
        $labelSignStatus.Text = "Status: Certificate selected"
        $labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
    }
})

# Checkbox state changed event
$checkBoxSavePublicKey.Add_CheckedChanged({
    if ($checkBoxSavePublicKey.Checked) {
        $labelSignStatus.Text = "Status: Public key will be copied to binary destination"
        $labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    } else {
        $labelSignStatus.Text = "Status: Public key will NOT be copied to binary destination"
        $labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 100, 0)
    }
})

# Generate Certificate Button
$btnGenerate.Add_Click({
    try {
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($textBoxDestPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a destination path.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxCNName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a certificate CN name.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxCertPassword.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a certificate password.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path $textBoxDestPath.Text)) {
            New-Item -Path $textBoxDestPath.Text -ItemType Directory -Force | Out-Null
        }
        
        $certName = $textBoxCNName.Text
        $cleanCertName = $certName -replace ' ', '_'
        $pfxFileName = "${cleanCertName}_CodeSigningPrivateKey.pfx"
        $certPath = Join-Path $textBoxDestPath.Text $pfxFileName
        $password = $textBoxCertPassword.Text
        $securePassword = ConvertTo-SecureString -String $password -Force -AsPlainText
        
        # Update UI for operation in progress
        $btnGenerate.Enabled = $false
        $btnGenerate.Text = "Generating Certificate..."
        $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
        $labelGenStatus.Text = "Status: Generating certificate... Please wait"
        $labelGenStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 100, 0)
        $form.Refresh()
        
        # Create self-signed certificate
        $cert = New-SelfSignedCertificate -Type CodeSigningCert `
            -Subject "CN=$certName" `
            -KeyAlgorithm RSA `
            -KeyLength 2048 `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -NotAfter (Get-Date).AddYears(5) `
            -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
        
        # Export certificate to PFX file
        Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $securePassword | Out-Null
        
        # Export public key
        $cerFileName = "${cleanCertName}_CodeSigningPublicKey.cer"
        $cerPath = Join-Path $textBoxDestPath.Text $cerFileName
        Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null

        # Create Key Details.txt file
        $keyDetailsFileName = "${cleanCertName}_KeyDetails.txt"
        $keyDetailsPath = Join-Path $textBoxDestPath.Text $keyDetailsFileName
        $keyDetailsContent = @"
Date of Key Generation: $(Get-Date)
Subject Name: $certName
Private Key Name: $pfxFileName
Public Key Name: $cerFileName
Password: $password
Hash Algorithm: SHA256
Valid Until: $($cert.NotAfter)
Thumbprint: $($cert.Thumbprint)
"@
        Set-Content -Path $keyDetailsPath -Value $keyDetailsContent
        
        # Reset button state
        $btnGenerate.Enabled = $true
        $btnGenerate.Text = "Generate Certificate"
        $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $labelGenStatus.Text = "Status: Certificate generated successfully!"
        $labelGenStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
        
        $successMsg = "Certificate generated successfully!`n`nLocation: $certPath`nThumbprint: $($cert.Thumbprint)`nFiles created:`n   - ${pfxFileName}`n   - ${cerFileName}`n   - ${keyDetailsFileName}"
        [System.Windows.Forms.MessageBox]::Show($successMsg, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Auto-populate the certificate path in the signing section
        $textBoxCertPath.Text = $certPath
        $textBoxSignPassword.Text = $password

    }
    catch {
        $btnGenerate.Enabled = $true
        $btnGenerate.Text = "Generate Certificate"
        $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
        $labelGenStatus.Text = "Status: Error during certificate generation"
        $labelGenStatus.ForeColor = [System.Drawing.Color]::Red
        
        $errorMsg = "Error generating certificate:`n`n$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Sign Binary Button
$btnSign.Add_Click({
    try {
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($textBoxBinaryPath.Text) -or -not (Test-Path $textBoxBinaryPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a valid binary file path.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxSignToolPath.Text) -or -not (Test-Path $textBoxSignToolPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a valid SignTool.exe path.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxCertPath.Text) -or -not (Test-Path $textBoxCertPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a valid certificate file path.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxSignPassword.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify the certificate password.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Show progress bar
        $progressBar.Value = 0
        $progressBar.Visible = $true
        $labelSignStatus.Visible = $false
        $btnSign.Enabled = $false
        $btnSign.Text = "Signing in progress..."
        $btnSign.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
        $form.Refresh()
        
        $binaryPath = $textBoxBinaryPath.Text
        $signToolPath = $textBoxSignToolPath.Text
        $certPath = $textBoxCertPath.Text
        $password = $textBoxSignPassword.Text
        $timestampUrl = $textBoxTimestamp.Text
        
        $progressBar.Value = 20
        $form.Refresh()
        
        # Build SignTool command
        $arguments = @(
            "sign",
            "/f", "`"$certPath`"",
            "/p", "`"$password`"",
            "/fd", "SHA256",
            "/tr", "`"$timestampUrl`"",
            "/td", "SHA256",
            "`"$binaryPath`""
        )
        
        $progressBar.Value = 40
        $form.Refresh()
        
        # Execute SignTool
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $signToolPath
        $processInfo.Arguments = $arguments -join " "
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $process.Start() | Out-Null
        
        $progressBar.Value = 60
        $form.Refresh()
        
        $output = $process.StandardOutput.ReadToEnd()
        $errorOutput = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        
        $progressBar.Value = 80
        $form.Refresh()
        
        if ($process.ExitCode -eq 0) {
            $progressBar.Value = 100
            $form.Refresh()
            Start-Sleep -Milliseconds 300
            
            # Check if user wants to copy public key to destination
            $publicKeyCopied = $false
            if ($checkBoxSavePublicKey.Checked) {
                # Copy public key to binary path
                $pfxFileName = Split-Path -Path $certPath -Leaf
                $cleanCertName = $pfxFileName -replace '_CodeSigningPrivateKey.pfx', ''
                $cerFileName = "${cleanCertName}_CodeSigningPublicKey.cer"
                $sourceCerPath = Join-Path (Split-Path -Path $certPath -Parent) $cerFileName
                $destinationCerPath = Join-Path (Split-Path -Path $binaryPath -Parent) $cerFileName

                if (Test-Path $sourceCerPath) {
                    Copy-Item -Path $sourceCerPath -Destination $destinationCerPath -Force
                    $publicKeyCopied = $true
                }
            }
            
            # Show success message
            if ($publicKeyCopied) {
                $successMsg = "Binary signed successfully!`n`nFile: $binaryPath`n`nPublic key copied to:`n   $destinationCerPath`n`nOutput:`n$output"
                [System.Windows.Forms.MessageBox]::Show($successMsg, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                if ($checkBoxSavePublicKey.Checked) {
                    $reasonMsg = "`n`nWarning: Public key was not found at expected location."
                } else {
                    $reasonMsg = "`n`nNote: Public key was not copied (option disabled)."
                }
                
                $successMsg = "Binary signed successfully!`n`nFile: $binaryPath$reasonMsg`n`nOutput:`n$output"
                [System.Windows.Forms.MessageBox]::Show($successMsg, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            
            $labelSignStatus.Text = "Status: Binary signed successfully!"
            $labelSignStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
        }
        else {
            throw "SignTool failed with exit code $($process.ExitCode)`n`nError: $errorOutput`n`nOutput: $output"
        }
        
    }
    catch {
        $labelSignStatus.Text = "Status: Error during signing"
        $labelSignStatus.ForeColor = [System.Drawing.Color]::Red
        
        $errorMsg = "Error signing binary:`n`n$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $labelSignStatus.Visible = $true
        $btnSign.Enabled = $true
        $btnSign.Text = "Sign Binary"
        $btnSign.BackColor = [System.Drawing.Color]::FromArgb(16, 137, 62)
    }
})

# ====================================
# SHOW FORM
# ====================================

# Show the form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()