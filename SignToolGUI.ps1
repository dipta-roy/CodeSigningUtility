#Requires -RunAsAdministrator

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Code Signing Certificate Manager'
$form.Size = New-Object System.Drawing.Size(700, 750)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# ====================================
# SECTION 1: Generate Code Signing Certificate
# ====================================

$groupBoxGenerate = New-Object System.Windows.Forms.GroupBox
$groupBoxGenerate.Location = New-Object System.Drawing.Point(20, 20)
$groupBoxGenerate.Size = New-Object System.Drawing.Size(640, 280)
$groupBoxGenerate.Text = 'Generate Code Signing Certificate'
$groupBoxGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($groupBoxGenerate)

# Certificate Destination Path
$labelDestPath = New-Object System.Windows.Forms.Label
$labelDestPath.Location = New-Object System.Drawing.Point(20, 35)
$labelDestPath.Size = New-Object System.Drawing.Size(200, 20)
$labelDestPath.Text = 'Destination Path:'
$labelDestPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxGenerate.Controls.Add($labelDestPath)

$textBoxDestPath = New-Object System.Windows.Forms.TextBox
$textBoxDestPath.Location = New-Object System.Drawing.Point(20, 60)
$textBoxDestPath.Size = New-Object System.Drawing.Size(480, 25)
$textBoxDestPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxGenerate.Controls.Add($textBoxDestPath)

$btnBrowseDestPath = New-Object System.Windows.Forms.Button
$btnBrowseDestPath.Location = New-Object System.Drawing.Point(510, 58)
$btnBrowseDestPath.Size = New-Object System.Drawing.Size(100, 27)
$btnBrowseDestPath.Text = 'Browse...'
$btnBrowseDestPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxGenerate.Controls.Add($btnBrowseDestPath)

# Certificate CN Name
$labelCNName = New-Object System.Windows.Forms.Label
$labelCNName.Location = New-Object System.Drawing.Point(20, 100)
$labelCNName.Size = New-Object System.Drawing.Size(250, 20)
$labelCNName.Text = 'Certificate Subject (CN Name):'
$labelCNName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxGenerate.Controls.Add($labelCNName)

$textBoxCNName = New-Object System.Windows.Forms.TextBox
$textBoxCNName.Location = New-Object System.Drawing.Point(20, 125)
$textBoxCNName.Size = New-Object System.Drawing.Size(590, 25)
$textBoxCNName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxCNName.Text = "My Code Signing Certificate"
$groupBoxGenerate.Controls.Add($textBoxCNName)

# Certificate Password
$labelCertPassword = New-Object System.Windows.Forms.Label
$labelCertPassword.Location = New-Object System.Drawing.Point(20, 165)
$labelCertPassword.Size = New-Object System.Drawing.Size(200, 20)
$labelCertPassword.Text = 'Certificate Password:'
$labelCertPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxGenerate.Controls.Add($labelCertPassword)

$textBoxCertPassword = New-Object System.Windows.Forms.TextBox
$textBoxCertPassword.Location = New-Object System.Drawing.Point(20, 190)
$textBoxCertPassword.Size = New-Object System.Drawing.Size(590, 25)
$textBoxCertPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxCertPassword.UseSystemPasswordChar = $true
$groupBoxGenerate.Controls.Add($textBoxCertPassword)

# Generate Button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Location = New-Object System.Drawing.Point(20, 230)
$btnGenerate.Size = New-Object System.Drawing.Size(590, 35)
$btnGenerate.Text = 'Generate Certificate'
$btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = 'Flat'
$btnGenerate.Cursor = [System.Windows.Forms.Cursors]::Hand
$groupBoxGenerate.Controls.Add($btnGenerate)

# ====================================
# SECTION 2: Sign Binary
# ====================================

$groupBoxSign = New-Object System.Windows.Forms.GroupBox
$groupBoxSign.Location = New-Object System.Drawing.Point(20, 320)
$groupBoxSign.Size = New-Object System.Drawing.Size(640, 380)
$groupBoxSign.Text = 'Sign Binary'
$groupBoxSign.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($groupBoxSign)

# Binary Path
$labelBinaryPath = New-Object System.Windows.Forms.Label
$labelBinaryPath.Location = New-Object System.Drawing.Point(20, 35)
$labelBinaryPath.Size = New-Object System.Drawing.Size(200, 20)
$labelBinaryPath.Text = 'Binary File Path:'
$labelBinaryPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($labelBinaryPath)

$textBoxBinaryPath = New-Object System.Windows.Forms.TextBox
$textBoxBinaryPath.Location = New-Object System.Drawing.Point(20, 60)
$textBoxBinaryPath.Size = New-Object System.Drawing.Size(480, 25)
$textBoxBinaryPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($textBoxBinaryPath)

$btnBrowseBinary = New-Object System.Windows.Forms.Button
$btnBrowseBinary.Location = New-Object System.Drawing.Point(510, 58)
$btnBrowseBinary.Size = New-Object System.Drawing.Size(100, 27)
$btnBrowseBinary.Text = 'Browse...'
$btnBrowseBinary.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($btnBrowseBinary)

# SignTool Path
$labelSignToolPath = New-Object System.Windows.Forms.Label
$labelSignToolPath.Location = New-Object System.Drawing.Point(20, 95)
$labelSignToolPath.Size = New-Object System.Drawing.Size(200, 20)
$labelSignToolPath.Text = 'SignTool.exe Path:'
$labelSignToolPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($labelSignToolPath)

$textBoxSignToolPath = New-Object System.Windows.Forms.TextBox
$textBoxSignToolPath.Location = New-Object System.Drawing.Point(20, 120)
$textBoxSignToolPath.Size = New-Object System.Drawing.Size(480, 25)
$textBoxSignToolPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($textBoxSignToolPath)

$btnBrowseSignTool = New-Object System.Windows.Forms.Button
$btnBrowseSignTool.Location = New-Object System.Drawing.Point(510, 118)
$btnBrowseSignTool.Size = New-Object System.Drawing.Size(100, 27)
$btnBrowseSignTool.Text = 'Browse...'
$btnBrowseSignTool.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($btnBrowseSignTool)

# Certificate Path
$labelCertPath = New-Object System.Windows.Forms.Label
$labelCertPath.Location = New-Object System.Drawing.Point(20, 155)
$labelCertPath.Size = New-Object System.Drawing.Size(200, 20)
$labelCertPath.Text = 'Certificate (.pfx) Path:'
$labelCertPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($labelCertPath)

$textBoxCertPath = New-Object System.Windows.Forms.TextBox
$textBoxCertPath.Location = New-Object System.Drawing.Point(20, 180)
$textBoxCertPath.Size = New-Object System.Drawing.Size(480, 25)
$textBoxCertPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($textBoxCertPath)

$btnBrowseCert = New-Object System.Windows.Forms.Button
$btnBrowseCert.Location = New-Object System.Drawing.Point(510, 178)
$btnBrowseCert.Size = New-Object System.Drawing.Size(100, 27)
$btnBrowseCert.Text = 'Browse...'
$btnBrowseCert.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($btnBrowseCert)

# Certificate Password for Signing
$labelSignPassword = New-Object System.Windows.Forms.Label
$labelSignPassword.Location = New-Object System.Drawing.Point(20, 215)
$labelSignPassword.Size = New-Object System.Drawing.Size(200, 20)
$labelSignPassword.Text = 'Certificate Password:'
$labelSignPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($labelSignPassword)

$textBoxSignPassword = New-Object System.Windows.Forms.TextBox
$textBoxSignPassword.Location = New-Object System.Drawing.Point(20, 240)
$textBoxSignPassword.Size = New-Object System.Drawing.Size(590, 25)
$textBoxSignPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxSignPassword.UseSystemPasswordChar = $true
$groupBoxSign.Controls.Add($textBoxSignPassword)

# Timestamp URL
$labelTimestamp = New-Object System.Windows.Forms.Label
$labelTimestamp.Location = New-Object System.Drawing.Point(20, 275)
$labelTimestamp.Size = New-Object System.Drawing.Size(200, 20)
$labelTimestamp.Text = 'Timestamp URL:'
$labelTimestamp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$groupBoxSign.Controls.Add($labelTimestamp)

$textBoxTimestamp = New-Object System.Windows.Forms.TextBox
$textBoxTimestamp.Location = New-Object System.Drawing.Point(20, 300)
$textBoxTimestamp.Size = New-Object System.Drawing.Size(590, 25)
$textBoxTimestamp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$textBoxTimestamp.Text = "http://timestamp.sectigo.com"
$groupBoxSign.Controls.Add($textBoxTimestamp)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 305)
$progressBar.Size = New-Object System.Drawing.Size(590, 25)
$progressBar.Style = 'Continuous'
$progressBar.Visible = $false
$groupBoxSign.Controls.Add($progressBar)

# Sign Button
$btnSign = New-Object System.Windows.Forms.Button
$btnSign.Location = New-Object System.Drawing.Point(20, 335)
$btnSign.Size = New-Object System.Drawing.Size(590, 35)
$btnSign.Text = 'Sign Binary'
$btnSign.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnSign.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$btnSign.ForeColor = [System.Drawing.Color]::White
$btnSign.FlatStyle = 'Flat'
$btnSign.Cursor = [System.Windows.Forms.Cursors]::Hand
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
    }
})

# Browse Binary File
$btnBrowseBinary.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Executable Files (*.exe;*.dll;*.msi)|*.exe;*.dll;*.msi|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Binary File to Sign"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxBinaryPath.Text = $openFileDialog.FileName
    }
})

# Browse SignTool
$btnBrowseSignTool.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "SignTool (signtool.exe)|signtool.exe|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select SignTool.exe"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxSignToolPath.Text = $openFileDialog.FileName
    }
})

# Browse Certificate
$btnBrowseCert.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Certificate Files (*.pfx;*.p12)|*.pfx;*.p12|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Certificate File"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $textBoxCertPath.Text = $openFileDialog.FileName
    }
})

# Generate Certificate Button
$btnGenerate.Add_Click({
    try {
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($textBoxDestPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a destination path.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxCNName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a certificate CN name.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxCertPassword.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a certificate password.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
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
        
        # Disable button during operation
        $btnGenerate.Enabled = $false
        $btnGenerate.Text = "Generating..."
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
        
        # Remove from store (optional - comment out if you want to keep it in the store)
        # Remove-Item -Path "Cert:\CurrentUser\My\$($cert.Thumbprint)" -Force
        
        $btnGenerate.Enabled = $true
        $btnGenerate.Text = "Generate Certificate"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Certificate generated successfully!`n`nLocation: $certPath`nThumbprint: $($cert.Thumbprint)", 
            "Success", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information)
        
                # Auto-populate the certificate path in the signing section
        
        $textBoxCertPath.Text = $certPath
        $textBoxSignPassword.Text = $password

        # Export public key
        $cerFileName = "${cleanCertName}_CodeSigningPublicKey.cer"
        $cerPath = Join-Path $textBoxDestPath.Text $cerFileName
        Export-Certificate -Cert $cert -FilePath $cerPath

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
"@
        Set-Content -Path $keyDetailsPath -Value $keyDetailsContent

    }
    catch {
        $btnGenerate.Enabled = $true
        $btnGenerate.Text = "Generate Certificate"
        
        [System.Windows.Forms.MessageBox]::Show(
            "Error generating certificate:`n`n$($_.Exception.Message)", 
            "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Sign Binary Button
$btnSign.Add_Click({
    try {
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($textBoxBinaryPath.Text) -or -not (Test-Path $textBoxBinaryPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a valid binary file path.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxSignToolPath.Text) -or -not (Test-Path $textBoxSignToolPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a valid SignTool.exe path.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxCertPath.Text) -or -not (Test-Path $textBoxCertPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a valid certificate file path.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($textBoxSignPassword.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify the certificate password.", "Validation Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Show progress bar
        $progressBar.Value = 0
        $progressBar.Visible = $true
        $btnSign.Enabled = $false
        $btnSign.Text = "Signing in progress..."
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
        
        $progressBar.Value = 100
        $form.Refresh()
        
        Start-Sleep -Milliseconds 500
        
        if ($process.ExitCode -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Binary signed successfully!`n`nFile: $binaryPath`n`nOutput:`n$output", 
                "Success", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Information)

            # Copy public key to binary path
            $pfxFileName = Split-Path -Path $certPath -Leaf
            $cleanCertName = $pfxFileName -replace '_CodeSigningPrivateKey.pfx', ''
            $cerFileName = "${cleanCertName}_CodeSigningPublicKey.cer"
            $sourceCerPath = Join-Path (Split-Path -Path $certPath -Parent) $cerFileName
            $destinationCerPath = Join-Path (Split-Path -Path $binaryPath -Parent) $cerFileName

            if (Test-Path $sourceCerPath) {
                Copy-Item -Path $sourceCerPath -Destination $destinationCerPath -Force
                [System.Windows.Forms.MessageBox]::Show(
                    "Public key copied to binary directory:`n`n$destinationCerPath", 
                    "Public Key Copied", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Public key not found at expected location: $sourceCerPath", 
                    "Warning", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
            }
        }
        else {
            throw "SignTool failed with exit code $($process.ExitCode)`n`nError: $errorOutput`n`nOutput: $output"
        }
        
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error signing binary:`n`n$($_.Exception.Message)", 
            "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $btnSign.Enabled = $true
        $btnSign.Text = "Sign Binary"
    }
})

# ====================================
# SHOW FORM
# ====================================

# Show the form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()