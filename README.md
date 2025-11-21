# Code Signing Utility

This utility provides a graphical user interface (GUI) to generate self-signed code signing certificates and sign binaries using `signtool.exe`.

## Features

*   Generate self-signed code signing certificates (`.pfx` files).
*   Export the public key (`.cer` file) for distribution.
*   Sign executable files (`.exe`, `.dll`, `.msi`) with the generated certificate.
*   User-friendly GUI for easy operation.

## How to Use

### GUI Mode

The primary way to use this utility is through its PowerShell-based GUI.

1.  **Run as Administrator**: Open a PowerShell window with administrative privileges.
2.  **Navigate to the Directory**: Change your current directory to where you've extracted the Code Signing Utility files.
3.  **Execute the Script**: Run the following command:

    ```powershell
    powershell -executionpolicy bypass .\SignToolGUI.ps1
    ```

This will launch the **Code Signing Certificate Manager** window.

#### Generate Code Signing Certificate

This section allows you to create a new code signing certificate.

1.  **Destination Path**: Choose a folder where the generated certificate files will be saved.
2.  **Certificate Subject (CN Name)**: Provide a name for your certificate (e.g., "My Company's Certificate").
3.  **Certificate Password**: Create a password for the private key (`.pfx` file).
4.  **Generate Certificate**: Click this button to generate the certificate. The following files will be created in the destination path:
    *   `_CodeSigningPrivateKey.pfx`: The private key and certificate.
    *   `_CodeSigningPublicKey.cer`: The public key.
    *   `_KeyDetails.txt`: A text file containing details about the generated key.

#### Sign Binary

This section allows you to sign a binary file using a certificate.

1.  **Binary File Path**: Select the executable file you want to sign.
2.  **SignTool.exe Path**: Locate the `signtool.exe` file. It is included in this utility's directory.
3.  **Certificate (.pfx) Path**: Select the `.pfx` file you generated.
4.  **Certificate Password**: Enter the password for the `.pfx` file.
5.  **Timestamp URL**: A default timestamp server is provided. You can change it if you have a different one.
6.  **Sign Binary**: Click this button to sign the file.

### Trusting the Certificate

To prevent security warnings when running your signed binaries, you or your users will need to install the public key (`.cer` file) to the "Trusted People" certificate store.

1.  Double-click the `.cer` file.
2.  Click "Open" -> "Install Certificate...".
3.  Select "Current User".
4.  Choose "Place all certificates in the following store".
5.  Browse to "Trusted People", click "OK", then "Next", and "Finish".

Your signed binaries should now run without warnings.

## Command-Line Tool

This utility uses `signtool.exe`, a command-line tool from the Windows SDK, for the actual code signing. While you can use `signtool.exe` directly, this GUI simplifies the process.

## License

This project is licensed under the terms of the MIT License.
