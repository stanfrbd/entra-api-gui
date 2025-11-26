<#
.SYNOPSIS
    Simple PowerShell GUI for Microsoft Entra ID API device actions.

.DESCRIPTION
    This tool uses the Microsoft Graph API to perform actions on Microsoft Entra ID (Azure AD) devices based on a list of hostnames from a CSV file.

    The tool finds the corresponding  Entra ID Object ID (Device ID) for the hostname and can then enable or disable the device.

    An Azure AD App ID and Secret are required to connect to the API. The tool requires the following **Application Permissions** in Entra ID (not delegated):

    - Directory.Read.All
    - Device.ReadWrite.All

    Ensure that you have granted administrator consent for these permissions.
#>


#===========================================================[Classes]===========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class ProcessDPI {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetProcessDPIAware();      
}
'@
$null = [ProcessDPI]::SetProcessDPIAware()


#===========================================================[Variables]===========================================================


$script:selectedmachines = @{} # Stores { 'Machine Name' = 'Entra ID ObjectID' }
$credspath = 'c:\temp\entrauicreds.txt'
$UnclickableColour = "#8d8989"
$ClickableColour = "#0078D4" # Microsoft Blue
$TextBoxFont = 'Microsoft Sans Serif,10'

#===========================================================[WinForm]===========================================================


[System.Windows.Forms.Application]::EnableVisualStyles()


$MainForm = New-Object system.Windows.Forms.Form
$MainForm.SuspendLayout()
$MainForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$MainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$MainForm.ClientSize = '950,750' 
$MainForm.text = "Entra ID Device Manager GUI"
$MainForm.BackColor = "#ffffff"
$MainForm.TopMost = $false

$Title = New-Object system.Windows.Forms.Label
$Title.text = "1 - Connect with Entra ID / Microsoft Graph Credentials"
$Title.AutoSize = $true
$Title.width = 25
$Title.height = 10
$Title.location = New-Object System.Drawing.Point(20, 20)
$Title.Font = 'Microsoft Sans Serif,12,style=Bold'

$AppIdBoxLabel = New-Object system.Windows.Forms.Label
$AppIdBoxLabel.text = "App Id:"
$AppIdBoxLabel.AutoSize = $true
$AppIdBoxLabel.width = 25
$AppIdBoxLabel.height = 10
$AppIdBoxLabel.location = New-Object System.Drawing.Point(20, 50)
$AppIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppIdBox = New-Object system.Windows.Forms.TextBox
$AppIdBox.multiline = $false
$AppIdBox.width = 314
$AppIdBox.height = 20
$AppIdBox.location = New-Object System.Drawing.Point(100, 50)
$AppIdBox.Font = $TextBoxFont
$AppIdBox.Visible = $true

$AppSecretBoxLabel = New-Object system.Windows.Forms.Label
$AppSecretBoxLabel.text = "App Secret:"
$AppSecretBoxLabel.AutoSize = $true
$AppSecretBoxLabel.width = 25
$AppSecretBoxLabel.height = 10
$AppSecretBoxLabel.location = New-Object System.Drawing.Point(20, 75)
$AppSecretBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppSecretBox = New-Object system.Windows.Forms.TextBox
$AppSecretBox.multiline = $false
$AppSecretBox.width = 314
$AppSecretBox.height = 20
$AppSecretBox.location = New-Object System.Drawing.Point(100, 75)
$AppSecretBox.Font = $TextBoxFont
$AppSecretBox.Visible = $true
$AppSecretBox.PasswordChar = '*'

$TenantIdBoxLabel = New-Object system.Windows.Forms.Label
$TenantIdBoxLabel.text = "Tenant Id:"
$TenantIdBoxLabel.AutoSize = $true
$TenantIdBoxLabel.width = 25
$TenantIdBoxLabel.height = 10
$TenantIdBoxLabel.location = New-Object System.Drawing.Point(20, 100)
$TenantIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$TenantIdBox = New-Object system.Windows.Forms.TextBox
$TenantIdBox.multiline = $false
$TenantIdBox.width = 314
$TenantIdBox.height = 20
$TenantIdBox.location = New-Object System.Drawing.Point(100, 100)
$TenantIdBox.Font = $TextBoxFont
$TenantIdBox.Visible = $true

$ConnectionStatusLabel = New-Object system.Windows.Forms.Label
$ConnectionStatusLabel.text = "Status:"
$ConnectionStatusLabel.AutoSize = $true
$ConnectionStatusLabel.width = 25
$ConnectionStatusLabel.height = 10
$ConnectionStatusLabel.location = New-Object System.Drawing.Point(20, 135)
$ConnectionStatusLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$ConnectionStatus = New-Object system.Windows.Forms.Label
$ConnectionStatus.text = "Disconnected"
$ConnectionStatus.AutoSize = $true
$ConnectionStatus.width = 25
$ConnectionStatus.height = 10
$ConnectionStatus.location = New-Object System.Drawing.Point(100, 135)
$ConnectionStatus.Font = 'Microsoft Sans Serif,10'

$SaveCredCheckbox = new-object System.Windows.Forms.checkbox
$SaveCredCheckbox.Location = New-Object System.Drawing.Point(200, 135)
$SaveCredCheckbox.AutoSize = $true
$SaveCredCheckbox.width = 60
$SaveCredCheckbox.height = 10
$SaveCredCheckbox.Text = "Save Credentials"
$SaveCredCheckbox.Font = 'Microsoft Sans Serif,10'
$SaveCredCheckbox.Checked = $false

$ConnectBtn = New-Object system.Windows.Forms.Button
$ConnectBtn.BackColor = $ClickableColour
$ConnectBtn.text = "Connect"
$ConnectBtn.width = 90
$ConnectBtn.height = 30
$ConnectBtn.location = New-Object System.Drawing.Point(325, 130)
$ConnectBtn.Font = 'Microsoft Sans Serif,10'
$ConnectBtn.ForeColor = "#ffffff"
$ConnectBtn.Visible = $True

# ACTION GROUP BOX
$EntraActionGroupBox = New-Object System.Windows.Forms.GroupBox
$EntraActionGroupBox.Location = New-Object System.Drawing.Point(500, 20)
$EntraActionGroupBox.width = 400
$EntraActionGroupBox.height = 90
$EntraActionGroupBox.Text = "3 - Manage Entra ID Device State"
$EntraActionGroupBox.Font = 'Microsoft Sans Serif,12,style=Bold'

$EnableDeviceBtn = New-Object system.Windows.Forms.Button
$EnableDeviceBtn.BackColor = $UnclickableColour
$EnableDeviceBtn.text = "Enable Device"
$EnableDeviceBtn.width = 150
$EnableDeviceBtn.height = 30
$EnableDeviceBtn.location = New-Object System.Drawing.Point(20, 40)
$EnableDeviceBtn.Font = 'Microsoft Sans Serif,10'
$EnableDeviceBtn.ForeColor = "#ffffff"
$EnableDeviceBtn.Visible = $true
$EnableDeviceBtn.Enabled = $false

$DisableDeviceBtn = New-Object system.Windows.Forms.Button
$DisableDeviceBtn.BackColor = $UnclickableColour
$DisableDeviceBtn.text = "Disable Device"
$DisableDeviceBtn.width = 150
$DisableDeviceBtn.height = 30
$DisableDeviceBtn.location = New-Object System.Drawing.Point(220, 40)
$DisableDeviceBtn.Font = 'Microsoft Sans Serif,10'
$DisableDeviceBtn.ForeColor = "#ffffff"
$DisableDeviceBtn.Visible = $true
$DisableDeviceBtn.Enabled = $false

$EntraActionGroupBox.Controls.AddRange(@($EnableDeviceBtn, $DisableDeviceBtn))


$InputCsvFileBox = New-Object System.Windows.Forms.GroupBox
$InputCsvFileBox.width = 880
$InputCsvFileBox.height = 200
$InputCsvFileBox.location = New-Object System.Drawing.Point(20, 190)
$InputCsvFileBox.text = "2 - Select Devices to Process (CSV)"
$InputCsvFileBox.Font = 'Microsoft Sans Serif,12,style=Bold'

$GetDevicesFromCsvBtn = New-Object System.Windows.Forms.Button
$GetDevicesFromCsvBtn.BackColor = $UnclickableColour
$GetDevicesFromCsvBtn.text = "Get Entra ID Devices"
$GetDevicesFromCsvBtn.width = 250
$GetDevicesFromCsvBtn.height = 30
$GetDevicesFromCsvBtn.location = New-Object System.Drawing.Point(610, 140)
$GetDevicesFromCsvBtn.Font = 'Microsoft Sans Serif,10'
$GetDevicesFromCsvBtn.ForeColor = "#ffffff"
$GetDevicesFromCsvBtn.Visible = $true

$SelectedDevicesBtn = New-Object system.Windows.Forms.Button
$SelectedDevicesBtn.BackColor = $UnclickableColour
$SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
$SelectedDevicesBtn.width = 200
$SelectedDevicesBtn.height = 30
$SelectedDevicesBtn.location = New-Object System.Drawing.Point(400, 140)
$SelectedDevicesBtn.Font = 'Microsoft Sans Serif,10'
$SelectedDevicesBtn.ForeColor = "#ffffff"
$SelectedDevicesBtn.Visible = $false

$ClearSelectedDevicesBtn = New-Object system.Windows.Forms.Button
$ClearSelectedDevicesBtn.BackColor = $UnclickableColour
$ClearSelectedDevicesBtn.text = "Clear Selection"
$ClearSelectedDevicesBtn.width = 180
$ClearSelectedDevicesBtn.height = 30
$ClearSelectedDevicesBtn.location = New-Object System.Drawing.Point(200, 140)
$ClearSelectedDevicesBtn.Font = 'Microsoft Sans Serif,10'
$ClearSelectedDevicesBtn.ForeColor = "#ffffff"
$ClearSelectedDevicesBtn.Visible = $false

# CSV file picker controls
$CsvPathBox = New-Object system.Windows.Forms.TextBox
$CsvPathBox.multiline = $false
$CsvPathBox.width = 700
$CsvPathBox.height = 25
$CsvPathBox.location = New-Object System.Drawing.Point(20, 60)
$CsvPathBox.Font = $TextBoxFont
$CsvPathBox.ReadOnly = $true
$CsvPathBox.Enabled = $false

$BrowseCsvBtn = New-Object system.Windows.Forms.Button
$BrowseCsvBtn.BackColor = $UnclickableColour
$BrowseCsvBtn.text = "Browse..."
$BrowseCsvBtn.width = 90
$BrowseCsvBtn.height = 25
$BrowseCsvBtn.location = New-Object System.Drawing.Point(730, 60)
$BrowseCsvBtn.Font = 'Microsoft Sans Serif,9'
$BrowseCsvBtn.ForeColor = "#ffffff"
$BrowseCsvBtn.Visible = $true
$BrowseCsvBtn.Enabled = $false

$CsvDescLabel = New-Object system.Windows.Forms.Label
$CsvDescLabel.text = "Select a CSV file with a header named 'Name' (single column) containing hostnames (one per line)."
$CsvDescLabel.width = 700
$CsvDescLabel.height = 40
$CsvDescLabel.location = New-Object System.Drawing.Point(20, 90)
$CsvDescLabel.Font = 'Microsoft Sans Serif,9'
$CsvDescLabel.ForeColor = "#000000"
$CsvDescLabel.Visible = $true

# OpenFileDialog for CSV selection
$OpenCsvDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenCsvDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenCsvDialog.Multiselect = $false


$InputCsvFileBox.Controls.AddRange(@(
        $CsvPathBox,
        $BrowseCsvBtn,
        $CsvDescLabel,
        $GetDevicesFromCsvBtn,
        $SelectedDevicesBtn,
        $ClearSelectedDevicesBtn
    ))

$LogBoxLabel = New-Object system.Windows.Forms.Label
$LogBoxLabel.text = "4 - Logs:"
$LogBoxLabel.width = 394
$LogBoxLabel.height = 20
$LogBoxLabel.location = New-Object System.Drawing.Point(20, 400)
$LogBoxLabel.Font = 'Microsoft Sans Serif,12,style=Bold'
$LogBoxLabel.Visible = $true

$LogBox = New-Object system.Windows.Forms.TextBox
$LogBox.multiline = $true
$LogBox.width = 880
$LogBox.height = 200
$LogBox.location = New-Object System.Drawing.Point(20, 430)
$LogBox.ScrollBars = 'Vertical'
$LogBox.Font = $TextBoxFont
$LogBox.Visible = $true

$ExportLogBtn = New-Object system.Windows.Forms.Button
$ExportLogBtn.BackColor = '#FFF0F8FF'
$ExportLogBtn.text = "Export Logs"
$ExportLogBtn.width = 120
$ExportLogBtn.height = 30
$ExportLogBtn.location = New-Object System.Drawing.Point(20, 650)
$ExportLogBtn.Font = 'Microsoft Sans Serif,10'
$ExportLogBtn.ForeColor = "#ff000000"
$ExportLogBtn.Visible = $true

$cancelBtn = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor = '#FFF0F8FF'
$cancelBtn.text = "Cancel"
$cancelBtn.width = 90
$cancelBtn.height = 30
$cancelBtn.location = New-Object System.Drawing.Point(810, 650)
$cancelBtn.Font = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor = "#ff000000"
$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$MainForm.CancelButton = $cancelBtn
$MainForm.Controls.Add($cancelBtn)


$MainForm.controls.AddRange(@($Title,
        $ConnectionStatusLabel, 
        $ConnectionStatus,
        $cancelBtn, 
        $AppIdBox, 
        $AppSecretBox,
        $TenantIdBox, 
        $AppIdBoxLabel, 
        $AppSecretBoxLabel, 
        $TenantIdBoxLabel, 
        $ConnectBtn, 
        $EntraActionGroupBox, 
        $LogBoxLabel, 
        $LogBox, 
        $SaveCredCheckbox,
        $InputCsvFileBox,
        $ExportLogBtn))


#===========================================================[Functions]===========================================================


function GetToken {
    # Function to authenticate to Microsoft Graph API
    $ConnectionStatus.ForeColor = "#000000"
    $ConnectionStatus.Text = 'Connecting...'
    $tenantId = $TenantIdBox.Text
    $appId = $AppIdBox.Text
    $appSecret = $AppSecretBox.Text
    
    # Microsoft Graph resource URI
    $resourceAppIdUri = 'https://graph.microsoft.com'
    # Azure AD v1.0 endpoint (compatible with client_credentials flow)
    $oAuthUri = "https://login.windows.net/$TenantId/oauth2/token"
    
    $authBody = [Ordered] @{
        resource      = "$resourceAppIdUri"
        client_id     = "$appId"
        client_secret = "$appSecret"
        grant_type    = 'client_credentials'
    }
    
    try {
        $authResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ErrorAction Stop
        $token = $authResponse.access_token
        $script:headers = @{
            'Content-Type' = 'application/json'
            Accept         = 'application/json'
            Authorization  = "Bearer $token"
        }
    }
    Catch {
        $ConnectionStatus.text = "Connection Failed"
        $LogBox.AppendText((get-date).ToString() + " Error: Failed to get Graph API access token. " + [Environment]::NewLine)
        [System.Windows.Forms.MessageBox]::Show("Connection Error: " + $_.Exception.Message , "Error")
        $ConnectionStatus.ForeColor = "#D0021B"
        $cancelBtn.text = "Close"
        return $null
    }

    if ($token) {
        $ConnectionStatus.text = "Connected to Microsoft Graph"
        $ConnectionStatus.ForeColor = "#7ed321"
        $LogBox.AppendText((get-date).ToString() + " Successfully connected to Entra ID (Graph API). " + [Environment]::NewLine)
        # hide "save creds" checkbox
        $SaveCredCheckbox.visible = $false
        # change text of connect to "Reconnect"
        $ConnectBtn.text = "Reconnect"
        
        # Enable device search/selection buttons
        ChangeButtonColours -Buttons $GetDevicesFromCsvBtn, $BrowseCsvBtn
        $CsvPathBox.Enabled = $true
        $BrowseCsvBtn.Enabled = $true
        SaveCreds
        return $headers
    }
}

function SaveCreds {
    if ($SaveCredCheckbox.Checked) {
        $securespassword = $AppSecretBox.Text | ConvertTo-SecureString -AsPlainText -Force
        $securestring = $securespassword | ConvertFrom-SecureString
        $creds = @($TenantIdBox.Text, $AppIdBox.Text, $securestring)
        $creds | Out-File $credspath
    }
}

function ChangeButtonColours {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        $Buttons
    )
    $ButtonsToChangeColour = $Buttons

    foreach ( $Button in $ButtonsToChangeColour) {
        $Button.BackColor = $ClickableColour
        $Button.Enabled = $true
    }
}

function EnableActionButtons {
    ChangeButtonColours -Buttons $EnableDeviceBtn, $DisableDeviceBtn, $SelectedDevicesBtn, $ClearSelectedDevicesBtn
    $EnableDeviceBtn.Enabled = $true
    $DisableDeviceBtn.Enabled = $true
}

function EnableDevice {
    # Enables the Entra ID device (sets accountEnabled: true)
    $LogBox.AppendText((get-date).ToString() + " Starting device Enable action..." + [Environment]::NewLine)
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Milliseconds 500
        $deviceId = $_.Value
        $MachineName = $_.Key
        $body = @{
            "accountEnabled" = $true;
        } | ConvertTo-Json

        $url = "https://graph.microsoft.com/v1.0/devices/$deviceId" 
        try { 
            # Use Invoke-RestMethod for PATCH
            $webResponse = Invoke-RestMethod -Method Patch -Uri $url -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
            $LogBox.AppendText((get-date).ToString() + " Device Enabled: " + $MachineName + " (ID: $deviceId)" + [Environment]::NewLine)
        }
        Catch {
            $ErrorMsg = $_.Exception.Response.StatusCode
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $ErrorMsg = $_.ErrorDetails.Message
            }
            $LogBox.AppendText((get-date).ToString() + " ERROR enabling " + $MachineName + ": " + $ErrorMsg + [Environment]::NewLine) 
            [System.Windows.Forms.MessageBox]::Show("Error for " + $MachineName + ": " + $ErrorMsg , "Error")
        }
    }
    $LogBox.AppendText((get-date).ToString() + " Device Enable action finished." + [Environment]::NewLine)
}

function DisableDevice {
    # Disables the Entra ID device (sets accountEnabled: false)
    $LogBox.AppendText((get-date).ToString() + " Starting device Disable action..." + [Environment]::NewLine)
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Milliseconds 500
        $deviceId = $_.Value
        $MachineName = $_.Key
        $body = @{
            "accountEnabled" = $false;
        } | ConvertTo-Json

        $url = "https://graph.microsoft.com/v1.0/devices/$deviceId" 
        try { 
            # Use Invoke-RestMethod for PATCH
            $webResponse = Invoke-RestMethod -Method Patch -Uri $url -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
            $LogBox.AppendText((get-date).ToString() + " Device Disabled: " + $MachineName + " (ID: $deviceId)" + [Environment]::NewLine) 
        }
        Catch {
            $ErrorMsg = $_.Exception.Response.StatusCode
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $ErrorMsg = $_.ErrorDetails.Message
            }
            $LogBox.AppendText((get-date).ToString() + " ERROR disabling " + $MachineName + ": " + $ErrorMsg + [Environment]::NewLine) 
            [System.Windows.Forms.MessageBox]::Show("Error for " + $MachineName + ": " + $ErrorMsg , "Error")
        }
    }
    $LogBox.AppendText((get-date).ToString() + " Device Disable action finished." + [Environment]::NewLine)
}


function ViewSelectedDevices {
    $devicesToView = @()
    $script:selectedmachines.GetEnumerator() | ForEach-Object {
        $devicesToView += [PSCustomObject]@{
            Name     = $_.Key
            ObjectID = $_.Value
        }
    }
    
    $filtermachines = $devicesToView | Out-GridView -Title "Select devices to process:" -PassThru 
    
    $script:selectedmachines.clear()
    foreach ($machine in $filtermachines) {
        $script:selectedmachines.Add($machine.Name, $machine.ObjectID)
    }
    
    $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
    if ($null -eq $script:selectedmachines.Keys.Count -or $script:selectedmachines.Keys.Count -eq 0) {
        $SelectedDevicesBtn.Visible = $false
        $ClearSelectedDevicesBtn.Visible = $false
        $EnableDeviceBtn.Enabled = $false
        $DisableDeviceBtn.Enabled = $false
        $EnableDeviceBtn.BackColor = $UnclickableColour
        $DisableDeviceBtn.BackColor = $UnclickableColour
    }
    else {
        EnableActionButtons
    }
    $LogBox.AppendText((get-date).ToString() + " Number of selected devices: " + $script:selectedmachines.Keys.count + [Environment]::NewLine + ($script:selectedmachines.Keys -join [Environment]::NewLine) + [Environment]::NewLine)
}

function ClearSelectedDevices {
    $script:selectedmachines = @{}
    $ClearSelectedDevicesBtn.Visible = $false
    $SelectedDevicesBtn.Visible = $false
    $SelectedDevicesBtn.text = "Selected Devices (0)"
    $EnableDeviceBtn.Enabled = $false
    $DisableDeviceBtn.Enabled = $false
    $EnableDeviceBtn.BackColor = $UnclickableColour
    $DisableDeviceBtn.BackColor = $UnclickableColour
    $LogBox.AppendText((get-date).ToString() + " Device selection cleared." + [Environment]::NewLine)
}


function GetDevicesFromCsv {
    # Searches for Entra ID devices using hostname from CSV
    if ((Test-Path $CsvPathBox.Text) -and ($CsvPathBox.Text).EndsWith(".csv")) {
        try {
            $machines = Import-Csv -Path $CsvPathBox.Text
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error importing CSV. Ensure it contains a column named 'Name'." , "CSV Import Error")
            return
        }

        $script:selectedmachines.Clear()
        $LogBox.AppendText((get-date).ToString() + " Searching for " + $machines.count + " Entra ID devices from CSV... Please wait." + [Environment]::NewLine)
        
        foreach ($machine in $machines) {
            # Add a small delay to avoid hitting throttling limits
            Start-Sleep -Milliseconds 500
            $MachineName = $machine.Name
            
            # Use Graph API filter for displayName (hostname)
            # Use ` to escape the $ in $filter and $select
            $url = "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$MachineName'&`$select=id,displayName"  
            
            try { 
                $webResponse = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -ErrorAction Stop
                
                if ($webResponse.value.Count -gt 0) {
                    $DeviceObject = $webResponse.value | Select-Object -First 1
                    $ObjectId = $DeviceObject.id
                    $DisplayName = $DeviceObject.displayName
                    
                    if (-not $script:selectedmachines.contains($DisplayName)) {
                        $script:selectedmachines.Add($DisplayName, $ObjectId)
                        $LogBox.AppendText((get-date).ToString() + " Found: " + $DisplayName + " (ID: $ObjectId)" + [Environment]::NewLine)
                    }
                }
                else {
                    $LogBox.AppendText((get-date).ToString() + " Not Found: " + $MachineName + " in Entra ID." + [Environment]::NewLine)
                }

            }
            Catch {
                $ErrorMsg = $_.Exception.Response.StatusCode
                if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                    $ErrorMsg = $_.ErrorDetails.Message
                }
                $LogBox.AppendText((get-date).ToString() + " Search ERROR for " + $MachineName + ": " + $ErrorMsg + [Environment]::NewLine)
            }
        }
        
        if ($script:selectedmachines.Keys.Count -gt 0) {
            # Open Out-GridView to filter the selection
            $devicesToView = @()
            $script:selectedmachines.GetEnumerator() | ForEach-Object {
                $devicesToView += [PSCustomObject]@{
                    Name     = $_.Key
                    ObjectID = $_.Value
                }
            }
            $filtermachines = $devicesToView | Out-GridView -Title "Select devices to process:" -PassThru 
            
            $script:selectedmachines.Clear()
            foreach ($machine in $filtermachines) {
                $script:selectedmachines.Add($machine.Name, $machine.ObjectID)
            }

            if ($script:selectedmachines.Keys.Count -gt 0) {
                EnableActionButtons
                $SelectedDevicesBtn.Visible = $true
                $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
                $ClearSelectedDevicesBtn.Visible = $true
            }
            else {
                $SelectedDevicesBtn.Visible = $false
                $ClearSelectedDevicesBtn.Visible = $false
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No devices found matching the hostnames in the CSV." , "Information")
            $SelectedDevicesBtn.Visible = $false
            $ClearSelectedDevicesBtn.Visible = $false
        }
        
        $LogBox.AppendText((get-date).ToString() + " Final number of selected devices: " + $script:selectedmachines.Keys.count + [Environment]::NewLine)

    } 
    else {
        [System.Windows.Forms.MessageBox]::Show($CsvPathBox.Text + " is not a valid CSV path or the file was not found." , "Error")
    }
}


function ExportLog {
    $LogBox.Text | Out-file .\entra_ui_log.txt
    $LogBox.AppendText((get-date).ToString() + " Log file created: " + (Get-Item .\entra_ui_log.txt).FullName + [Environment]::NewLine)
}


#===========================================================[Script]===========================================================


if (test-path $credspath) {
    # Load saved credentials
    $creds = Get-Content $credspath
    $pass = $creds[2] | ConvertTo-SecureString
    $unsecurePassword = [PSCredential]::new(0, $pass).GetNetworkCredential().Password
    $TenantIdBox.Text = $creds[0]
    $AppIdBox.Text = $creds[1]
    $AppSecretBox.Text = $unsecurePassword
}


$ConnectBtn.Add_Click({ GetToken })

# Event Handlers for Entra ID Actions
$EnableDeviceBtn.Add_Click({ EnableDevice })
$DisableDeviceBtn.Add_Click({ DisableDevice })

$GetDevicesFromCsvBtn.Add_Click({ GetDevicesFromCsv })

$BrowseCsvBtn.Add_Click({
        if ($OpenCsvDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $CsvPathBox.Text = $OpenCsvDialog.FileName
            $LogBox.AppendText((get-date).ToString() + " CSV selected: " + $CsvPathBox.Text + [Environment]::NewLine)
        }
    })

$SelectedDevicesBtn.Add_Click({ ViewSelectedDevices })

$ClearSelectedDevicesBtn.Add_Click({ ClearSelectedDevices })

$ExportLogBtn.Add_Click({ ExportLog })

$MainForm.ResumeLayout()
[void]$MainForm.ShowDialog()