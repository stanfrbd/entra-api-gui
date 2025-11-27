<#
.SYNOPSIS
    Simple PowerShell GUI for Microsoft Entra ID User actions using Microsoft Graph API.

.DESCRIPTION
    This tool uses the Microsoft Graph API to perform bulk and specific actions on users based on a list of User Principal Names (UPNs) from a CSV file.

    The tool looks up the user's Object ID and performs administrative actions like enabling/disabling the account, managing risk state, and session revocation.

    An Azure AD App ID and Secret are required to connect to the API. The tool requires the following **Application Permissions** in Entra ID (App-only scopes):

    - User.Read.All
    - User.ReadWrite.All
    - Group.ReadWrite.All
    - IdentityRiskyUser.ReadWrite.All
    
    Administrator consent is required for all these permissions.
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


$script:selectedusers = @{} # Stores { 'User Principal Name' = 'Entra ID ObjectID' }
$credspath = 'c:\temp\entrauseruicreds.txt'
$UnclickableColour = "#8d8989"
$ClickableColour = "#0078D4" # Microsoft Blue
$TextBoxFont = 'Microsoft Sans Serif,10'
$GraphApiUrl = "https://graph.microsoft.com/v1.0"
$GraphApiBetaUrl = "https://graph.microsoft.com/beta"

# Authentication method mapping removed

#===========================================================[WinForm]===========================================================


[System.Windows.Forms.Application]::EnableVisualStyles()


$MainForm = New-Object system.Windows.Forms.Form
$MainForm.SuspendLayout()
$MainForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$MainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$MainForm.ClientSize = '950,750' # Adjusted height
$MainForm.text = "Entra ID User actions API GUI"
$MainForm.BackColor = "#ffffff"
$MainForm.TopMost = $false

# 1 - Connection Section (Same as before)
$Title = New-Object system.Windows.Forms.Label
$Title.text = "1 - Connect with Entra ID / Microsoft Graph Credentials"
$Title.AutoSize = $true
$Title.location = New-Object System.Drawing.Point(20, 20)
$Title.Font = 'Microsoft Sans Serif,12,style=Bold'

$AppIdBoxLabel = New-Object system.Windows.Forms.Label
$AppIdBoxLabel.text = "App Id:"
$AppIdBoxLabel.AutoSize = $true
$AppIdBoxLabel.location = New-Object System.Drawing.Point(20, 50)
$AppIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppIdBox = New-Object system.Windows.Forms.TextBox
$AppIdBox.multiline = $false
$AppIdBox.width = 314
$AppIdBox.height = 20
$AppIdBox.location = New-Object System.Drawing.Point(100, 50)
$AppIdBox.Font = $TextBoxFont

$AppSecretBoxLabel = New-Object system.Windows.Forms.Label
$AppSecretBoxLabel.text = "App Secret:"
$AppSecretBoxLabel.AutoSize = $true
$AppSecretBoxLabel.location = New-Object System.Drawing.Point(20, 75)
$AppSecretBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppSecretBox = New-Object system.Windows.Forms.TextBox
$AppSecretBox.multiline = $false
$AppSecretBox.width = 314
$AppSecretBox.height = 20
$AppSecretBox.location = New-Object System.Drawing.Point(100, 75)
$AppSecretBox.Font = $TextBoxFont
$AppSecretBox.PasswordChar = '*'

$TenantIdBoxLabel = New-Object system.Windows.Forms.Label
$TenantIdBoxLabel.text = "Tenant Id:"
$TenantIdBoxLabel.AutoSize = $true
$TenantIdBoxLabel.location = New-Object System.Drawing.Point(20, 100)
$TenantIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$TenantIdBox = New-Object system.Windows.Forms.TextBox
$TenantIdBox.multiline = $false
$TenantIdBox.width = 314
$TenantIdBox.height = 20
$TenantIdBox.location = New-Object System.Drawing.Point(100, 100)
$TenantIdBox.Font = $TextBoxFont

$ConnectionStatusLabel = New-Object system.Windows.Forms.Label
$ConnectionStatusLabel.text = "Status:"
$ConnectionStatusLabel.AutoSize = $true
$ConnectionStatusLabel.location = New-Object System.Drawing.Point(20, 135)
$ConnectionStatusLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$ConnectionStatus = New-Object system.Windows.Forms.Label
$ConnectionStatus.text = "Disconnected"
$ConnectionStatus.AutoSize = $true
$ConnectionStatus.location = New-Object System.Drawing.Point(100, 135)
$ConnectionStatus.Font = 'Microsoft Sans Serif,10'

$SaveCredCheckbox = new-object System.Windows.Forms.checkbox
$SaveCredCheckbox.Location = New-Object System.Drawing.Point(200, 135)
$SaveCredCheckbox.AutoSize = $true
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

# 2 - User Selection Section
$InputCsvFileBox = New-Object System.Windows.Forms.GroupBox
$InputCsvFileBox.width = 880
$InputCsvFileBox.height = 200
$InputCsvFileBox.location = New-Object System.Drawing.Point(20, 190)
$InputCsvFileBox.text = "2 - Select Users to Process (CSV)"
$InputCsvFileBox.Font = 'Microsoft Sans Serif,12,style=Bold'

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
$BrowseCsvBtn.Enabled = $false

$CsvDescLabel = New-Object system.Windows.Forms.Label
$CsvDescLabel.text = "Select a CSV file with a header named 'UserPrincipalName' (single column) containing user IDs."
$CsvDescLabel.width = 700
$CsvDescLabel.height = 40
$CsvDescLabel.location = New-Object System.Drawing.Point(20, 90)
$CsvDescLabel.Font = 'Microsoft Sans Serif,9'
$CsvDescLabel.ForeColor = "#000000"
$CsvDescLabel.Visible = $true

$GetUsersFromCsvBtn = New-Object System.Windows.Forms.Button
$GetUsersFromCsvBtn.BackColor = $UnclickableColour
$GetUsersFromCsvBtn.text = "Get Entra ID Users"
$GetUsersFromCsvBtn.width = 250
$GetUsersFromCsvBtn.height = 30
$GetUsersFromCsvBtn.location = New-Object System.Drawing.Point(610, 140)
$GetUsersFromCsvBtn.Font = 'Microsoft Sans Serif,10'
$GetUsersFromCsvBtn.ForeColor = "#ffffff"
$GetUsersFromCsvBtn.Enabled = $false

$SelectedUsersBtn = New-Object system.Windows.Forms.Button
$SelectedUsersBtn.BackColor = $UnclickableColour
$SelectedUsersBtn.text = "Selected Users (" + $script:selectedusers.Keys.count + ")"
$SelectedUsersBtn.width = 200
$SelectedUsersBtn.height = 30
$SelectedUsersBtn.location = New-Object System.Drawing.Point(400, 140)
$SelectedUsersBtn.Font = 'Microsoft Sans Serif,10'
$SelectedUsersBtn.ForeColor = "#ffffff"
$SelectedUsersBtn.Visible = $false

$ClearSelectedUsersBtn = New-Object system.Windows.Forms.Button
$ClearSelectedUsersBtn.BackColor = $UnclickableColour
$ClearSelectedUsersBtn.text = "Clear Selection"
$ClearSelectedUsersBtn.width = 180
$ClearSelectedUsersBtn.height = 30
$ClearSelectedUsersBtn.location = New-Object System.Drawing.Point(200, 140)
$ClearSelectedUsersBtn.Font = 'Microsoft Sans Serif,10'
$ClearSelectedUsersBtn.ForeColor = "#ffffff"
$ClearSelectedUsersBtn.Visible = $false

$OpenCsvDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenCsvDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenCsvDialog.Multiselect = $false


$InputCsvFileBox.Controls.AddRange(@(
        $CsvPathBox,
        $BrowseCsvBtn,
        $CsvDescLabel,
        $GetUsersFromCsvBtn,
        $SelectedUsersBtn,
        $ClearSelectedUsersBtn
    ))

# 3 - User Actions Section
$ActionGroupBox = New-Object System.Windows.Forms.GroupBox
$ActionGroupBox.Location = New-Object System.Drawing.Point(20, 410)
$ActionGroupBox.width = 910
$ActionGroupBox.height = 180 # Reduced height for cleaner layout
$ActionGroupBox.Text = "3 - Perform Action on Selected Users"
$ActionGroupBox.Font = 'Microsoft Sans Serif,12,style=Bold'

# Row 1: Account Status (Now only Enable/Disable)
$StatusLabel = New-Object system.Windows.Forms.Label
$StatusLabel.text = "Account Status:"
$StatusLabel.AutoSize = $true
$StatusLabel.location = New-Object System.Drawing.Point(20, 40)
$StatusLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$EnableUserBtn = New-Object system.Windows.Forms.Button
$EnableUserBtn.BackColor = $UnclickableColour
$EnableUserBtn.text = "Enable User"
$EnableUserBtn.width = 150
$EnableUserBtn.height = 30
$EnableUserBtn.location = New-Object System.Drawing.Point(150, 35)
$EnableUserBtn.Font = 'Microsoft Sans Serif,10'
$EnableUserBtn.ForeColor = "#ffffff"
$EnableUserBtn.Enabled = $false

$DisableUserBtn = New-Object system.Windows.Forms.Button
$DisableUserBtn.BackColor = $UnclickableColour
$DisableUserBtn.text = "Disable User"
$DisableUserBtn.width = 150
$DisableUserBtn.height = 30
$DisableUserBtn.location = New-Object System.Drawing.Point(310, 35)
$DisableUserBtn.Font = 'Microsoft Sans Serif,10'
$DisableUserBtn.ForeColor = "#ffffff"
$DisableUserBtn.Enabled = $false

# Row 2: Risk Status & Session Revocation
$RiskLabel = New-Object system.Windows.Forms.Label
$RiskLabel.text = "Risk & Session:"
$RiskLabel.AutoSize = $true
$RiskLabel.location = New-Object System.Drawing.Point(20, 80)
$RiskLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$MarkRiskyBtn = New-Object system.Windows.Forms.Button
$MarkRiskyBtn.BackColor = $UnclickableColour
$MarkRiskyBtn.text = "Mark as Risky"
$MarkRiskyBtn.width = 150
$MarkRiskyBtn.height = 30
$MarkRiskyBtn.location = New-Object System.Drawing.Point(150, 75)
$MarkRiskyBtn.Font = 'Microsoft Sans Serif,10'
$MarkRiskyBtn.ForeColor = "#ffffff"
$MarkRiskyBtn.Enabled = $false

$UnmarkRiskyBtn = New-Object system.Windows.Forms.Button
$UnmarkRiskyBtn.BackColor = $UnclickableColour
$UnmarkRiskyBtn.text = "Unmark as Risky"
$UnmarkRiskyBtn.width = 150
$UnmarkRiskyBtn.height = 30
$UnmarkRiskyBtn.location = New-Object System.Drawing.Point(310, 75)
$UnmarkRiskyBtn.Font = 'Microsoft Sans Serif,10'
$UnmarkRiskyBtn.ForeColor = "#ffffff"
$UnmarkRiskyBtn.Enabled = $false

$RevokeSignInBtn = New-Object system.Windows.Forms.Button
$RevokeSignInBtn.BackColor = $UnclickableColour
$RevokeSignInBtn.text = "Revoke Sign-in Sessions"
$RevokeSignInBtn.width = 250
$RevokeSignInBtn.height = 30
$RevokeSignInBtn.location = New-Object System.Drawing.Point(470, 75) # Moved to fill gap
$RevokeSignInBtn.Font = 'Microsoft Sans Serif,10'
$RevokeSignInBtn.ForeColor = "#ffffff"
$RevokeSignInBtn.Enabled = $false


# Row 3: Group Membership
$GroupLabel = New-Object system.Windows.Forms.Label
$GroupLabel.text = "Group Membership:"
$GroupLabel.AutoSize = $true
$GroupLabel.location = New-Object System.Drawing.Point(20, 130)
$GroupLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$GroupTextBox = New-Object system.Windows.Forms.TextBox
$GroupTextBox.multiline = $false
$GroupTextBox.width = 250
$GroupTextBox.height = 20
$GroupTextBox.location = New-Object System.Drawing.Point(150, 125)
$GroupTextBox.Font = $TextBoxFont
$GroupTextBox.PlaceholderText = "Group Display Name or ID"

$AddUserToGroupBtn = New-Object system.Windows.Forms.Button
$AddUserToGroupBtn.BackColor = $UnclickableColour
$AddUserToGroupBtn.text = "Add to Group"
$AddUserToGroupBtn.width = 150
$AddUserToGroupBtn.height = 30
$AddUserToGroupBtn.location = New-Object System.Drawing.Point(410, 125)
$AddUserToGroupBtn.Font = 'Microsoft Sans Serif,10'
$AddUserToGroupBtn.ForeColor = "#ffffff"
$AddUserToGroupBtn.Enabled = $false

$RemoveUserFromGroupBtn = New-Object system.Windows.Forms.Button
$RemoveUserFromGroupBtn.BackColor = $UnclickableColour
$RemoveUserFromGroupBtn.text = "Remove from Group"
$RemoveUserFromGroupBtn.width = 150
$RemoveUserFromGroupBtn.height = 30
$RemoveUserFromGroupBtn.location = New-Object System.Drawing.Point(570, 125)
$RemoveUserFromGroupBtn.Font = 'Microsoft Sans Serif,10'
$RemoveUserFromGroupBtn.ForeColor = "#ffffff"
$RemoveUserFromGroupBtn.Enabled = $false

# Row 4: Authentication Methods - REMOVED ALL BUTTONS/LABELS

$ActionGroupBox.Controls.AddRange(@(
        # R1
        $StatusLabel, $EnableUserBtn, $DisableUserBtn,
        # R2
        $RiskLabel, $MarkRiskyBtn, $UnmarkRiskyBtn, $RevokeSignInBtn,
        # R3
        $GroupLabel, $GroupTextBox, $AddUserToGroupBtn, $RemoveUserFromGroupBtn
    ))


# 4 - Logs and Footer
$LogBoxLabel = New-Object system.Windows.Forms.Label
$LogBoxLabel.text = "4 - Logs:"
$LogBoxLabel.width = 394
$LogBoxLabel.height = 20
$LogBoxLabel.location = New-Object System.Drawing.Point(20, 610) # New position
$LogBoxLabel.Font = 'Microsoft Sans Serif,12,style=Bold'

$LogBox = New-Object system.Windows.Forms.TextBox
$LogBox.multiline = $true
$LogBox.width = 880
$LogBox.height = 100
$LogBox.location = New-Object System.Drawing.Point(20, 640) # New position
$LogBox.ScrollBars = 'Vertical'
$LogBox.Font = $TextBoxFont

$ExportLogBtn = New-Object system.Windows.Forms.Button
$ExportLogBtn.BackColor = '#FFF0F8FF'
$ExportLogBtn.text = "Export Logs"
$ExportLogBtn.width = 120
$ExportLogBtn.height = 30
$ExportLogBtn.location = New-Object System.Drawing.Point(20, 750) # Adjusted position
$ExportLogBtn.Font = 'Microsoft Sans Serif,10'
$ExportLogBtn.ForeColor = "#ff000000"

$cancelBtn = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor = '#FFF0F8FF'
$cancelBtn.text = "Cancel"
$cancelBtn.width = 90
$cancelBtn.height = 30
$cancelBtn.location = New-Object System.Drawing.Point(810, 750) # Adjusted position
$cancelBtn.Font = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor = "#ff000000"
$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$MainForm.CancelButton = $cancelBtn

$MainForm.controls.AddRange(@(
        $Title, $AppIdBoxLabel, $AppIdBox, $AppSecretBoxLabel, $AppSecretBox, $TenantIdBoxLabel, $TenantIdBox,
        $ConnectionStatusLabel, $ConnectionStatus, $SaveCredCheckbox, $ConnectBtn, 
        $InputCsvFileBox, $ActionGroupBox, 
        $LogBoxLabel, $LogBox, $ExportLogBtn, $cancelBtn
    ))


#===========================================================[Functions]===========================================================


# --- Utility Functions ---

function GetToken {
    # Function to authenticate to Microsoft Graph API
    $ConnectionStatus.ForeColor = "#000000"
    $ConnectionStatus.Text = 'Connecting...'
    $tenantId = $TenantIdBox.Text
    $appId = $AppIdBox.Text
    $appSecret = $AppSecretBox.Text
    
    $resourceAppIdUri = 'https://graph.microsoft.com'
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
        return $null
    }

    if ($token) {
        $ConnectionStatus.text = "Connected to Microsoft Graph"
        $ConnectionStatus.ForeColor = "#7ed321"
        $LogBox.AppendText((get-date).ToString() + " Successfully connected to Entra ID (Graph API). " + [Environment]::NewLine)
        ChangeButtonColours -Buttons $GetUsersFromCsvBtn, $BrowseCsvBtn, $SelectedUsersBtn, $ClearSelectedUsersBtn
        $CsvPathBox.Enabled = $true
        $BrowseCsvBtn.Enabled = $true
        # change text from connect to Reconnect
        $ConnectBtn.text = "Reconnect"
        # hide checkbox after successful connection
        $SaveCredCheckbox.Visible = $false
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
    # Only working action buttons remain
    $actionButtons = @(
        $EnableUserBtn, $DisableUserBtn, $MarkRiskyBtn, $UnmarkRiskyBtn, 
        $RevokeSignInBtn, $AddUserToGroupBtn, $RemoveUserFromGroupBtn
    )
    ChangeButtonColours -Buttons $actionButtons
}

function HandleGraphError {
    Param(
        [Parameter(Mandatory = $true)]
        $ErrorRecord,
        [Parameter(Mandatory = $true)]
        $ActionName,
        [Parameter(Mandatory = $true)]
        $TargetName
    )
    # Default message if we can't extract details
    $ErrorMsg = "Unknown error or connection failure. Check PowerShell console for JIT details."
    $response = $ErrorRecord.Exception.Response

    # Safely check for the response object
    if ($response) {
        # Try to extract standard status code
        $ErrorMsg = $response.StatusCode.ToString()
        try {
            $ErrorDetails = $response.GetResponseStream()
            $Reader = New-Object System.IO.StreamReader($ErrorDetails)
            $Body = $Reader.ReadToEnd() | ConvertFrom-Json
            # Clean up error message for display
            $ErrorMsg = $Body.error.message -replace "`r|`n|`t", " "
        }
        catch {
            $ErrorMsg += " (Failed to parse API error body)"
        }
    }
    
    $LogBox.AppendText((get-date).ToString() + " ERROR during $ActionName for $TargetName : $ErrorMsg" + [Environment]::NewLine)
}


# --- User Selection Functions ---

function ViewSelectedUsers {
    $usersToView = @()
    $script:selectedusers.GetEnumerator() | ForEach-Object {
        $usersToView += [PSCustomObject]@{
            UserPrincipalName = $_.Key
            ObjectID          = $_.Value
        }
    }
    
    $filterUsers = $usersToView | Out-GridView -Title "Select users to process:" -PassThru 
    
    $script:selectedusers.clear()
    foreach ($user in $filterUsers) {
        $script:selectedusers.Add($user.UserPrincipalName, $user.ObjectID)
    }
    
    $SelectedUsersBtn.text = "Selected Users (" + $script:selectedusers.Keys.count + ")"
    if ($script:selectedusers.Keys.Count -gt 0) {
        EnableActionButtons
        $ClearSelectedUsersBtn.Visible = $true
    }
    else {
        ClearSelectedUsers
    }
    $LogBox.AppendText((get-date).ToString() + " Number of selected users: " + $script:selectedusers.Keys.count + [Environment]::NewLine + ($script:selectedusers.Keys -join [Environment]::NewLine) + [Environment]::NewLine)
}

function ClearSelectedUsers {
    $script:selectedusers = @{}
    $ClearSelectedUsersBtn.Visible = $false
    $SelectedUsersBtn.Visible = $false
    $SelectedUsersBtn.text = "Selected Users (0)"
    # Only working action buttons remain
    $actionButtons = @(
        $EnableUserBtn, $DisableUserBtn, $MarkRiskyBtn, $UnmarkRiskyBtn, 
        $RevokeSignInBtn, $AddUserToGroupBtn, $RemoveUserFromGroupBtn
    )
    foreach ($btn in $actionButtons) {
        $btn.Enabled = $false
        $btn.BackColor = $UnclickableColour
    }
    $LogBox.AppendText((get-date).ToString() + " User selection cleared." + [Environment]::NewLine)
}


function GetUsersFromCsv {
    if ((Test-Path $CsvPathBox.Text) -and ($CsvPathBox.Text).EndsWith(".csv")) {
        try {
            # **IMPORTANT:** CSV header MUST be 'UserPrincipalName'
            $users = Import-Csv -Path $CsvPathBox.Text
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error importing CSV. Ensure it contains a column named 'UserPrincipalName'." , "CSV Import Error")
            return
        }

        $script:selectedusers.Clear()
        $LogBox.AppendText((get-date).ToString() + " Searching for " + $users.count + " users from CSV... Please wait." + [Environment]::NewLine)
        
        foreach ($user in $users) {
            Start-Sleep -Milliseconds 200
            $UPN = $user.UserPrincipalName
            
            $url = "$GraphApiUrl/users/?`$filter=userPrincipalName eq '$UPN'&`$select=id,userPrincipalName,accountEnabled"  
            
            try { 
                $webResponse = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -ErrorAction Stop
                
                if ($webResponse.value.Count -gt 0) {
                    $UserObject = $webResponse.value | Select-Object -First 1
                    $ObjectId = $UserObject.id
                    $UPN_Found = $UserObject.userPrincipalName
                    
                    if (-not $script:selectedusers.contains($UPN_Found)) {
                        $script:selectedusers.Add($UPN_Found, $ObjectId)
                        $LogBox.AppendText((get-date).ToString() + " Found: " + $UPN_Found + " (Enabled: " + $UserObject.accountEnabled + ")" + [Environment]::NewLine)
                    }
                }
                else {
                    $LogBox.AppendText((get-date).ToString() + " Not Found: " + $UPN + " in Entra ID." + [Environment]::NewLine)
                }

            }
            Catch {
                HandleGraphError -ErrorRecord $_ -ActionName "User Lookup" -TargetName $UPN
            }
        }
        
        if ($script:selectedusers.Keys.Count -gt 0) {
            ViewSelectedUsers
            $SelectedUsersBtn.Visible = $true
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No users found matching the UPNs in the CSV." , "Information")
            ClearSelectedUsers
        }
        
        $LogBox.AppendText((get-date).ToString() + " Final number of selected users: " + $script:selectedusers.Keys.count + [Environment]::NewLine)

    } 
    else {
        [System.Windows.Forms.MessageBox]::Show($CsvPathBox.Text + " is not a valid CSV path or the file was not found." , "Error")
    }
}


# --- Action Functions (Account Status) ---

function SetUserAccountEnabled {
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )
    $ActionName = If ($Enabled) { "Enable User" } else { "Disable User" }
    $LogBox.AppendText((get-date).ToString() + " Starting $ActionName action..." + [Environment]::NewLine)
    $script:selectedusers.GetEnumerator() | foreach-object {
        Start-Sleep -Milliseconds 500
        $userId = $_.Value
        $UPN = $_.Key
        $body = @{
            "accountEnabled" = $Enabled;
        } | ConvertTo-Json

        $url = "$GraphApiUrl/users/$userId" 
        try { 
            Invoke-RestMethod -Method Patch -Uri $url -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
            $LogBox.AppendText((get-date).ToString() + " $ActionName Succeeded: " + $UPN + [Environment]::NewLine) 
        }
        Catch {
            HandleGraphError -ErrorRecord $_ -ActionName $ActionName -TargetName $UPN
        }
    }
    $LogBox.AppendText((get-date).ToString() + " $ActionName action finished." + [Environment]::NewLine)
}


# --- Action Functions (Risk Status) ---

function MarkUserRisky {
    Param(
        [Parameter(Mandatory = $true)]
        [bool]$IsRisky
    )
    
    $ActionEndpoint = If ($IsRisky) { "confirmCompromised" } else { "dismiss" }
    $ActionName = If ($IsRisky) { "Mark User as Risky" } else { "Unmark User as Risky" }

    $LogBox.AppendText((get-date).ToString() + " Starting $ActionName action..." + [Environment]::NewLine)

    $userIDsBody = @{
        "userIds" = $script:selectedusers.Values
    } | ConvertTo-Json

    # Note: identityProtection endpoints use the v1.0 URL
    $url = "$GraphApiUrl/identityProtection/riskyUsers/$ActionEndpoint"
    
    try {
        Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $userIDsBody -ContentType "application/json" -ErrorAction Stop
        $LogBox.AppendText((get-date).ToString() + " $ActionName Succeeded for all " + $script:selectedusers.Keys.Count + " users." + [Environment]::NewLine)
    }
    Catch {
        HandleGraphError -ErrorRecord $_ -ActionName $ActionName -TargetName "Batch"
        [System.Windows.Forms.MessageBox]::Show("Error during batch $ActionName operation. See logs for details." , "Error")
    }
    $LogBox.AppendText((get-date).ToString() + " $ActionName action finished." + [Environment]::NewLine)
}


# --- Action Functions (Session Revocation) ---

function RevokeSignInSessions {
    $ActionName = "Revoke Sign-in Sessions"
    $LogBox.AppendText((get-date).ToString() + " Starting $ActionName action (also invalidates all refresh tokens)..." + [Environment]::NewLine)

    $script:selectedusers.GetEnumerator() | foreach-object {
        Start-Sleep -Milliseconds 500
        $userId = $_.Value
        $UPN = $_.Key

        $url = "$GraphApiUrl/users/$userId/revokeSignInSessions" 
        try { 
            # POST request with no body
            Invoke-RestMethod -Method Post -Uri $url -Headers $headers -ContentType "application/json" -ErrorAction Stop
            $LogBox.AppendText((get-date).ToString() + " $ActionName Succeeded: " + $UPN + [Environment]::NewLine) 
        }
        Catch {
            HandleGraphError -ErrorRecord $_ -ActionName $ActionName -TargetName $UPN
        }
    }
    $LogBox.AppendText((get-date).ToString() + " $ActionName action finished." + [Environment]::NewLine)
}


# --- Authentication Method Deletion Functions (REMOVED) ---
# Removed functions: DeleteAuthenticationMethodByType, DeleteAllAuthMethods

# --- Action Functions (Group Membership) ---

function GetGroupIdByDisplayNameOrId {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$GroupIdentifier
    )
    
    # Try to treat as object ID first (GUID)
    if ($GroupIdentifier -as [System.Guid]) {
        return $GroupIdentifier
    }
    
    # Otherwise, search by display name
    $url = "$GraphApiUrl/groups/?`$filter=displayName eq '$GroupIdentifier'&`$select=id"
    try {
        $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -ErrorAction Stop
        if ($response.value.Count -gt 0) {
            return $response.value[0].id
        }
    }
    catch {
        HandleGraphError -ErrorRecord $_ -ActionName "Group Lookup" -TargetName $GroupIdentifier
    }
    return $null
}

function ManageGroupMembership {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Action # "Add" or "Remove"
    )
    
    $GroupIdentifier = $GroupTextBox.Text
    if ([string]::IsNullOrWhiteSpace($GroupIdentifier)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a Group Display Name or ID.", "Error")
        return
    }

    $GroupId = GetGroupIdByDisplayNameOrId -GroupIdentifier $GroupIdentifier
    if (!$GroupId) {
        $LogBox.AppendText((get-date).ToString() + " ERROR: Group '$GroupIdentifier' not found. Cannot proceed." + [Environment]::NewLine)
        [System.Windows.Forms.MessageBox]::Show("Group '$GroupIdentifier' not found. Check the group name/ID and logs.", "Error")
        return
    }

    $ActionName = "$Action User to Group"
    $LogBox.AppendText((get-date).ToString() + " Starting $ActionName action for Group ID $GroupId..." + [Environment]::NewLine)

    $script:selectedusers.GetEnumerator() | foreach-object {
        Start-Sleep -Milliseconds 500
        $userId = $_.Value
        $UPN = $_.Key
        
        try {
            if ($Action -eq "Add") {
                $url = "$GraphApiUrl/groups/$GroupId/members/`$ref"
                $body = @{ '@odata.id' = "$GraphApiUrl/users/$userId" } | ConvertTo-Json
                Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $body -ContentType "application/json" -ErrorAction Stop
                $LogBox.AppendText((get-date).ToString() + " $ActionName Succeeded: $UPN added to group $GroupId." + [Environment]::NewLine)
            }
            elseif ($Action -eq "Remove") {
                $url = "$GraphApiUrl/groups/$GroupId/members/$userId/`$ref"
                Invoke-RestMethod -Method Delete -Uri $url -Headers $headers -ErrorAction Stop
                $LogBox.AppendText((get-date).ToString() + " $ActionName Succeeded: $UPN removed from group $GroupId." + [Environment]::NewLine)
            }
        }
        Catch {
            HandleGraphError -ErrorRecord $_ -ActionName $ActionName -TargetName $UPN
        }
    }
    $LogBox.AppendText((get-date).ToString() + " $ActionName action finished." + [Environment]::NewLine)
}

# --- Export Log Function ---

function ExportLog {
    $LogBox.Text | Out-file .\entra_user_ui_log.txt
    $LogBox.AppendText((get-date).ToString() + " Log file created: " + (Get-Item .\entra_user_ui_log.txt).FullName + [Environment]::NewLine)
}


#===========================================================[Script]===========================================================


if (test-path $credspath) {
    $creds = Get-Content $credspath
    $pass = $creds[2] | ConvertTo-SecureString
    $unsecurePassword = [PSCredential]::new(0, $pass).GetNetworkCredential().Password
    $TenantIdBox.Text = $creds[0]
    $AppIdBox.Text = $creds[1]
    $AppSecretBox.Text = $unsecurePassword
}


$ConnectBtn.Add_Click({ GetToken })

# User Selection Handlers
$GetUsersFromCsvBtn.Add_Click({ GetUsersFromCsv })
$BrowseCsvBtn.Add_Click({
        if ($OpenCsvDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $CsvPathBox.Text = $OpenCsvDialog.FileName
            $LogBox.AppendText((get-date).ToString() + " CSV selected: " + $CsvPathBox.Text + [Environment]::NewLine)
        }
    })
$SelectedUsersBtn.Add_Click({ ViewSelectedUsers })
$ClearSelectedUsersBtn.Add_Click({ ClearSelectedUsers })


# Action Handlers (Kept button variable names as you had them, but linked to corrected functions)
# R1
$EnableUserBtn.Add_Click({ SetUserAccountEnabled -Enabled $true })
$DisableUserBtn.Add_Click({ SetUserAccountEnabled -Enabled $false })
# R2
$MarkRiskyBtn.Add_Click({ MarkUserRisky -IsRisky $true })
$UnmarkRiskyBtn.Add_Click({ MarkUserRisky -IsRisky $false })
$RevokeSignInBtn.Add_Click({ RevokeSignInSessions })
# R3 (Group Management)
$AddUserToGroupBtn.Add_Click({ ManageGroupMembership -Action "Add" })
$RemoveUserFromGroupBtn.Add_Click({ ManageGroupMembership -Action "Remove" })
# R4 (Auth Methods) - Removed all handlers for R4 buttons.

$ExportLogBtn.Add_Click({ ExportLog })

$MainForm.ResumeLayout()
[void]$MainForm.ShowDialog()