# GUI for Entra API sample app
Simple PowerShell GUI for Microsoft Entra ID actions on devices.

Inspired from this repo from Microsoft: https://github.com/microsoft/mde-api-gui/

> [!IMPORTANT]
> This project has nothing to do with Microsoft.

> [!NOTE]
> If you intend to use this with many machines (100+), consider adding throttling handling to avoid API rate limiting. There is already one with 500 milliseconds delay between each request, but it may not be enough. And note that it takes a lot of time to process many machines in sequence=> 100 machines ~ 1min30s

## Pros

- No installation of SDK needed
- Quick to execute and simple GUI
- Very useful in case of critical incident
- Has a file picker for CSV

## Cons

- Will be more difficult to keep up to date

<img width="942" height="772" alt="image" src="https://github.com/user-attachments/assets/fecb94c0-e80a-44ae-86c8-9ea8c35d31ed" />

## Why this?

This tool is particularly useful when you need to prevent users from accessing company resources with **non-compliant machines** or during a **security incident**, especially for unmanaged devices (BYOD, no EDR) that **cannot be isolated** through other means. It implies that you have a way to identify the affected devices (e.g. compromised user accounts, IP addresses, etc.) and map them to device names in Entra ID. 

> [!NOTE]
> You could probaly already do that in the GUI but sometimes it's just faster to do it this way (e.g. CSIRT doesn't have admin rights in Entra ID portal, but has API app credentials).

## Get started
1. Create Azure AD application as described here: https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app
2. Grant the following API permissions to the application:

| Permission | Description |
|-------------------------|----------------------|
| `Device.ReadWrite.All` | Allows the app to read and write properties of devices without a signed-in user. |
| `Directory.Read.All` | Allows the app to read data in your organization's directory, such as users and groups, without a signed-in user. |

3. Create application secret.
## Usage
1. **Connect** with AAD Tenant ID, Application Id and Application Secret of the application created earlier.
2. **Get Devices** that you want to perform actions on, using one of the following methods:
    * CSV file: single Name column with machine hostnames ("Device Name" in Entra ID)
3. Confirm selection in PowerShell forms pop-up.
4. Choose action that you want to perform on **Selected Devices**, the following actions are currently available:
    * Disable Device in Entra ID
    * Enable Device in Entra ID
5. Verify actions result with **Logs** text box.
