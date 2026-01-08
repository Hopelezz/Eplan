using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;

// Created Suplanus and modified by Mark Spratt

/* How to use this script:
 * 1. Copy this script into your EPLAN Scripts directory (e.g. C:\EPLAN\Scripts\)
 * 2. Open EPLAN and go to File -> Extras -> Interfaces -> Script: Load
 * 3. Select the script "HiddenSettings_Extension.cs" in the file dialog and open
 * 4. Click "OK" in the Register Script Dialog if it's already registered (this will update the script)
 * 5. A confirmation message will appear once the script is loaded successfully
 */

/// <summary>
/// Extension to toggle EPLAN extended context menu settings
/// </summary>
public class HiddenSettingsExtension
{
    #region Main Entry Point
    [Start]
    public void Function()
    {
        try
        {
            Settings settings = new Settings();
            bool currentValue = GetCurrentSettingValue(settings);
            
            if (ShowConfirmationDialog(currentValue))
            {
                ToggleSettingValue(settings, currentValue);
                ShowSuccessMessage(!currentValue);
            }
            else
            {
                ShowCancelledMessage();
            }
        }
        catch (System.Exception ex)
        {
            ShowErrorMessage(ex.Message);
        }
    }
    #endregion

    #region Helper Methods
    /// <summary>
    /// Gets the current value of the extended context menu setting
    /// </summary>
    /// <param name="settings">Settings object</param>
    /// <returns>Current setting value</returns>
    private bool GetCurrentSettingValue(Settings settings)
    {
        return settings.GetBoolSetting("USER.EnfMVC.ContextMenuSetting.ShowExtended", 0);
    }

    /// <summary>
    /// Shows confirmation dialog with current state and asks user to proceed
    /// </summary>
    /// <param name="currentValue">Current setting value</param>
    /// <returns>True if user wants to proceed, false otherwise</returns>
    private bool ShowConfirmationDialog(bool currentValue)
    {
        string currentState = currentValue ? "ENABLED" : "DISABLED";
        string newAction = currentValue ? "DISABLE" : "ENABLE";
        
        string message = "Extended Context Menu Setting\n\n" +
                        "Current state: " + currentState + "\n\n" +
                        "Do you want to " + newAction + " this setting?\n" +
                        "(EPLAN restart will be required)";
        
        DialogResult result = MessageBox.Show(
            message,
            "Toggle Extended Context Menu",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question
        );
        
        return result == DialogResult.Yes;
    }

    /// <summary>
    /// Toggles the setting to the opposite value
    /// </summary>
    /// <param name="settings">Settings object</param>
    /// <param name="currentValue">Current setting value</param>
    private void ToggleSettingValue(Settings settings, bool currentValue)
    {
        bool newValue = !currentValue;
        settings.SetBoolSetting("USER.EnfMVC.ContextMenuSetting.ShowExtended", newValue, 0);
    }

    /// <summary>
    /// Shows success message after setting has been changed
    /// </summary>
    /// <param name="newValue">New setting value</param>
    private void ShowSuccessMessage(bool newValue)
    {
        string status = newValue ? "enabled" : "disabled";
        string message = "Extended context menu setting has been " + status + ".\n\n" +
                        "Please restart EPLAN for changes to take effect.";
        
        MessageBox.Show(
            message,
            "Setting Changed",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    /// <summary>
    /// Shows cancellation message when user chooses not to proceed
    /// </summary>
    private void ShowCancelledMessage()
    {
        MessageBox.Show(
            "No changes were made.",
            "Cancelled",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    /// <summary>
    /// Shows error message when an exception occurs
    /// </summary>
    /// <param name="errorMessage">Error message to display</param>
    private void ShowErrorMessage(string errorMessage)
    {
        MessageBox.Show(
            "An error occurred while toggling the setting:\n" + errorMessage,
            "Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
    #endregion
}
