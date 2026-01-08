using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;

// Created by Mark Spratt

/* How to use this script:
 * 1. Copy this script into your EPLAN Scripts directory (e.g. C:\EPLAN\Scripts\)
 * 2. Open EPLAN and go to File -> Extras -> Interfaces -> Script: Load
 * 3. Select the script "OpenProjectFolder_Extension.cs" in the file dialog and open
 * 4. Click "OK" in the Register Script Dialog if it's already registered (this will update the script)
 * 5. A confirmation message will appear once the script is loaded successfully
 */

/// <summary>
/// Extension to add context menu options for opening project directories
/// </summary>
public class ProjectFolderExtensions
{
    #region Constants
    private const string CONTEXT_MENU_DIALOG = "PmPageObjectTreeDialog";
    private const string CONTEXT_MENU_NAME = "1007";
    #endregion

    #region Registration
    [DeclareRegister]
    public void Register()
    {
        MessageBox.Show(
            "Project Folder Extension script has been loaded successfully!\n\n" +
            "Features added:\n" +
            "• Open DOC directory via right-click menu in the Page Object Tree\n" +
            "• Open IMG directory via right-click menu in the Page Object Tree",
            "Script Loaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    [DeclareUnregister]
    public void UnRegister()
    {
        MessageBox.Show(
            "Project Folder Extension script removed",
            "Script Unloaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }
    #endregion

    #region Menu Setup
    [DeclareMenu]
    public void SetupContextMenu()
    {
        var contextMenu = new Eplan.EplApi.Gui.ContextMenu();
        var menuLocation = new ContextMenuLocation
        {
            DialogName = CONTEXT_MENU_DIALOG,
            ContextMenuName = CONTEXT_MENU_NAME
        };

        // Add menu items for both DOC and IMG directories
        contextMenu.AddMenuItem(menuLocation, "Open DOC Directory", "OpenDOCFolder", false, false);
        contextMenu.AddMenuItem(menuLocation, "Open IMG Directory", "OpenIMGFolder", false, false);
    }
    #endregion

    #region Action Methods
    [DeclareAction("OpenDOCFolder")]
    public void OpenDOCFolder()
    {
        OpenProjectDirectory("$(DOC)", "DOC", true);
    }

    [DeclareAction("OpenIMGFolder")]
    public void OpenIMGFolder()
    {
        OpenProjectDirectory("$(IMG)", "IMG", false);
    }
    #endregion

    #region Helper Methods
    /// <summary>
    /// Opens a project directory in Windows Explorer
    /// </summary>
    /// <param name="pathVariable">EPLAN path variable (e.g., "$(DOC)", "$(IMG)")</param>
    /// <param name="directoryName">Display name for the directory type</param>
    /// <param name="createIfMissing">Whether to create the directory if it doesn't exist</param>
    private void OpenProjectDirectory(string pathVariable, string directoryName, bool createIfMissing)
    {
        try
        {
            string directoryPath = PathMap.SubstitutePath(pathVariable);
            
            if (Directory.Exists(directoryPath))
            {
                Process.Start("explorer.exe", directoryPath);
                return;
            }

            // Directory doesn't exist
            string message = "There is no " + directoryName + " directory in the project.";
            
            if (createIfMissing)
            {
                var result = MessageBox.Show(
                    message + "\n\nWould you like to create it?",
                    "Directory Not Found",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.Yes)
                {
                    Directory.CreateDirectory(directoryPath);
                    Process.Start("explorer.exe", directoryPath);
                }
            }
            else
            {
                MessageBox.Show(
                    message,
                    "Directory Not Found",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Failed to open " + directoryName + " directory:\n" + ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }
    #endregion
}
