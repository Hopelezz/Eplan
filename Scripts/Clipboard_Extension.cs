using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;

// Created by Mark Spratt

/* How to use this script:
 * 1. Copy this script into your EPLAN Scripts directory (e.g., C:\EPLAN\Scripts\)
 * 2. Open EPLAN and go to File -> Extras -> Interfaces -> Script: Load
 * 3. Select the script "Clipboard_Extension.cs" in the file dialog and open
 * 4. Click "OK" in the Register Script Dialog if it's already registered (this will update the script)
 * 5. A confirmation message will appear once the script is loaded successfully
 */

/// <summary>
/// Extension to add clipboard image functionality to the EPLAN editor
/// </summary>
public class ClipboardExtension
{
    #region Constants
    private const string CONTEXT_MENU_DIALOG = "Editor";
    private const string CONTEXT_MENU_NAME = "Ged";
    private const string COMPANY_PREFIX = "JEUS"; // Change this to your company name or initials
    private const string IMAGE_FORMAT = ".png";
    #endregion

    #region Registration
    [DeclareRegister]
    public void Register()
    {
        MessageBox.Show(
            "Clipboard Extension script has been loaded successfully!\n\n" +
            "Features added:\n" +
            "• Insert images from clipboard via right-click menu in the editor",
            "Script Loaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    [DeclareUnregister]
    public void UnRegister()
    {
        MessageBox.Show(
            "Clipboard Extension script removed",
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

        contextMenu.AddMenuItem(menuLocation, "Insert Image from Clipboard", "ClipboardImage", false, false);
    }
    #endregion

    #region Action Methods
    [DeclareAction("ClipboardImage")]
    public void InsertClipboardImage()
    {
        try
        {
            if (!Clipboard.ContainsImage())
            {
                MessageBox.Show(
                    "No image found in your clipboard.\n\nPlease copy an image first and try again.",
                    "No Clipboard Image",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return;
            }

            string imagePath = GenerateImageFileName();
            SaveClipboardImage(imagePath);
            InsertImageIntoEditor(imagePath);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Failed to insert image from clipboard:\n" + ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }
    #endregion

    #region Helper Methods
    /// <summary>
    /// Generates a unique filename for the clipboard image
    /// </summary>
    /// <returns>Full path to the image file</returns>
    private string GenerateImageFileName()
    {
        string imagesDirectory = PathMap.SubstitutePath("$(IMG)");
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        string fileName = COMPANY_PREFIX + "_" + timestamp + IMAGE_FORMAT;
        
        return Path.Combine(imagesDirectory, fileName);
    }

    /// <summary>
    /// Saves the clipboard image to the specified path
    /// </summary>
    /// <param name="imagePath">Path where to save the image</param>
    private void SaveClipboardImage(string imagePath)
    {
        // Ensure the directory exists
        string directory = Path.GetDirectoryName(imagePath);
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using (Image clipboardImage = Clipboard.GetImage())
        {
            clipboardImage.Save(imagePath, ImageFormat.Png);
        }
    }

    /// <summary>
    /// Inserts the saved image into the EPLAN editor
    /// </summary>
    /// <param name="imagePath">Path to the image file</param>
    private void InsertImageIntoEditor(string imagePath)
    {
        var actionContext = new ActionCallingContext();
        actionContext.AddParameter("Name", "XGedIaInsertImage");
        actionContext.AddParameter("Filename", imagePath);

        var commandInterpreter = new CommandLineInterpreter();
        commandInterpreter.Execute("XGedStartInteractionAction2D", actionContext);
    }
    #endregion
}

