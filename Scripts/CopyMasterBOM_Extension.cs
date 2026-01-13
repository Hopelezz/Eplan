using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;

// Created by Mark Spratt

/* How to use this script:
 * 1. Copy this script into your EPLAN Scripts directory (e.g. C:\EPLAN\Scripts\)
 * 2. Open EPLAN and go to File -> Extras -> Interfaces -> Script: Load
 * 3. Select the script "MasterBOM_Extension.cs" in the file dialog and open
 * 4. Click "OK" in the Register Script Dialog if it's already registered (this will update the script)
 * 5. A confirmation message will appear once the script is loaded successfully
 * 6. Use the "Copy Master BOM" button in Tools > Scripts ribbon
    * Features:
    * • Copy Master BOM to project DOC folder with automatic renaming
    * • Extract project number from F#### format
    * • Auto-populate BOM with project name, number, and editor name
    * • Built-in progress indicator and file overwrite protection
 */

/// <summary>
/// Extension to automate Master BOM copying and project information population
/// </summary>
public class MasterBOM_Extension
{
    #region Constants
    private const string TAB_NAME_EN = "Tools";
    private const string TAB_NAME_DE = "Extras";
    private const string GROUP_NAME = "Scripts";

    // Master BOM file locations - add more paths as needed
    // Multiple common locations are used in case the Hard Drive letter or network path differs due to structural changes.
    private static readonly string[] MASTER_BOM_PATHS = {
        @"E:\Electrical Design Team\Form\Master BoM.xlsm",
        @"\\name.intra\dfs01\Engineering\Electrical Design Team\Form\Master BoM.xlsm",
        // Add more possible locations here
    };
    
    /*
        Excel worksheet and cell coordinates for project information
        For our Excel BOM template, we have a Master that displays all worksheets:
        - Project Name goes to cell B5
        - Project Number goes to cell D3
        - Editor Name goes to cell D6
     */
    private const string WORKSHEET_NAME = "005";
    private const string PROJECT_NAME_CELL = "B5";
    private const string PROJECT_NUMBER_CELL = "D3";
    private const string EDITOR_NAME_CELL = "D6";
    #endregion

    #region Registration
    [DeclareRegister]
    public void Register()
    {
        SetupRibbonInterface();
        
        MessageBox.Show(
            "Master BOM Extension script has been loaded successfully!\n\n" +
            "Features added:\n" +
            "• Copy Master BOM to project DOC folder with automatic renaming\n" +
            "• Extract project number from F#### format\n" +
            "• Auto-populate BOM with project name, number, and editor name\n" +
            "• Built-in progress indicator and file overwrite protection\n\n" +
            "Find the 'Copy Master BOM' button in Tools > Scripts",
            "Script Loaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    [DeclareUnregister]
    public void UnRegister()
    {
        RemoveRibbonInterface();
        
        MessageBox.Show(
            "Master BOM Extension script removed successfully!\n\n" +
            "The 'Copy Master BOM' button has been removed from the Tools tab.",
            "Script Unloaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }
    #endregion

    #region Ribbon Interface Setup
    /// <summary>
    /// Gets the localized tab name for the ribbon
    /// </summary>
    private MultiLangString GetTabName()
    {
        MultiLangString tabName = new MultiLangString();
        tabName.AddString(ISOCode.Language.L_de_DE, TAB_NAME_DE);
        tabName.AddString(ISOCode.Language.L_en_US, TAB_NAME_EN);
        return tabName;
    }

    /// <summary>
    /// Sets up the ribbon interface with the Copy Master BOM button
    /// </summary>
    private void SetupRibbonInterface()
    {
        RibbonBar ribbonBar = new RibbonBar();

        RibbonTab ribbonTab = ribbonBar.GetTab(GetTabName(), true);
        if (ribbonTab == null)
        {
            ribbonTab = ribbonBar.AddTab(GetTabName());
        }

        RibbonCommandGroup ribbonCommandGroup = ribbonTab.AddCommandGroup(GROUP_NAME);
        ribbonCommandGroup.AddCommand("Copy Master BOM", "CopyMasterBOMToProject", new RibbonIcon(CommandIcon.Application));
    }

    /// <summary>
    /// Removes the ribbon interface elements
    /// </summary>
    private void RemoveRibbonInterface()
    {
        RibbonBar ribbonBar = new RibbonBar();
        RibbonTab ribbonTab = ribbonBar.GetTab(GetTabName(), true);
        if (ribbonTab != null)
        {
            RibbonCommandGroup ribbonCommandGroup = ribbonTab.GetCommandGroup(GROUP_NAME);
            if (ribbonCommandGroup != null)
            {
                ribbonCommandGroup.Remove();
            }
        }
    }
    #endregion

    #region Main Action Methods
    [DeclareAction("CopyMasterBOMToProject")]
    public void CopyMasterBOMToProject()
    {
        Progress progress = new Progress("CopyMasterBOM");
        progress.SetAllowCancel(false);
        progress.ShowImmediately();
        
        try
        {
            // Step 1: Get project information
            progress.BeginPart(15, "Getting project information...");
            string projectPath = PathMap.SubstitutePath("$(PROJECTPATH)");
            string projectName = PathMap.SubstitutePath("$(PROJECTNAME)");
            
            if (string.IsNullOrEmpty(projectPath) || string.IsNullOrEmpty(projectName))
            {
                progress.EndPart(true);
                MessageBox.Show("No project is currently open.", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            progress.EndPart();

            // Step 2: Create DOC folder
            progress.BeginPart(10, "Creating DOC folder...");
            string docFolderPath = Path.Combine(projectPath, "DOC");

            if (!Directory.Exists(docFolderPath))
            {
                Directory.CreateDirectory(docFolderPath);
            }
            progress.EndPart();

            // Step 3: Find Master BOM
            progress.BeginPart(15, "Locating Master BOM file...");
            string masterBOMPath = FindMasterBOMFile();
            if (string.IsNullOrEmpty(masterBOMPath) || !File.Exists(masterBOMPath))
            {
                progress.EndPart(true);
                MessageBox.Show("Master BOM file not found. Please check the Master BOM location.", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            progress.EndPart();

            // Step 4: Check if destination file exists and handle overwrite
            progress.BeginPart(20, "Checking destination file...");
            string destinationPath = Path.Combine(docFolderPath, projectName + " BOM.xlsm");
            
            if (File.Exists(destinationPath))
            {
                progress.EndPart(true); // Close progress for user interaction
                
                DialogResult overwriteResult = MessageBox.Show(
                    "The BOM file already exists:\n" + destinationPath + "\n\nDo you want to overwrite it?",
                    "File Already Exists",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (overwriteResult == DialogResult.No)
                {
                    // Ask if they want to open the existing file instead
                    DialogResult openResult = MessageBox.Show(
                        "Would you like to open the existing BOM file?",
                        "Open Existing File",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    
                    if (openResult == DialogResult.Yes)
                    {
                        try
                        {
                            Process.Start(destinationPath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not open BOM file: " + ex.Message, "Error", 
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    return; // Exit the method
                }
                
                // User chose to overwrite, restart progress
                progress = new Progress("CopyMasterBOM");
                progress.SetAllowCancel(false);
                progress.ShowImmediately();
                progress.BeginPart(30, "Overwriting existing BOM file...");
            }
            else
            {
                progress.EndPart();
                progress.BeginPart(30, "Copying Master BOM file...");
            }
            
            // Copy the file (overwrite if user confirmed)
            File.Copy(masterBOMPath, destinationPath, true);
            progress.EndPart();

            // Step 5: Extract information
            progress.BeginPart(10, "Extracting project information...");
            string projectNumber = ExtractProjectNumber(projectName);
            string editorName = GetEditorName();
            progress.EndPart();

            // Step 6: Update BOM
            progress.BeginPart(30, "Updating BOM with project information...");
            bool bomUpdated = UpdateBOMWithProjectInfo(destinationPath, projectName, projectNumber, editorName);
            progress.EndPart();

            // Show success message and ask what the user wants to open
            string message = bomUpdated ? 
                "Master BOM copied and updated successfully to:\n" + destinationPath + "\n\nProject name has been applied to cell " + PROJECT_NAME_CELL + " in worksheet '" + WORKSHEET_NAME + "'." :
                "Master BOM copied successfully to:\n" + destinationPath + "\n\nNote: Could not automatically update project name in BOM.";
            
            // Use custom dialog with properly named buttons
            CustomActionDialog actionDialog = new CustomActionDialog(message);
            DialogResult result = actionDialog.ShowDialog();

            if (result == DialogResult.Yes) // BOM button
            {
                // Open the BOM file
                try
                {
                    Process.Start(destinationPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open BOM file: " + ex.Message, "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (result == DialogResult.No) // DOC button
            {
                // Open the DOC folder
                Process.Start("explorer", docFolderPath);
            }
            // If Cancel, do nothing
        }
        catch (Exception ex)
        {
            progress.EndPart(true);
            MessageBox.Show("Error copying Master BOM: " + ex.Message, 
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            progress.EndPart(true);
        }
    }

    [Start]
    public void Function()
    {
        CopyMasterBOMToProject();
    }
    #endregion

    #region Helper Methods
    /// <summary>
    /// Locates the Master BOM file from predefined locations or prompts user to select
    /// </summary>
    /// <returns>Path to the Master BOM file</returns>
    private string FindMasterBOMFile()
    {
        // Check predefined locations first
        foreach (string path in MASTER_BOM_PATHS)
        {
            if (File.Exists(path))
            {
                return path;
            }
        }

        // If not found in common locations, prompt user to select the file
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Title = "Select Master BOM File";
            openFileDialog.Filter = "Excel files (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = @"C:\EPLAN\Data";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Extracts project number from project name using F#### pattern
    /// </summary>
    /// <param name="projectName">The project name</param>
    /// <returns>Extracted project number or empty string</returns>
    private string ExtractProjectNumber(string projectName)
    {
        // Use regex to find F followed by digits (F####)
        Match match = Regex.Match(projectName, @"F(\d+)");
        if (match.Success)
        {
            return match.Groups[1].Value; // Return just the number part
        }
        return ""; // Return empty if no F#### pattern found
    }

    /// <summary>
    /// Gets the current editor's name from environment
    /// </summary>
    /// <returns>Formatted editor name</returns>
    private string GetEditorName()
    {
        try
        {
            string userName = Environment.UserName;
            return FormatUserName(userName);
        }
        catch
        {
            return "";
        }
    }

    /// <summary>
    /// Formats username for display (e.g., "first.last" -> "F. Last")
    /// </summary>
    /// <param name="userName">Raw username</param>
    /// <returns>Formatted username</returns>
    private string FormatUserName(string userName)
    {
        if (string.IsNullOrEmpty(userName))
            return "";

        try
        {
            // Handle format like "first.last" -> "F. Last"
            if (userName.Contains("."))
            {
                string[] parts = userName.Split('.');
                if (parts.Length >= 2)
                {
                    string firstName = parts[0];
                    string lastName = parts[1];
                    
                    // Capitalize first letter of first name, first letter of last name
                    string formattedFirstName = firstName.Length > 0 ? firstName.Substring(0, 1).ToUpper() : "";
                    string formattedLastName = lastName.Length > 0 ? 
                        lastName.Substring(0, 1).ToUpper() + lastName.Substring(1).ToLower() : "";
                    
                    return formattedFirstName + ". " + formattedLastName;
                }
            }
            
            // Handle other formats or fallback - just capitalize first letter
            if (userName.Length > 0)
            {
                return userName.Substring(0, 1).ToUpper() + userName.Substring(1).ToLower();
            }
            
            return userName;
        }
        catch
        {
            // If formatting fails, return original username
            return userName;
        }
    }

    /// <summary>
    /// Updates the BOM file with project information using Excel COM automation
    /// </summary>
    /// <param name="filePath">Path to the BOM file</param>
    /// <param name="projectName">Project name</param>
    /// <param name="projectNumber">Project number</param>
    /// <param name="editorName">Editor name</param>
    /// <returns>True if successful, false otherwise</returns>
    private bool UpdateBOMWithProjectInfo(string filePath, string projectName, string projectNumber, string editorName)
    {
        object excelApp = null;
        object workbook = null;
        
        try
        {
            // Create Excel application
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                MessageBox.Show("Excel is not installed on this system.", "Warning", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            excelApp = Activator.CreateInstance(excelType);
            
            // Keep Excel hidden initially - user will choose whether to open it
            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { false });
            
            // Open the workbook
            object workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            workbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath });
            
            // Get the worksheets collection
            object worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null);
            
            // Try to get the worksheet
            object worksheet = null;
            try
            {
                worksheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { WORKSHEET_NAME });
            }
            catch
            {
                MessageBox.Show("Could not find worksheet '" + WORKSHEET_NAME + "' in the BOM file.", "Warning", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            
            // Set project name in designated cell
            object rangeProjectName = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { PROJECT_NAME_CELL });
            rangeProjectName.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, rangeProjectName, new object[] { projectName });
            
            // Set project number in designated cell
            object rangeProjectNumber = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { PROJECT_NUMBER_CELL });
            rangeProjectNumber.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, rangeProjectNumber, new object[] { projectNumber });
            
            // Set editor name in designated cell
            object rangeEditorName = worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, worksheet, new object[] { EDITOR_NAME_CELL });
            rangeEditorName.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, rangeEditorName, new object[] { editorName });
            
            // Save and close the workbook
            workbook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, workbook, null);
            workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
            
            // Quit Excel application
            excelType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
            
            // Release COM objects
            if (rangeProjectName != null) Marshal.ReleaseComObject(rangeProjectName);
            if (rangeProjectNumber != null) Marshal.ReleaseComObject(rangeProjectNumber);
            if (rangeEditorName != null) Marshal.ReleaseComObject(rangeEditorName);
            if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            if (worksheets != null) Marshal.ReleaseComObject(worksheets);
            
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error updating BOM file: " + ex.Message, "Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        finally
        {
            // Clean up COM objects
            if (workbook != null)
            {
                try
                {
                    Marshal.ReleaseComObject(workbook);
                }
                catch { }
            }
            
            if (excelApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(excelApp);
                }
                catch { }
            }
        }
    }
    #endregion
}

/// <summary>
/// Custom dialog with properly named buttons for post-completion actions
/// </summary>
public class CustomActionDialog : Form
{
    #region Fields
    private Button bomButton;
    private Button docButton;
    private Button cancelButton;
    private Label messageLabel;
    #endregion
    
    #region Constructor
    public CustomActionDialog(string message)
    {
        InitializeDialog(message);
    }
    #endregion
    
    #region Private Methods
    /// <summary>
    /// Initializes the dialog components and layout
    /// </summary>
    /// <param name="message">Message to display</param>
    private void InitializeDialog(string message)
    {
        // Form properties
        this.Text = "Success - Choose Action";
        this.Size = new System.Drawing.Size(500, 200);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.TopMost = true;
        
        // Message label
        messageLabel = new Label();
        messageLabel.Text = message + "\n\nWhat would you like to do next?";
        messageLabel.Location = new System.Drawing.Point(20, 20);
        messageLabel.Size = new System.Drawing.Size(450, 100);
        messageLabel.AutoSize = false;
        
        // BOM button
        bomButton = new Button();
        bomButton.Text = "BOM";
        bomButton.Location = new System.Drawing.Point(120, 130);
        bomButton.Size = new System.Drawing.Size(75, 23);
        bomButton.Click += (sender, e) => { this.DialogResult = DialogResult.Yes; this.Close(); };
        
        // DOC button
        docButton = new Button();
        docButton.Text = "DOC";
        docButton.Location = new System.Drawing.Point(205, 130);
        docButton.Size = new System.Drawing.Size(75, 23);
        docButton.Click += (sender, e) => { this.DialogResult = DialogResult.No; this.Close(); };
        
        // Cancel button
        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Location = new System.Drawing.Point(290, 130);
        cancelButton.Size = new System.Drawing.Size(75, 23);
        cancelButton.Click += (sender, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
        
        // Add controls to form
        this.Controls.Add(messageLabel);
        this.Controls.Add(bomButton);
        this.Controls.Add(docButton);
        this.Controls.Add(cancelButton);
        
        // Set default button
        this.AcceptButton = bomButton;
        this.CancelButton = cancelButton;
    }
    #endregion
}
