using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;

// Created by Mark Spratt
// ExportLabelSelection Extension - Exports Label (Parts List) for selected area and adds to BOM

/* How to use this script:
 * 1. Copy this script into your EPLAN Scripts directory (e.g. C:\EPLAN\Scripts\)
 * 2. Open EPLAN and go to File -> Extras -> Interfaces -> Script: Load
 * 3. Select the script "ExportLabelSelection_Extension.cs" in the file dialog and open
 * 4. Select an area/section in EPLAN (e.g., =S+HS, =R1++MLA1+HSM1, =N1+...)
 * 5. Use the "Add Selection to BOM" button in Tools > Scripts ribbon
 * Features:
 * • Export Label (Parts List) for currently selected area/section
 * • Automatically add to BOM with proper naming
 * • Move sheets to end of BOM workbook
 * • Built-in progress indicator and error handling
 */

/// <summary>
/// Extension to export Label (Parts List) for selected area and add to BOM
/// </summary>
public class ExportLabelSelection_Extension
{
    #region Constants
    private const string CONFIG_SCHEME = "Jensen BOM RevB";
    private const string BOM_FILE_PATTERN = " BOM.xlsm";
    private const string CONTEXT_MENU_DIALOG = "PmPageObjectTreeDialog";
    private const string CONTEXT_MENU_NAME = "1007";
    #endregion

    #region Registration
    [DeclareRegister]
    public void Register()
    {
        MessageBox.Show(
            "ExportLabelSelection Extension script has been loaded successfully!\n\n" +
            "Features added:\n" +
            "• Export Label (Parts List) for selected area/section\n" +
            "• Automatically add to BOM with proper naming\n" +
            "• Move sheets to end of BOM workbook\n" +
            "• Support for multiple selections (run multiple times)\n\n" +
            "Find the 'Add Selection to BOM' option by right-clicking in the Page Navigator",
            "Script Loaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    [DeclareUnregister]
    public void UnRegister()
    {
        MessageBox.Show(
            "ExportLabelSelection Extension script removed successfully!\n\n" +
            "The 'Add Selection to BOM' context menu option has been removed from the Page Navigator.",
            "Script Unloaded",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }
    #endregion

    #region Context Menu Setup
    [DeclareMenu]
    public void SetupContextMenu()
    {
        var contextMenu = new Eplan.EplApi.Gui.ContextMenu();
        var menuLocation = new ContextMenuLocation
        {
            DialogName = CONTEXT_MENU_DIALOG,
            ContextMenuName = CONTEXT_MENU_NAME
        };

        contextMenu.AddMenuItem(menuLocation, "Add Selection to BOM", "AddSelectionToBOM", false, false);
    }
    #endregion

    #region Main Action Methods
    [DeclareAction("AddSelectionToBOM")]
    public void AddSelectionToBOM()
    {
        AutoTreat();
    }

    [Start]
    public void Function()
    {
        AutoTreat();
    }
    public bool AutoTreat()
    {
        Progress progress = new Progress("ExportLabelSelection_Extension");
        progress.SetAllowCancel(false);
        progress.ShowImmediately();
        
        try
        {
            // Step 1: Get project information
            progress.BeginPart(10, "Getting project information...");
            string projectPath = PathMap.SubstitutePath("$(PROJECTPATH)");
            string projectName = PathMap.SubstitutePath("$(PROJECTNAME)");
            
            if (string.IsNullOrEmpty(projectPath) || string.IsNullOrEmpty(projectName))
            {
                progress.EndPart(true);
                MessageBox.Show("No project is currently open.", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            progress.EndPart();

            // Step 2: Get description from user
            progress.BeginPart(15, "Getting export description...");
            string selectionDescription = GetDescriptionFromUser();
            if (string.IsNullOrEmpty(selectionDescription))
            {
                progress.EndPart(true);
                return false; // User cancelled
            }
            progress.EndPart();

            // Step 3: Create DOC folder
            progress.BeginPart(5, "Ensuring DOC folder exists...");
            string docFolderPath = Path.Combine(projectPath, "DOC");
            if (!Directory.Exists(docFolderPath))
            {
                Directory.CreateDirectory(docFolderPath);
            }
            progress.EndPart();

            // Step 4: Determine BOM file path
            progress.BeginPart(10, "Locating BOM file...");
            string bomFilePath = Path.Combine(docFolderPath, projectName + BOM_FILE_PATTERN);
            bool bomExists = File.Exists(bomFilePath);
            progress.EndPart();

            // Step 5: Export Label to temporary file
            progress.BeginPart(30, "Exporting Label (Parts List)...");
            string tempExportFile = @"C:\TEMP\Jensen BOM.xlsx";
            
            // Ensure temp directory exists
            string tempDir = Path.GetDirectoryName(tempExportFile);
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }
            
            bool exportResult = ExportLabelForSelection(tempExportFile);
            if (!exportResult)
            {
                progress.EndPart(true);
                MessageBox.Show("Failed to export Label (Parts List) for the selected area.", "Export Failed", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            progress.EndPart();

            // Step 6: Process BOM file
            progress.BeginPart(30, "Checking BOM file...");
            
            if (!bomExists)
            {
                progress.EndPart(true);
                
                DialogResult result = MessageBox.Show(
                    "No BOM file found in the DOC folder.\n\n" +
                    "You need to create the Master BOM first using the 'Copy Master BOM' button.\n\n" +
                    "Would you like to run the Copy Master BOM script now?\n\n" +
                    "Note: After creating the Master BOM, you will need to rerun this Export Selection script.",
                    "BOM Not Found",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    // Since we can't directly instantiate the other extension, 
                    // just inform the user to run it manually
                    MessageBox.Show(
                        "Please run the 'Copy Master BOM' script manually:\n\n" +
                        "1. Go to Tools > Scripts in the EPLAN ribbon\n" +
                        "2. Click 'Copy Master BOM' button\n" +
                        "3. After it completes, rerun this 'Add Selection to BOM' script\n\n" +
                        "This will create the Master BOM file that this script needs.",
                        "Manual Action Required",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "Please run the 'Copy Master BOM' button first from Tools > Scripts, " +
                        "then rerun this Export Selection script.",
                        "Manual Action Required",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                
                return false;
            }
            
            progress.EndPart();
            progress.BeginPart(30, "Adding to BOM file...");
            
            // Add new sheet to existing BOM
            bool bomProcessResult = AddLabelToBOM(bomFilePath, tempExportFile, selectionDescription);
            
            if (!bomProcessResult)
            {
                progress.EndPart(true);
                MessageBox.Show("Failed to process BOM file. The Label export was successful but could not be added to the BOM.", 
                    "BOM Processing Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            progress.EndPart();

            // Step 7: Success - Ask user what to do next
            string successMessage = bomExists ? 
                "Label (Parts List) for '" + selectionDescription + "' has been successfully added to the existing BOM:\n" + bomFilePath :
                "New BOM created with Label (Parts List) for '" + selectionDescription + "':\n" + bomFilePath;
            
            DialogResult nextAction = MessageBox.Show(
                successMessage + "\n\nWhat would you like to do next?\n\nYes = Open BOM file\nNo = Run script again for new selection\nCancel = Exit",
                "Success - Choose Next Action",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Information);
            
            if (nextAction == DialogResult.Yes)
            {
                // Open BOM file
                try
                {
                    System.Diagnostics.Process.Start(bomFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open BOM file: " + ex.Message, "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (nextAction == DialogResult.No)
            {
                // User wants to run script again - show instruction and exit
                MessageBox.Show("Please select your next area/section in EPLAN and run the script again.", 
                    "Ready for Next Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            // If Cancel, just exit normally
            
            return true;
        }
        catch (Exception ex)
        {
            progress.EndPart(true);
            MessageBox.Show("Error in MasterBOM Helper: " + ex.Message, 
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        finally
        {
            progress.EndPart(true);
        }
    }

    /// <summary>
    /// Processes a single selection export and adds it to the BOM
    /// </summary>
    /// <param name="bomFilePath">Path to the BOM file</param>
    /// <returns>True if successful</returns>
    private bool ProcessSingleSelection(string bomFilePath)
    {
        Progress progress = new Progress("MasterBOM_Helper");
        progress.SetAllowCancel(false);
        progress.ShowImmediately();
        
        try
        {
            // Step 1: Get description from user
            progress.BeginPart(20, "Getting export description...");
            string selectionDescription = GetDescriptionFromUser();
            if (string.IsNullOrEmpty(selectionDescription))
            {
                progress.EndPart(true);
                return false; // User cancelled
            }
            progress.EndPart();

            // Step 2: Ensure temp directory exists and export
            progress.BeginPart(40, "Exporting Label (Parts List)...");
            string tempExportFile = @"C:\TEMP\Jensen BOM.xlsx";
            string tempDir = Path.GetDirectoryName(tempExportFile);
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }
            
            bool exportResult = ExportLabelForSelection(tempExportFile);
            if (!exportResult)
            {
                progress.EndPart(true);
                MessageBox.Show("Failed to export Label (Parts List) for the selected area.", "Export Failed", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            progress.EndPart();

            // Step 3: Move sheet from temp file to BOM
            progress.BeginPart(40, "Adding to BOM file...");
            bool bomExists = File.Exists(bomFilePath);
            bool bomProcessResult = false;
            
            if (bomExists)
            {
                // Add new sheet to existing BOM
                bomProcessResult = AddLabelToBOM(bomFilePath, tempExportFile, selectionDescription);
            }
            else
            {
                // Create new BOM from export
                bomProcessResult = CreateBOMFromExport(tempExportFile, bomFilePath, selectionDescription);
            }
            
            if (!bomProcessResult)
            {
                progress.EndPart(true);
                MessageBox.Show("Failed to process BOM file. The Label export was successful but could not be added to the BOM.", 
                    "BOM Processing Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            progress.EndPart();
            
            return true;
        }
        catch (Exception ex)
        {
            progress.EndPart(true);
            MessageBox.Show("Error processing selection: " + ex.Message, 
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        finally
        {
            progress.EndPart(true);
        }
    }
    #endregion

    #region Helper Methods
    /// <summary>
    /// Gets description from user for the selection
    /// </summary>
    /// <returns>Description for the selection or empty string if cancelled</returns>
    private string GetDescriptionFromUser()
    {
        using (SelectionDescriptionForm form = new SelectionDescriptionForm())
        {
            if (form.ShowDialog() == DialogResult.OK)
            {
                return SanitizeSheetName(form.SelectionDescription);
            }
            return string.Empty; // User cancelled
        }
    }

    /// <summary>
    /// Sanitizes a string to be a valid Excel sheet name
    /// </summary>
    /// <param name="input">Input string</param>
    /// <returns>Sanitized sheet name</returns>
    private string SanitizeSheetName(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "Selection";
        
        // Remove invalid characters for Excel sheet names
        string sanitized = Regex.Replace(input, @"[\\\/\*\?\[\]:]+", "_");
        
        // Limit length (Excel limit is 31 characters)
        if (sanitized.Length > 28) // Leave room for potential numbering
            sanitized = sanitized.Substring(0, 28);
        
        return sanitized;
    }

    /// <summary>
    /// Exports Label (Parts List) for the current selection to a temporary file
    /// </summary>
    /// <param name="tempFilePath">Path for temporary export file</param>
    /// <returns>True if successful</returns>
    private bool ExportLabelForSelection(string tempFilePath)
    {
        try
        {
            // Get the currently selected pages
            CommandLineInterpreter cliCheck = new CommandLineInterpreter();
            ActionCallingContext checkContext = new ActionCallingContext();
            checkContext.AddParameter("TYPE", "PAGES");
            bool hasSelection = cliCheck.Execute("selectionset", checkContext);
            
            string selectedPages = string.Empty;
            if (hasSelection)
            {
                checkContext.GetParameter("PAGES", ref selectedPages);
            }
            
            if (string.IsNullOrEmpty(selectedPages))
            {
                MessageBox.Show(
                    "No pages are currently selected in EPLAN.\n\n" +
                    "To use this script:\n" +
                    "1. Select one or more pages in the Page Navigator\n" +
                    "2. Right-click and choose 'Add Selection to BOM'\n\n" +
                    "The script will export parts from the selected pages only.",
                    "No Selection",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }
            
            // Delete any existing temp file to ensure we start clean
            if (File.Exists(tempFilePath))
            {
                try 
                { 
                    File.Delete(tempFilePath); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Cannot delete existing temporary file: " + tempFilePath + "\n\n" +
                        "Error: " + ex.Message + "\n\n" +
                        "Please close any Excel files that might be using this location.",
                        "File Access Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return false;
                }
            }

            CommandLineInterpreter cli = new CommandLineInterpreter();
            ActionCallingContext context = new ActionCallingContext();
            
            context.AddParameter("configscheme", CONFIG_SCHEME);
            context.AddParameter("filterscheme", "");
            context.AddParameter("sortscheme", "");
            context.AddParameter("language", "en_US");
            context.AddParameter("destinationfile", tempFilePath);
            context.AddParameter("recrepeat", "1");
            context.AddParameter("taskrepeat", "1");
            context.AddParameter("showoutput", "0");
            context.AddParameter("type", "PROJECT");
            context.AddParameter("USESELECTION", "1");
            
            bool result = cli.Execute("label", context);
            
            if (!result)
            {
                MessageBox.Show(
                    "Label export command failed.\n\n" +
                    "This might happen if:\n" +
                    "• No objects are currently selected in EPLAN\n" +
                    "• The selection contains no parts to export\n" +
                    "• The selected objects don't have exportable part data\n" +
                    "• The export configuration scheme '" + CONFIG_SCHEME + "' is not available\n\n" +
                    "Debug info:\n" +
                    "• Attempted export to: " + tempFilePath + "\n" +
                    "• Config scheme: " + CONFIG_SCHEME + "\n\n" +
                    "Please make sure you have selected objects in EPLAN that contain parts/components and try again.\n\n" +
                    "To select objects: Use Ctrl+click or drag to select components, then run this script.",
                    "Export Command Failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
            
            // Verify the export file was actually created and has content
            if (!File.Exists(tempFilePath))
            {
                MessageBox.Show(
                    "Export completed but no output file was created.\n\n" +
                    "This usually means:\n" +
                    "• The selection contains no exportable parts\n" +
                    "• The selected objects don't match the export filter criteria\n" +
                    "• No objects were actually selected when the export ran\n\n" +
                    "Please check your selection includes components/parts and try again.",
                    "No Export Data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }
            
            // Check if the file has meaningful content (more than just headers)
            FileInfo fileInfo = new FileInfo(tempFilePath);
            if (fileInfo.Length < 100) // Arbitrary small size check
            {
                MessageBox.Show(
                    "Export file was created but appears to be empty or contain only headers.\n\n" +
                    "This usually means:\n" +
                    "• The selection contains no exportable parts\n" +
                    "• The selected objects don't have part data to export\n\n" +
                    "Please verify your selection includes components with part numbers and try again.",
                    "Export Contains No Data",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }
            
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error during label export: " + ex.Message, "Export Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    /// <summary>
    /// Adds the exported label data to an existing BOM file as a new sheet
    /// </summary>
    /// <param name="bomFilePath">Path to existing BOM file</param>
    /// <param name="tempExportFile">Path to temporary export file</param>
    /// <param name="selectionDescription">Description for the new sheet name</param>
    /// <returns>True if successful</returns>
    private bool AddLabelToBOM(string bomFilePath, string tempExportFile, string selectionDescription)
    {
        object excelApp = null;
        object bomWorkbook = null;
        object exportWorkbook = null;
        
        try
        {
            // Create Excel application
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                MessageBox.Show("Excel is not installed on this system.", "Excel Required", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            excelApp = Activator.CreateInstance(excelType);
            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { false });
            
            // Open both workbooks
            object workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            bomWorkbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { bomFilePath });
            exportWorkbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { tempExportFile });
            
            // Get worksheets
            object bomWorksheets = bomWorkbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, bomWorkbook, null);
            object exportWorksheets = exportWorkbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, exportWorkbook, null);
            
            // Get the first (and likely only) sheet from the export file
            object exportDataSheet = exportWorksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, exportWorksheets, new object[] { 1 });
            
            // Get the last sheet in the BOM workbook to position the new sheet after it
            int sheetCount = (int)bomWorksheets.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, bomWorksheets, null);
            object lastSheet = bomWorksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, bomWorksheets, new object[] { sheetCount });
            
            // Copy the export sheet to the BOM workbook (after the last sheet)
            exportDataSheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, exportDataSheet, new object[] { Type.Missing, lastSheet });
            
            // Get the newly created sheet (it will be the active sheet)
            object newSheet = bomWorkbook.GetType().InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, bomWorkbook, null);
            
            // Rename the new sheet
            string newSheetName = GetUniqueSheetName(bomWorksheets, selectionDescription);
            newSheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, newSheet, new object[] { newSheetName });
            
            // Close export workbook without saving
            exportWorkbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, exportWorkbook, new object[] { false });
            
            // Save BOM workbook
            bomWorkbook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, bomWorkbook, null);
            bomWorkbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, bomWorkbook, new object[] { false });
            
            // Quit Excel
            excelType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
            
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error adding label to BOM: " + ex.Message, "BOM Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        finally
        {
            // Clean up COM objects
            if (exportWorkbook != null)
            {
                try { Marshal.ReleaseComObject(exportWorkbook); } catch { }
            }
            if (bomWorkbook != null)
            {
                try { Marshal.ReleaseComObject(bomWorkbook); } catch { }
            }
            if (excelApp != null)
            {
                try { Marshal.ReleaseComObject(excelApp); } catch { }
            }
        }
    }

    /// <summary>
    /// Creates a new BOM file from the export data
    /// </summary>
    /// <param name="tempExportFile">Path to temporary export file</param>
    /// <param name="bomFilePath">Path for new BOM file</param>
    /// <param name="selectionDescription">Description for the sheet name</param>
    /// <returns>True if successful</returns>
    private bool CreateBOMFromExport(string tempExportFile, string bomFilePath, string selectionDescription)
    {
        try
        {
            // Copy the export file to the BOM location with proper extension
            File.Copy(tempExportFile, bomFilePath, true);
            
            // Open and rename the sheet to match the selection
            return RenameFirstSheet(bomFilePath, selectionDescription);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error creating BOM from export: " + ex.Message, "BOM Creation Error", 
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    /// <summary>
    /// Renames the first sheet in an Excel file
    /// </summary>
    /// <param name="filePath">Path to Excel file</param>
    /// <param name="newName">New name for the sheet</param>
    /// <returns>True if successful</returns>
    private bool RenameFirstSheet(string filePath, string newName)
    {
        object excelApp = null;
        object workbook = null;
        
        try
        {
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            excelApp = Activator.CreateInstance(excelType);
            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { false });
            
            object workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);
            workbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath });
            
            object worksheets = workbook.GetType().InvokeMember("Worksheets", BindingFlags.GetProperty, null, workbook, null);
            object firstSheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { 1 });
            
            firstSheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, firstSheet, new object[] { newName });
            
            workbook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, workbook, null);
            workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
            excelType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
            
            return true;
        }
        catch
        {
            return false; // Non-critical error, BOM still created
        }
        finally
        {
            if (workbook != null) try { Marshal.ReleaseComObject(workbook); } catch { }
            if (excelApp != null) try { Marshal.ReleaseComObject(excelApp); } catch { }
        }
    }

    /// <summary>
    /// Gets a unique sheet name by checking existing sheets and adding a number if needed
    /// </summary>
    /// <param name="worksheets">Worksheets collection</param>
    /// <param name="baseName">Base name for the sheet</param>
    /// <returns>Unique sheet name</returns>
    private string GetUniqueSheetName(object worksheets, string baseName)
    {
        try
        {
            string proposedName = baseName;
            int counter = 1;
            
            while (SheetExists(worksheets, proposedName))
            {
                proposedName = baseName + "_" + counter;
                counter++;
                
                // Prevent infinite loop
                if (counter > 100)
                {
                    proposedName = baseName + "_" + (DateTime.Now.Ticks % 1000);
                    break;
                }
            }
            
            return proposedName;
        }
        catch
        {
            return baseName + "_" + (DateTime.Now.Ticks % 1000);
        }
    }

    /// <summary>
    /// Checks if a sheet with the given name exists
    /// </summary>
    /// <param name="worksheets">Worksheets collection</param>
    /// <param name="sheetName">Name to check</param>
    /// <returns>True if sheet exists</returns>
    private bool SheetExists(object worksheets, string sheetName)
    {
        try
        {
            object sheet = worksheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, worksheets, new object[] { sheetName });
            return sheet != null;
        }
        catch
        {
            return false; // Sheet doesn't exist
        }
    }
    #endregion
}



/// <summary>
/// Form to get selection description from user
/// </summary>
public class SelectionDescriptionForm : Form
{
    #region Fields
    private TextBox selectionTextBox;
    private Button okButton;
    private Button cancelButton;
    private Label instructionLabel;
    #endregion
    
    #region Properties
    public string SelectionDescription { get; private set; }
    #endregion
    
    #region Constructor
    public SelectionDescriptionForm()
    {
        InitializeForm();
    }
    #endregion
    
    #region Private Methods
    /// <summary>
    /// Initializes the form components and layout
    /// </summary>
    private void InitializeForm()
    {
        // Form properties
        this.Text = "Selection Description";
        this.Size = new System.Drawing.Size(400, 180);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.TopMost = true;
        
        // Instruction label
        instructionLabel = new Label();
        instructionLabel.Text = "Enter a name for the Excel tab that will be added to the BOM:\n(e.g., HS, MLA1, Motor Control, R1_Panel)";
        instructionLabel.Location = new System.Drawing.Point(20, 20);
        instructionLabel.Size = new System.Drawing.Size(350, 40);
        instructionLabel.AutoSize = false;
        
        // Selection text box
        selectionTextBox = new TextBox();
        selectionTextBox.Location = new System.Drawing.Point(20, 70);
        selectionTextBox.Size = new System.Drawing.Size(350, 23);
        selectionTextBox.Text = "Selection";
        selectionTextBox.SelectAll();
        
        // OK button
        okButton = new Button();
        okButton.Text = "OK";
        okButton.Location = new System.Drawing.Point(220, 110);
        okButton.Size = new System.Drawing.Size(75, 23);
        okButton.Click += OkButton_Click;
        
        // Cancel button
        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Location = new System.Drawing.Point(305, 110);
        cancelButton.Size = new System.Drawing.Size(75, 23);
        cancelButton.Click += CancelButton_Click;
        
        // Add controls to form
        this.Controls.Add(instructionLabel);
        this.Controls.Add(selectionTextBox);
        this.Controls.Add(okButton);
        this.Controls.Add(cancelButton);
        
        // Set default button and focus
        this.AcceptButton = okButton;
        this.CancelButton = cancelButton;
        selectionTextBox.Focus();
    }
    
    private void OkButton_Click(object sender, EventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(selectionTextBox.Text))
        {
            SelectionDescription = selectionTextBox.Text.Trim();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        else
        {
            MessageBox.Show("Please enter a description for the selection.", "Input Required", 
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            selectionTextBox.Focus();
        }
    }
    
    private void CancelButton_Click(object sender, EventArgs e)
    {
        this.DialogResult = DialogResult.Cancel;
        this.Close();
    }
    #endregion
}
