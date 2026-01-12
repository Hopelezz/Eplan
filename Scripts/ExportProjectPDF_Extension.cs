using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Eplan.EplApi.Base;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Gui;

public class ExportProjectPDF_Extension
{
    #region Constants
    private const string TAB_NAME_EN = "Tools";
    private const string TAB_NAME_DE = "Extras";
    private const string GROUP_NAME = "Scripts";
    #endregion

    #region Registration
    [DeclareRegister]
    public void Register()
    {
        try
        {
            SetupRibbonInterface();
            
            MessageBox.Show(
                "Export PDF Rev script has been loaded successfully!\n\n" +
                "Find the 'Export PDF Rev' button in Tools > Scripts",
                "Script Loaded",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                string.Format("Error during registration: {0}\n\nStack trace:\n{1}", ex.Message, ex.StackTrace),
                "Registration Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    [DeclareUnregister]
    public void UnRegister()
    {
        try
        {
            RemoveRibbonInterface();
            
            MessageBox.Show(
                "Export PDF Rev script removed successfully!",
                "Script Unloaded",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception)
        {
            // Silent catch during unregistration to avoid interfering with EPLAN shutdown
            // or other scripts' unregistration process
        }
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
    /// Sets up the ribbon interface with the Export PDF Rev button
    /// </summary>
    private void SetupRibbonInterface()
    {
        try
        {
            RibbonBar ribbonBar = new RibbonBar();

            // Try to find existing Tools tab
            RibbonTab toolsTab = null;
            
            // Try English first
            MultiLangString tabNameEN = new MultiLangString();
            tabNameEN.AddString(ISOCode.Language.L_en_US, TAB_NAME_EN);
            toolsTab = ribbonBar.GetTab(tabNameEN, false);
            
            if (toolsTab == null)
            {
                // Try German
                MultiLangString tabNameDE = new MultiLangString();
                tabNameDE.AddString(ISOCode.Language.L_de_DE, TAB_NAME_DE);
                toolsTab = ribbonBar.GetTab(tabNameDE, false);
            }

            if (toolsTab == null)
            {
                MessageBox.Show("Tools tab not found. Please make sure EPLAN is fully loaded.", "Warning");
                return;
            }

            // Try to find existing Scripts group
            RibbonCommandGroup scriptsGroup = toolsTab.GetCommandGroup(GROUP_NAME);
            
            if (scriptsGroup == null)
            {
                // Create Scripts group if it doesn't exist
                scriptsGroup = toolsTab.AddCommandGroup(GROUP_NAME);
            }

            if (scriptsGroup == null)
            {
                MessageBox.Show("Failed to find or create Scripts group in Tools tab", "Error");
                return;
            }

            // Add our command to the existing/created Scripts group
            scriptsGroup.AddCommand("Export PDF Rev", "ExportPDFRev", new RibbonIcon(CommandIcon.Application));
        }
        catch (Exception ex)
        {
            MessageBox.Show(string.Format("Ribbon setup error: {0}", ex.Message), "Error");
        }
    }

    /// <summary>
    /// Clean unregistration - EPLAN will handle command cleanup automatically
    /// </summary>
    private void RemoveRibbonInterface()
    {
        // EPLAN automatically cleans up commands when scripts are unloaded
        // No manual removal needed - just ensure clean shutdown
    }
    #endregion

    #region Main Action Methods
    [DeclareAction("ExportPDFRev")]
    public void Function()
    {
        try
        {
            // Get the project name using EPLAN's PathMap - same as MasterBOM extension
            string projectName = PathMap.SubstitutePath("$(PROJECTNAME)");
            
            if (string.IsNullOrEmpty(projectName))
            {
                MessageBox.Show("No project is currently open.", "Error");
                return;
            }

            // Get job number for revision extraction
            CommandLineInterpreter cli = new CommandLineInterpreter();
            ActionCallingContext jobCtx = new ActionCallingContext();
            jobCtx.AddParameter("PropertyId", "10013"); // Job number property ID
            jobCtx.AddParameter("PropertyIndex", "0");
            
            cli.Execute("XEsGetProjectPropertyAction", jobCtx);
            
            string jobNumber = string.Empty;
            jobCtx.GetParameter("PropertyValue", ref jobNumber);
            
            // Extract revision from job number
            string revision = ExtractRevision(jobNumber);
            if (string.IsNullOrEmpty(revision))
            {
                // Prompt user to enter revision name
                string userRevision = PromptForRevision();
                if (string.IsNullOrEmpty(userRevision))
                {
                    // User cancelled or entered empty string
                    return;
                }
                revision = userRevision;
            }

            // Create PDF filename: ProjectName + Revision
            string pdfFileName = string.Format("{0} {1}", projectName.Trim(), revision.Trim());

            // Ask user if they want to generate reports before PDF export
            DialogResult generateReportsResult = MessageBox.Show(
                "Do you want to generate project reports before exporting the PDF?\n\n" +
                "• Yes: Generate reports first, then export PDF\n" +
                "• No: Export PDF only (faster)\n" +
                "• Cancel: Cancel the operation",
                "Generate Reports?",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);
            
            if (generateReportsResult == DialogResult.Cancel)
            {
                return; // User cancelled the operation
            }
            
            if (generateReportsResult == DialogResult.Yes)
            {
                // Generate project reports before PDF export
                bool reportsGenerated = GenerateProjectReports();
                if (!reportsGenerated)
                {
                    DialogResult continueResult = MessageBox.Show(
                        "Project reports generation failed or was cancelled.\n\nDo you want to continue with PDF export anyway?",
                        "Reports Generation Issue",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    
                    if (continueResult == DialogResult.No)
                    {
                        return;
                    }
                }
            }

            // Export PDF
            ExportProjectToPDF(pdfFileName);
        }
        catch (Exception ex)
        {
            MessageBox.Show(string.Format("An error occurred: {0}", ex.Message), "Error");
        }
    }

    /// <summary>
    /// Prompts the user to enter a revision name
    /// </summary>
    /// <returns>The revision name entered by the user, or empty string if cancelled</returns>
    private string PromptForRevision()
    {
        try
        {
            // Create a simple input form
            using (Form inputForm = new Form())
            {
                inputForm.Text = "Enter Revision";
                inputForm.Width = 400;
                inputForm.Height = 150;
                inputForm.StartPosition = FormStartPosition.CenterScreen;
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputForm.MaximizeBox = false;
                inputForm.MinimizeBox = false;

                Label label = new Label();
                label.Text = "No revision found in job number.\nPlease enter the revision name:";
                label.Location = new System.Drawing.Point(20, 20);
                label.Width = 350;
                label.Height = 40;
                inputForm.Controls.Add(label);

                TextBox textBox = new TextBox();
                textBox.Text = "RevA";
                textBox.Location = new System.Drawing.Point(20, 65);
                textBox.Width = 250;
                textBox.SelectAll();
                inputForm.Controls.Add(textBox);

                Button okButton = new Button();
                okButton.Text = "OK";
                okButton.Location = new System.Drawing.Point(280, 63);
                okButton.Width = 75;
                okButton.DialogResult = DialogResult.OK;
                inputForm.Controls.Add(okButton);

                Button cancelButton = new Button();
                cancelButton.Text = "Cancel";
                cancelButton.Location = new System.Drawing.Point(280, 90);
                cancelButton.Width = 75;
                cancelButton.DialogResult = DialogResult.Cancel;
                inputForm.Controls.Add(cancelButton);

                inputForm.AcceptButton = okButton;
                inputForm.CancelButton = cancelButton;

                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    return textBox.Text != null ? textBox.Text.Trim() : string.Empty;
                }
                else
                {
                    return string.Empty; // User cancelled
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                string.Format("Error getting revision input: {0}", ex.Message),
                "Input Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return string.Empty;
        }
    }

    private string ExtractRevision(string jobNumber)
    {
        // Check for null or empty input
        if (string.IsNullOrEmpty(jobNumber))
        {
            return string.Empty;
        }

        try
        {
            // Assuming the revision is always at the end of the job number
            // Example: "F7000 RevA" -> "RevA"
            string[] parts = jobNumber.Split(' ');
            if (parts != null && parts.Length > 1)
            {
                string lastPart = parts[parts.Length - 1];
                return !string.IsNullOrEmpty(lastPart) ? lastPart : string.Empty;
            }
            return string.Empty;
        }
        catch (Exception)
        {
            // If any error occurs during string processing, return empty string
            return string.Empty;
        }
    }

    /// <summary>
    /// Generates project reports using EPLAN CLI command
    /// </summary>
    /// <returns>True if successful, false otherwise</returns>
    private bool GenerateProjectReports()
    {
        try
        {
            // Create progress indicator
            Progress progress = new Progress("Generate Reports");
            progress.SetAllowCancel(false);
            progress.BeginPart(100, "Generating project reports...");
            progress.ShowImmediately();

            CommandLineInterpreter cli = new CommandLineInterpreter();
            
            // Execute the reports command for the currently opened project
            // This will generate all configured project reports
            cli.Execute("reports /TYPE:PROJECT");
            
            progress.EndPart(true);
            
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                string.Format("Project reports generation failed: {0}", ex.Message),
                "Reports Generation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return false;
        }
    }

    /// <summary>
    /// Exports the current project to PDF using the specified filename
    /// </summary>
    /// <param name="fileName">The filename for the PDF (without .pdf extension)</param>
    private void ExportProjectToPDF(string fileName)
    {
        try
        {
            // Get the DOC folder path
            string docPath = PathMap.SubstitutePath("$(DOC)");
            
            // Create full file path
            string fullFilePath = Path.Combine(docPath, fileName);
            string fullFilePathWithExt = fullFilePath + ".pdf";
            
            // Check if file already exists
            if (File.Exists(fullFilePathWithExt))
            {
                DialogResult overwriteResult = MessageBox.Show(
                    string.Format("The PDF file already exists:\n{0}\n\nDo you want to overwrite it?", fullFilePathWithExt),
                    "File Already Exists",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (overwriteResult == DialogResult.No)
                {
                    return; // Exit without exporting
                }
                
                // Delete the existing file so EPLAN can create a new one
                try
                {
                    File.Delete(fullFilePathWithExt);
                    
                    // Wait a moment to ensure file system updates
                    System.Threading.Thread.Sleep(100);
                    
                    // Verify file was actually deleted
                    if (File.Exists(fullFilePathWithExt))
                    {
                        MessageBox.Show("File deletion failed - file still exists. Export cancelled.", "Error");
                        return;
                    }
                }
                catch (Exception deleteEx)
                {
                    MessageBox.Show(
                        string.Format("Could not delete existing file: {0}\n\nPossible causes:\n• The PDF file is currently open in a PDF viewer\n• The file is being previewed in Windows Explorer\n• The file is being used by another application\n\nPlease close the file and try again.", deleteEx.Message),
                        "Delete Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
            }
            
            // Create progress indicator
            Progress progress = new Progress("PDF Export");
            progress.SetAllowCancel(false);
            progress.BeginPart(100, "Exporting PDF...");
            progress.ShowImmediately();

            CommandLineInterpreter cli = new CommandLineInterpreter();
            ActionCallingContext ctx = new ActionCallingContext();

            // Set the correct parameters for PDF export
            ctx.AddParameter("TYPE", "PDFPROJECTSCHEME");
            ctx.AddParameter("EXPORTFILE", fullFilePath);
            ctx.AddParameter("EXPORTSCHEME", "EPLAN_default_value");
            
            // Execute the export action
            cli.Execute("export", ctx);
            
            progress.EndPart(true);
            
            // Verify the export actually created the file
            if (File.Exists(fullFilePathWithExt))
            {
                ShowExportCompleteDialog(fileName, docPath, fullFilePathWithExt);
            }
            else
            {
                MessageBox.Show(
                    string.Format("Export command completed but no PDF file was created.\n\nExpected location: {0}", fullFilePathWithExt),
                    "Export Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                string.Format("PDF export failed: {0}", ex.Message),
                "Export Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    /// <summary>
    /// Shows the export complete dialog with options to open folder, copy to clipboard, or close
    /// </summary>
    /// <param name="fileName">The PDF filename without extension</param>
    /// <param name="docPath">The DOC folder path</param>
    /// <param name="fullFilePathWithExt">The complete file path including .pdf extension</param>
    private void ShowExportCompleteDialog(string fileName, string docPath, string fullFilePathWithExt)
    {
        try
        {
            using (Form dialog = new Form())
            {
                dialog.Text = "Export Complete";
                dialog.Width = 480;
                dialog.Height = 200;
                dialog.StartPosition = FormStartPosition.CenterScreen;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;

                Label messageLabel = new Label();
                messageLabel.Text = string.Format("PDF export completed successfully!\n\nFile: {0}.pdf\nLocation: {1}", fileName, docPath);
                messageLabel.Location = new System.Drawing.Point(20, 20);
                messageLabel.Width = 420;
                messageLabel.Height = 80;
                dialog.Controls.Add(messageLabel);

                Button openFolderButton = new Button();
                openFolderButton.Text = "Open Folder";
                openFolderButton.Location = new System.Drawing.Point(20, 120);
                openFolderButton.Width = 100;
                openFolderButton.Click += (sender, e) => {
                    try
                    {
                        System.Diagnostics.Process.Start("explorer.exe", docPath);
                    }
                    catch (Exception folderEx)
                    {
                        MessageBox.Show(
                            string.Format("Could not open folder: {0}", folderEx.Message),
                            "Folder Open Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    }
                };
                dialog.Controls.Add(openFolderButton);

                Button copyToClipboardButton = new Button();
                copyToClipboardButton.Text = "Copy to Clipboard";
                copyToClipboardButton.Location = new System.Drawing.Point(140, 120);
                copyToClipboardButton.Width = 120;
                copyToClipboardButton.Click += (sender, e) => {
                    try
                    {
                        System.Collections.Specialized.StringCollection files = new System.Collections.Specialized.StringCollection();
                        files.Add(fullFilePathWithExt);
                        Clipboard.SetFileDropList(files);
                        
                        MessageBox.Show(
                            string.Format("PDF file copied to clipboard!\n\nYou can now paste it into the contract folder."),
                            "Copied to Clipboard",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    }
                    catch (Exception clipboardEx)
                    {
                        MessageBox.Show(
                            string.Format("Could not copy file to clipboard: {0}", clipboardEx.Message),
                            "Clipboard Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    }
                };
                dialog.Controls.Add(copyToClipboardButton);

                Button closeButton = new Button();
                closeButton.Text = "Close";
                closeButton.Location = new System.Drawing.Point(280, 120);
                closeButton.Width = 75;
                closeButton.DialogResult = DialogResult.OK;
                dialog.Controls.Add(closeButton);

                dialog.AcceptButton = closeButton;
                dialog.ShowDialog();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                string.Format("Error showing export dialog: {0}", ex.Message),
                "Dialog Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }
    #endregion
}
