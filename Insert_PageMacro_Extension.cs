using System;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
using Eplan.EplApi.Scripting;

// Created by Mark Spratt

/* How to use this script:
 * 1. Copy this script into your EPLAN Scripts directory (e.g. C:\EPLAN\Scripts\)
 * 2. Open EPLAN and go to File -> Extras -> Interfaces -> Script: Load
 * 3. Select the script "Insert_X67_Extension.cs" in the file dialog and open
 * 4. Click "OK" in the Register Script Dialog if it's already registered (this will update the script)
 * 5. A confirmation message will appear once the script is loaded successfully
  *
 * Eventually, I want to increase the macros added dynamically.
 */

/// <summary>
/// Extension to add context menu options for inserting macros
/// Right-click on a page in the Page Navigator to see the options. These are specific to common insertable pages.
/// </summary>
public class InsertMacroContextMenu
{
    private const string CONTEXT_MENU_DIALOG = "PmPageObjectTreeDialog";
    private const string CONTEXT_MENU_NAME = "1007";

    private const string MacroPath1 = "Y:\\EPLAN\\Macros\\STD\\N\\Modules\\B_X67DM1321_DIGITAL_MIXED.emp";
    private const string MacroName1 = "X67DM1321";

    private const string MacroPath2 = "Y:\\EPLAN\\Macros\\STD\\N\\Modules\\A_X67PS1300_POWER_SUPPLY.emp";
    private const string MacroName2 = "X67PS1300";

    private const string MacroPath3 = "Y:\\EPLAN\\Macros\\STD\\N\\Modules\\C_X67UM1352_ANALOG_INPUT.emp";
    private const string MacroName3 = "X67UM1352";

    [DeclareMenu]
    public void SetupContextMenu()
    {
        var contextMenu = new Eplan.EplApi.Gui.ContextMenu();
        var menuLocation = new ContextMenuLocation
        {
            DialogName = CONTEXT_MENU_DIALOG,
            ContextMenuName = CONTEXT_MENU_NAME
        };

        // Add menu items for each macro
        contextMenu.AddMenuItem(menuLocation, "Insert " + MacroName1, "InsertPageMacro1", false, false);
        contextMenu.AddMenuItem(menuLocation, "Insert " + MacroName2, "InsertPageMacro2", false, false);
        contextMenu.AddMenuItem(menuLocation, "Insert " + MacroName3, "InsertPageMacro3", false, false);
    }

    [DeclareAction("InsertPageMacro1")]
    public void InsertPageMacro1(ActionCallingContext context)
    {
        InsertMacro(MacroPath1, MacroName1);
    }

    [DeclareAction("InsertPageMacro2")]
    public void InsertPageMacro2(ActionCallingContext context)
    {
        InsertMacro(MacroPath2, MacroName2);
    }

    [DeclareAction("InsertPageMacro3")]
    public void InsertPageMacro3(ActionCallingContext context)
    {
        InsertMacro(MacroPath3, MacroName3);
    }

    private void InsertMacro(string macroPath, string macroName)
    {
        try
        {
            ActionCallingContext oAcc = new ActionCallingContext();
            CommandLineInterpreter oCLI = new CommandLineInterpreter();

            oAcc.AddParameter("filename", macroPath);
            oAcc.AddParameter("pagename", macroName);
            oAcc.AddParameter("variant", "0"); // Variant A
            oAcc.AddParameter("RepresentationType", "1"); // MultiLine
            oAcc.AddParameter("AskForNumeration", "false");
            oAcc.AddParameter("NumberModus", "1");
            oAcc.AddParameter("RenumberPrefixes", "false");

            oCLI.Execute("XMInsertPageMacro", oAcc);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error inserting macro: " + ex.Message);
        }
    }
}