using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.IO;

namespace OutlookOldDefaultThemeAddIn
{
    public partial class ThisAddIn
    {
        private string templatePath = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                templatePath = Registry.LocalMachine
                    .OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration")?
                    .GetValue("InstallationPath") as string;

                var targetTemplate = Registry.CurrentUser
                    .OpenSubKey(@"Software\Microsoft\Office\16.0\Common\MailSettings")?
                    .GetValue("OverrideTheme") as string;

                if (string.IsNullOrEmpty(targetTemplate))
                    targetTemplate = "Office 2013 - 2022 Theme.thmx";

                if (!string.IsNullOrEmpty(templatePath))
                {
                    templatePath = Path.Combine(templatePath, "root", "Document Themes 16", targetTemplate);
                    if (!File.Exists(templatePath))
                        templatePath = null;
                }
            }
            catch (System.Exception) { }

            if(templatePath != null)
                this.Application.ItemLoad += Application_ItemLoad;
        }

        private void Application_ItemLoad(object Item)
        {
            if(Item != null && Item is Outlook.MailItem) {
                var item = (Item as Outlook.MailItem);
                if(!item.Sent)
                    item.Open += (ref bool cancel) => MailItem_Open(ref cancel, item);
            }
        }

        private void MailItem_Open(ref bool Cancel, Outlook.MailItem Item)
        {
            // We got exceptions in the VBA Prototype
            // but it still set the template, so we just catch
            // any exceptions here, to prevent a bad UX.
            try
            {
                var inspector = Item.GetInspector;
                if (inspector != null &&
                    inspector.EditorType == OlEditorType.olEditorWord &&
                    inspector.WordEditor != null &&
                    !string.IsNullOrEmpty(templatePath))
                {
                    inspector.WordEditor.ApplyDocumentTheme(templatePath);
                }
            }
            catch (System.Exception) { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            this.Application.ItemLoad -= Application_ItemLoad;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
