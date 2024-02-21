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
using System.Threading;

namespace OutlookOldDefaultThemeAddIn
{
    public partial class ThisAddIn
    {
        private string templatePath = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // https://stackoverflow.com/questions/13728491/opensubkey-returns-null-for-a-registry-key-that-i-can-see-in-regedit-exe
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                {
                    templatePath = hklm
                        .OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration")?
                        .GetValue("InstallationPath") as string;
                }

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
            if (Item != null && Item is Outlook.MailItem) {
                var item = (Item as Outlook.MailItem);

                // We are not 100% happy with doing this for every MailItem,
                // initially we checked beforehand if the mail was sent, but
                // this led to inconsistent COM-Exceptions, which we assume
                // to be from the mail not being fully loaded..
                item.Open += (ref bool cancel) => SetDocumentThemeForMail(item);

                var itemWithEvent = item as ItemEvents_10_Event;
                itemWithEvent.Reply += (object response, ref bool cancel) => FixReadingPaneTheme(response as Outlook.MailItem);
                itemWithEvent.ReplyAll += (object response, ref bool cancel) => FixReadingPaneTheme(response as Outlook.MailItem);
                itemWithEvent.Forward += (object forward, ref bool cancel) => FixReadingPaneTheme(forward as Outlook.MailItem);
            }
        }


        private void FixReadingPaneTheme(Outlook.MailItem Item)
        {
            new Timer((state) =>
            {
                Explorer explorer = this.Application.ActiveExplorer();
                explorer.ActiveInlineResponseWordEditor.ApplyDocumentTheme(templatePath);
            }, null, 1000, Timeout.Infinite);

            return;


            // Checking to only run code if it's a new mail
            if (Item == null || Item.Sent)
                return;

            var inspector = Item.GetInspector;
            if (inspector != null && !string.IsNullOrEmpty(templatePath))
            {
                // This fixes applying the theme in the sidepanel
                // but it introduces this weird behavoir where it opens
                // the same mail in the background as a new window.
                // Since it closes it and keeps it in sync with the
                // sidepanel, it seems to be a necessary evil..
                // Using .Display() instead of .Activate() would work too.
                inspector.Activate();

                // This works to get rid of the extra window, but
                // decided to not include it, since it's probably
                // worse than just keeping a editor window open in
                // the background and letting Outlook handle it.
                /*
                new Timer((state) =>
                {
                    inspector.Close(OlInspectorClose.olDiscard);
                }, null, 500, Timeout.Infinite);
                */
            }
        }


        private void SetDocumentThemeForMail(Outlook.MailItem Item)
        {
            // Checking to only run code if it's a new mail
            if (Item == null || Item.Sent)
                return;

            // We got exceptions in the VBA Prototype
            // but it still set the template, so we just catch
            // any exceptions here, to prevent a bad UX.
            try
            {
                var inspector = Item.GetInspector;
                if (inspector != null && !string.IsNullOrEmpty(templatePath))
                {
                    // This fixes applying the theme in the sidepanel
                    // but it introduces this weird behavoir where it opens
                    // the same mail in the background as a new window.
                    // Since it closes it and keeps it in sync with the
                    // sidepanel, it seems to be a necessary evil..
                    //inspector.Activate();

                    if (inspector.EditorType == OlEditorType.olEditorWord && inspector.WordEditor != null)
                        inspector.WordEditor.ApplyDocumentTheme(templatePath);
                    else if (inspector.EditorType == OlEditorType.olEditorHTML && inspector.HTMLEditor != null)
                        inspector.HTMLEditor.ApplyDocumentTheme(templatePath);
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
