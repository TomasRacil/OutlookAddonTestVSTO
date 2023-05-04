using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace OutlookAddonTestVSTO
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(InboxItems_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }

        private void InboxItems_ItemAdd(object Item)
        {
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mailItem = Item as Outlook.MailItem;

                // Kontrola, jestli je email utajovaný
                //if (mailItem.Sensitivity == Outlook.OlSensitivity.olConfidential)
                //{
                    // Kontrola, zda uživatel splňuje zadané požadavky
                    if (!MeetsRequirements(mailItem))
                    {
                        mailItem.UnRead = true; // Nastaví email jako nepřečtený
                        mailItem.Subject = "NEPŘEČTENO: " + mailItem.Subject; // Změní předmět emailu
                        // Zobrazení zprávy uživateli, že email byl označen jako nepřečtený
                        MessageBox.Show("Tento email byl označen jako nepřečtený.", "Upozornění", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                //}
            }
        }

        private bool MeetsRequirements(Outlook.MailItem mailItem)
        {

            // Zpráva bude zablokována, když předmět bude obsahovat klíčové slovo
            if (mailItem.Subject.Contains("SLOVO"))
            {
                return true; // Zpráva splňuje požadavky pro blokování
            }
            else
            {
                return false; // Zpráva nesplňuje požadavky pro blokování
            }
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
