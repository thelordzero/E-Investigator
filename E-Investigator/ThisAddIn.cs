using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace E_Investigator
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddACategory("Clean", OlCategoryColor.olCategoryColorGreen, OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF3);
            AddACategory("Spam", OlCategoryColor.olCategoryColorYellow, OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF4);
            AddACategory("Possible Malicious", OlCategoryColor.olCategoryColorDarkOrange, OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF5);
            AddACategory("Verified Malicious", OlCategoryColor.olCategoryColorDarkRed, OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF6);
            AddACategory("Possible Targeted", OlCategoryColor.olCategoryColorGray, OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF7);
            AddACategory("Verified Targeted", OlCategoryColor.olCategoryColorRed, OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF8);

            //AddACategory("Malicious", OlCategoryColor.olCategoryColorRed, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            //AddACategory("Reported to MS", OlCategoryColor.olCategoryColorOrange, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            //AddACategory("Unsubscribe", OlCategoryColor.olCategoryColorGreen, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            //AddACategory("Manual Block/Other", OlCategoryColor.olCategoryColorYellow, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            //AddACategory("Targeted", OlCategoryColor.olCategoryColorPurple, OlCategoryShortcutKey.olCategoryShortcutKeyNone);
            //AddACategory("No Further Action", OlCategoryColor.olCategoryColorBlue, OlCategoryShortcutKey.olCategoryShortcutKeyNone);

            //currentExplorer = this.ActiveExplorer();
            //currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);

            this.Application.NewMail += new Outlook.ApplicationEvents_11_NewMailEventHandler(ThisAddIn_NewMail);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void ThisAddIn_NewMail()
        {
            Outlook.NameSpace outlookNameSpace = this.Application.GetNamespace("MAPI");

            Outlook.MAPIFolder inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Items unreadMailItems = inbox.Items.Restrict("[Unread]= true");

            foreach (Object omailItem in unreadMailItems)
            {
                Outlook.MailItem unreadMailItem = omailItem as Outlook.MailItem;

                if (unreadMailItem != null)
                {
                    if (unreadMailItem.To == "lordzero")
                    {
                        string existingCategories = unreadMailItem.Categories;
                        if (String.IsNullOrEmpty(existingCategories))
                        {
                            unreadMailItem.Categories = "SPAM";
                        }
                        else
                        {
                            if (unreadMailItem.Categories.Contains("SPAM") == false)
                            {
                                unreadMailItem.Categories = existingCategories + ", SPAM";
                            }
                        }
                        unreadMailItem.Save();
                    }
                }
            }
        }

        /// <summary>
        /// Used for verifying categories during debug stage.
        /// </summary>
        private void EnumerateCategories()
        {
            Outlook.Categories categories = Application.Session.Categories;
            foreach (Outlook.Category category in categories)
            {
                Debug.WriteLine(category.Name);
                Debug.WriteLine(category.CategoryID);
            }
        }

        private void AddACategory(string categoryName, Outlook.OlCategoryColor categoryColor, Outlook.OlCategoryShortcutKey categoryShortcut)
        {
            Outlook.Categories categories = Application.Session.Categories;
            if (!CategoryExists(categoryName))
            {
                Outlook.Category category = categories.Add(categoryName, categoryColor, categoryShortcut);
            }
        }

        private bool CategoryExists(string categoryName)
        {
            try
            {
                Outlook.Category category = Application.Session.Categories[categoryName];
                if (category != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch { return false; }
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
