using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace E_Investigator
{
    public partial class DefaultRibbon
    {
        private void DefaultRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        #region Button Actions
        /// <summary>
        /// Categorizes selected message(s) as spam.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bSpam_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleItemCategory("Spam");
        }

        /// <summary>
        /// Categorizes selected message(s) as being possibly malicious.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bPossMal_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleItemCategory("Possible Malicious");
        }

        /// <summary>
        /// Categorizes selected message(s) as being verified malicious.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bVerMal_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleItemCategory("Verified Malicious");
        }

        /// <summary>
        /// Categorizes selected message(s) as possible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bPoss_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleItemCategory("Possible Targeted");
        }

        /// <summary>
        /// Categorizes selected message(s) as verified.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bVer_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleItemCategory("Verified Targeted");
        }

        /// <summary>
        /// Categorizes selected message(s) as having been inspected and clean.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bClean_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleItemCategory("Clean");
        }

        /// <summary>
        /// Display information regarding message(s).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bInspect_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Inspecting");
        }

        /// <summary>
        /// Executes a Forefront Online Protection for Exchange (FOPE) search for selected message(s).
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bFOPESearch_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Searching");
        }
        #endregion

        private void ToggleItemCategory(string category)
        {
            Inspector inspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (inspector != null)
            {
                if (inspector.CurrentItem != null)
                {
                    if (inspector.CurrentItem is MailItem)
                    {
                        MailItem _mailItem = inspector.CurrentItem;

                        if (_mailItem.Categories != null)
                        {
                            if (_mailItem.Categories.Contains(category))
                                _mailItem.Categories = _mailItem.Categories.Replace(string.Format("{0}, ", category), "").Replace(string.Format("{0}", category), "");
                            else
                                _mailItem.Categories = string.Format("{0}, {1}", category, _mailItem.Categories);
                        }
                        else
                            _mailItem.Categories = category;
                    }
                }
            }
        }

        private List<Microsoft.Office.Interop.Outlook.Categories> GetCategories()
        {
            Outlook.NameSpace nameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            List<Outlook.Categories> catList = null;

            foreach (Outlook.Categories cat in nameSpace.Categories)
            {
                catList.Add(cat);
            }
            return catList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subjectEmail">Subject of the email.</param>
        /// <param name="toEmail">Who the email is to (semi-colon delimited).</param>
        /// <param name="bodyEmail">Body of the email.</param>
        /// <param name="attachment">Boolean that determines if selected object is attached to the email.</param>
        private void CreateEmailItem(string subjectEmail, string toEmail, string bodyEmail, bool attachment)
        {
            Outlook.MailItem eMail = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            ((Outlook._MailItem)eMail).Send();
        }        
    }
}
