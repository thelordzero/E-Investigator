using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

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
            System.Windows.Forms.MessageBox.Show("Spam");
        }

        /// <summary>
        /// Categorizes selected message(s) as being possibly malicious.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bPossMal_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Possible Malicious");
        }

        /// <summary>
        /// Categorizes selected message(s) as being verified malicious.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bVerMal_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Verified Malicious");
        }

        /// <summary>
        /// Categorizes selected message(s) as possible.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bPoss_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Possible");
        }

        /// <summary>
        /// Categorizes selected message(s) as verified.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bVer_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Verified");
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

        private List<Microsoft.Office.Interop.Outlook.Categories> GetCategories()
        {
            Microsoft.Office.Interop.Outlook.NameSpace nameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            List<Microsoft.Office.Interop.Outlook.Categories> catList = null;

            foreach (Microsoft.Office.Interop.Outlook.Categories cat in nameSpace.Categories)
            {
                catList.Add(cat);
            }
            return catList;
        }
    }
}
