using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Outlook;

namespace EInspectorUC
{
    /// <summary>
    /// Interaction logic for EInspectorUC.xaml
    /// </summary>
    public partial class EInspectorUC : UserControl
    {
        public EInspectorUC()
        {
            InitializeComponent();
        }

        private void TestB_Click(object sender, RoutedEventArgs e)
        {
            // get a reference to our mail item 
            //Outlook.MailItem curMail = (Outlook.MailItem).OutlookItem;
            //HeaderTB.Text = (string)curMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E"); 
        }
    }


}
