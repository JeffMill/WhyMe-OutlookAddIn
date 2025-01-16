using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Windows.Forms;

namespace WhyMe
{
    public partial class WhyMeRibbon
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application application = Globals.ThisAddIn.Application;
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem myMailItem = (Outlook.MailItem)inspector.CurrentItem;
            ThisAddIn.Membership found = null;

            foreach (Outlook.Recipient r in myMailItem.Recipients)
            {
                Debug.WriteLine("Searching for: " + r.AddressEntry.Name);
                found = Globals.ThisAddIn.IsMemberOf(r.AddressEntry);
                if (found != null)
                    break;
            }

            if (found != null)
            {
                Debug.WriteLine(found.ToString());
                MessageBox.Show(found.ToString(), "Why Me?", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                Debug.WriteLine("NOT FOUND");
                MessageBox.Show("Couldn't Find you. Perhaps you were BCC'd?", "Why Me?", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
