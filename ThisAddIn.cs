using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;

namespace WhyMe
{
    public partial class ThisAddIn
    {
        private Outlook.AddressEntries addrEntries;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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

        public class Membership
        {
            public bool IsDirect { get; set; }
            public IList<string> MemberOfDLs { get; set; }
            public override string ToString()
            {
                if (IsDirect)
                {
                    return "Mail was sent directly To/CC you.";
                }
                else
                {
                    return "Found you in " + string.Join("\r\n-> ", MemberOfDLs);
                }
            }
        }
        public Membership IsMemberOf(Outlook.AddressEntry groupname)
        {
            Membership ret = null;
            try
            {
            Outlook.Application a = new Outlook.Application();
            Outlook.AddressEntry currentUser = a.Session.CurrentUser.AddressEntry;
                if (currentUser.Type == "EX")
                {
                    Outlook.ExchangeUser exchUser = currentUser.GetExchangeUser();
                    if (exchUser != null)
                    {
                        if (addrEntries == null)
                            addrEntries = exchUser.GetMemberOfList();
                        if (addrEntries != null)
                        {
                            if (groupname.Name.Equals(currentUser.Name))    //User is on To/CC
                            {
                                ret = new Membership();
                                ret.IsDirect = true;
                                return ret;
                                // return "Mail was sent directly To/CC you.";
                            }
                            else
                            {
                                Debug.WriteLine("Searching for " + groupname.Name + " in your groups...");
                                foreach (Outlook.AddressEntry addrEntry in addrEntries) //See if this group is in the user's groups list
                                {
                                    if (addrEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                                        if (addrEntry.Name.Equals(groupname.Name))
                                        {
                                            ret = new Membership();
                                            ret.MemberOfDLs = new List<string>();
                                            ret.MemberOfDLs.Add(addrEntry.Name);
                                            return ret;
                                            // return "Found you in " + addrEntry.Name;
                                        }
                                }
                                if (groupname.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)    //If not, recurse into member-DLs
                                {
                                    Outlook.ExchangeDistributionList exchDL = groupname.GetExchangeDistributionList();
                                    Outlook.AddressEntries subAddrEntries = exchDL.GetExchangeDistributionListMembers();
                                    if (subAddrEntries != null)
                                    {
                                        foreach (Outlook.AddressEntry exchDLMember in subAddrEntries)
                                        {
                                            if (exchDLMember.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                                            {
                                                Debug.WriteLine("Attempting to recurse into " + exchDLMember.Name + "(child of " + groupname.Name + ")...");
                                                ret = IsMemberOf(exchDLMember);
                                                if (ret != null)
                                                {
                                                    ret.MemberOfDLs.Add(groupname.Name);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("Exception: " + e);
            }
            return ret;
        }
    }
}
