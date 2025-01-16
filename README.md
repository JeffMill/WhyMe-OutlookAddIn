# WhyMe-OutlookAddIn

Outlook AddIn to determine why a receiver got an email.

Uses:

[Microsoft.Office.Tools.Outlook.OutlookAddInBase](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.outlook.outlookaddinbase?view=vsto-2022) to implement Outlook Add-In.

[Microsoft.Office.Tools.Ribbon.RibbonBase](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.ribbon.ribbonbase?view=vsto-2022) for the Ribbon.

[Microsoft.Office.Outlook.Application](https://learn.microsoft.com/en-us/office/vba/api/Outlook.Application) to get Outlook application instance.

[Microsoft.Office.Outlook.ExchangeUser](https://learn.microsoft.com/en-us/office/vba/api/Outlook.ExchangeUser) to get [AddressEntry](https://learn.microsoft.com/en-us/office/vba/api/outlook.addressentry) for an Exchange user.

[Microsoft.Office.Outlook.ExchangeDistributionList](https://learn.microsoft.com/en-us/office/vba/api/Outlook.ExchangeDistributionList) to get [AddressEntry](https://learn.microsoft.com/en-us/office/vba/api/outlook.addressentry) for an Exchange distribution list.

[Microsoft.Office.AddressEntries](https://learn.microsoft.com/en-us/office/vba/api/outlook.addressentries) as collection of AddressEntry objects.