using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSmartMove
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void detectButton_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                        
            Outlook.MAPIFolder temp = inBox.Folders["FTA"];
            Outlook.MAPIFolder destFolder = temp.Folders["_Test"];

            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                object item = explorer.Selection[1];
                if (item is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = item as Outlook.MailItem;
                    mailItem.Move(destFolder);
                }
            }
        }
    }
}
