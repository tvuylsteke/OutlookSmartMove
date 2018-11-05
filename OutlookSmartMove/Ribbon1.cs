using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using Newtonsoft.Json;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.Text.RegularExpressions;
using System.Drawing;

namespace OutlookSmartMove
{
    public partial class Ribbon1
    {
        
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {            
            customerFile = @"c:\temp\customers.json";
            customerFolder = @"\\Thomas.Vuylsteke@microsoft.com\Inbox\FTA";
            using (StreamReader r = new StreamReader(customerFile))
            {
                string json = r.ReadToEnd();
                customers = JsonConvert.DeserializeObject<List<OutlookSmartMove.Customer>>(json);                
            }
            if (customers == null)
            {
                customers = new List<OutlookSmartMove.Customer>();
            }
            moveButton.Enabled = false;
            moveOptions.Enabled = false;            
            folderBox.Text = null;            
        }

        private void writeError(string err)
        {
            MessageBox.Show(err);
        }

        private void updateCustomerInfo(string Folder, string update, string updateType)
        {
            bool newFolder = true;
            OutlookSmartMove.Customer target = new OutlookSmartMove.Customer(Folder);
            
            //check list of known customers to see if the folder is in there already
            foreach (OutlookSmartMove.Customer cust in this.customers)
            {
                if(cust.FolderName == Folder)
                {
                    //work with the customer we already have
                    target = cust;
                    newFolder = false;
                }
            }

            switch (updateType)
            {
                case "email":
                    if (!target.EmailAddresses.Contains(update.ToLower()))
                    {
                        target.EmailAddresses.Add(update.ToLower());
                    }
                    break;
                case "keyword":
                    if (!target.Keywords.Contains(update.ToLower()))
                    {
                        target.Keywords.Add(update.ToLower());
                    }
                    break;
                default:
                    writeError("Error updating customers");
                    break;
            }

            if (newFolder)
            {
                customers.Add(target);
            }
            //write file
            using (StreamWriter file = File.CreateText(customerFile))            
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, customers);                
            }
        }

        private void learnCustomerItem()
        {            
            Outlook.MAPIFolder parentFolder = currentItem.Parent as Outlook.MAPIFolder;
            string currentItemFolder = parentFolder.FolderPath;

            //building a list of suffixes from the current item
            List<string> suffixes = new List<string>();
            foreach (Outlook.Recipient r in currentItem.Recipients)
            {
                if (r.Address.Contains("@"))
                {
                    string suffix = "@" + r.Address.Split('@')[1];
                    if (!suffixes.Contains(suffix))
                    {
                        suffixes.Add(suffix);
                    }
                }
            }

            //for optimization (does it make sense?) we do this as a separate step
            //and only once per email suffix domain
            foreach (string suffix in suffixes)
            {
                updateCustomerInfo(currentItemFolder, suffix, "email");
            }

            //Matching customer name in between the -
            // e.g. Fast Track for Azure - Fabrikam - Follow up and Next Steps
            Regex rx = new Regex(@".*[-](.*)[-].*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            MatchCollection matches = rx.Matches(currentItem.TaskSubject);           

            string customer = "";
            if (matches.Count == 1)
            {
                foreach (Match match in matches)
                {
                    GroupCollection groups = match.Groups;
                    if (groups.Count == 2)
                    {
                        customer = groups[1].Value.Trim();
                        updateCustomerInfo(currentItemFolder, customer, "keyword");
                    }
                }
            }
        }

        private void findCustomerFolder(string subject, string addresses)
        {
            List<string> matches = new List<string>();
            //check each known customer in the config file
            foreach (OutlookSmartMove.Customer cust in customers)
            {
                foreach(string word in cust.Keywords)
                {
                    if (subject.ToLower().Contains(word) && !matches.Contains(cust.FolderName))
                    {
                        matches.Add(cust.FolderName);
                    }
                }

                foreach (string address in cust.EmailAddresses)
                {
                    if (addresses.Contains(address) && !matches.Contains(cust.FolderName))
                    {
                        matches.Add(cust.FolderName);
                    }
                }
            }

            switch (matches.Count)
            {
                case 0:
                    writeError("No matches found");
                    break;
                case 1:
                    folderBox.Text = getFolderShortName(matches.First());
                    folderBox.Tag = matches.First();
                    moveButton.Enabled = true;
                    moveOptions.Enabled = false;
                    break;
                default:
                    moveOptions.Enabled = true;
                    moveButton.Enabled = false;
                    folderBox.Text = null;
                    foreach (string match in matches)
                    {                     
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = getFolderShortName(match);
                        item.Tag = match;
                        moveOptions.Items.Add(item);
                    }
                    writeError("Click Moves");
                    break;                
            }
        }

        //in: "\\\\john.doe@contoso.com\\Inbox\\FTA\\Fabrikam"
        //out: Fabrikam
        private string getFolderShortName(string folderPath)
        {
            string[] res = folderPath.Split('\\');
            return res[res.Length - 1];
        }

        private void initializeCurrentItem()
        {
            this.currentItem = null;
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                object item = explorer.Selection[1];
                if (item is Outlook.MailItem)
                {
                    this.currentItem = item as Outlook.MailItem;
                }
            }            
        }

        //in 
        //  folderPath: "\\\\john.doe@contoso.com\\Inbox\\FTA\\Fabrikam"
        //  folders" starting point to search below.
        //      (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        //out reference to MAPIFolder
        private Outlook.MAPIFolder GetFolder(string folderPath, Outlook.Folders folders)
        {
            if(folders == null)
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                folders = inBox.Folders;
            }
            string dir = folderPath.Substring(0, folderPath.Substring(4).IndexOf("\\") + 4);
            try
            {
                foreach (Outlook.MAPIFolder mf in folders)
                {
                    if (!(mf.FullFolderPath.StartsWith(dir))) continue;
                    if (mf.FullFolderPath == folderPath) return mf;
                    else
                    {
                        Outlook.MAPIFolder temp = GetFolder(folderPath, mf.Folders);
                        if (temp != null)
                            return temp;
                    }
                }
                return null;
            }
            catch { return null; }
        }

        //only searches one level deep!
        private void searchFolder(string query, Outlook.Folders folders)
        {
            if (folders == null)
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                folders = inBox.Folders;
            }           

            Hashtable res = new Hashtable();
            foreach (Outlook.MAPIFolder folder in folders)
            {
                if (folder.Name.ToLower().Contains(query))
                {
                    res[folder.Name] = folder.FolderPath;
                }
            }           

            switch (res.Keys.Count)
            {
                case 0:
                    writeError("Folder not found");
                    //searchButton.Image = new Bitmap(Properties.Resources.foundNot);
                    break;
                case 1:
                    //there's only 1, but not sure if there's a better way to get it.
                    foreach (string key in res.Keys)
                    {
                        folderBox.Text = key;
                        folderBox.Tag = res[key];
                        moveButton.Enabled = true;
                        moveOptions.Enabled = false;
                        //searchButton.Image = new Bitmap(Properties.Resources.foundOk);
                    }
                    break;
                default:
                    moveOptions.Enabled = true;                    
                    moveButton.Enabled = false;
                    moveOptions.Items.Clear();
                    //clearing the field triggers the folderBox_TextChanged option as well!
                    //folderBox.Text = null;
                    //searchButton.Image = new Bitmap(Properties.Resources.foundOk);
                    foreach (string key in res.Keys)
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = key;
                        item.Tag = res[key];
                        moveOptions.Items.Add(item);
                    }
                    break;
            }
        }

        //in 
        //  folder: "Fabrikam"
        //  folders" starting point to create below.
        //      (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        //out reference to MAPIFolder
        private void createFolder(string folder, Outlook.MAPIFolder MAPIfolder)
        {
            Outlook.Folders folders = MAPIfolder.Folders;
            if (folders == null)
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                folders = inBox.Folders;
            }                        
            
            try
            {
                folders.Add(folder, Outlook.OlDefaultFolders.olFolderInbox);
                folderBox.Text = null;
            }
            catch (Exception ex)
            {
                writeError("Error creating folder");
            }          
        }

        //in customerFolder: "\\\\john.doe@contoso.com\\Inbox\\FTA\\Fabrikam"
        private void moveCurrentItem(string customerFolder)
        {            
            try
            {                 
                this.currentItem.Move(GetFolder(customerFolder, null));               
            }
            catch (Exception ex)
            {
                writeError("Bad Folder");
            }            
        }

        private void detectButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();
            folderBox.Text = null;
            moveOptions.Items.Clear();
            
            //convert the list of recipients to one long string which will use to check if the email address sufix is part of it
            string mailaddresses = "";
            foreach(Outlook.Recipient r in currentItem.Recipients)
            {
                mailaddresses += r.Address + ",";                      
            }
            //also include the sender address
            mailaddresses += currentItem.SenderEmailAddress;
            //try to match the mailItem based on it's subject and the string of email addresses
            findCustomerFolder(currentItem.TaskSubject, mailaddresses);            
        }

        //Clicking one of the options, when multiple are provided, will simply set it as active.
        //At the same time it will disable the control.
        private void gallery1_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDownItem item = moveOptions.SelectedItem;
            folderBox.Text = item.Label;
            folderBox.Tag = item.Tag;
            moveOptions.Enabled = false;
            moveButton.Enabled = true;            
        }

        private void moveButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();
            moveCurrentItem((string)folderBox.Tag);            
        }

        private void learnButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();
            learnCustomerItem();               
        }               
   
        private void createButton_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MAPIFolder customerMapiFolder = GetFolder(customerFolder, null);
            createFolder(folderBox.Text, customerMapiFolder);
        }

        private void searchButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();

            //search below the customer folder
            Outlook.MAPIFolder customerFolders = GetFolder(customerFolder, null);
            searchFolder(folderBox.Text.ToLower(), customerFolders.Folders);             
        }

        private void focusButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {                
                Outlook.MAPIFolder destFolder = GetFolder((string)folderBox.Tag, null);
                Globals.ThisAddIn.Application.ActiveExplorer().Activate();
                Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder = destFolder;
            }
            catch(Exception ex)
            {
                writeError("Error focussing");
            }
        }

        private void homeButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);                
                Globals.ThisAddIn.Application.ActiveExplorer().Activate();
                Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder = inBox;
            }
            catch (Exception ex)
            {
                writeError("Error going inbox");
            }
        }

        private void folderBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();

            //search below the customer folder
            Outlook.MAPIFolder customerFolders = GetFolder(customerFolder, null);
            searchFolder(folderBox.Text.ToLower(), customerFolders.Folders);
        }

        private Outlook.MailItem currentItem { get; set; }
        private List<OutlookSmartMove.Customer> customers { get; set; }
        private string customerFile { get; set; }
        private string customerFolder { get; set; }
    }
}
