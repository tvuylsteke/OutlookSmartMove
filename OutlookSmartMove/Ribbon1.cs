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

namespace OutlookSmartMove
{
    public partial class Ribbon1
    {
        List<OutlookSmartMove.Customer> customers;
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            string customerFile = @"c:\temp\customers.json";   
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

        private void updateCustomers(string Folder, string update, string updateType)
        {
            bool newFolder = true;
            OutlookSmartMove.Customer target = new OutlookSmartMove.Customer(Folder);
            
            foreach (OutlookSmartMove.Customer cust in this.customers)
            {
                if(cust.FolderName == Folder)
                {
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
                    folderBox.Text = "Error updating customers";
                    break;
            }

            if (newFolder)
            {
                this.customers.Add(target);
            }
            //write file
            using (StreamWriter file = File.CreateText(@"c:\temp\customers.json"))            
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, this.customers);                
            }
        }

        private void findCustomerFolder(string subject, string addresses)
        {
            Hashtable matches = new Hashtable();
            //check each known customer in the config file
            foreach (OutlookSmartMove.Customer cust in this.customers)
            {
                foreach(string word in cust.Keywords)
                {
                    if (subject.ToLower().Contains(word))
                    {
                        if (!matches.ContainsKey(cust.FolderName))
                        {
                            matches.Add(cust.FolderName, "");
                        }
                    }
                }

                foreach (string address in cust.EmailAddresses)
                {
                    if (addresses.Contains(address))
                    {
                        if (!matches.ContainsKey(cust.FolderName))
                        {
                            matches.Add(cust.FolderName, "");
                        }
                    }
                }
            }

            if (matches.Count == 0)
            {
                folderBox.Text = "No matches found";
            }
            else if(matches.Count == 1)
            {
                foreach(string key in matches.Keys)
                {
                    folderBox.Text = getFolderShortName(key);
                    folderBox.Tag = key;
                    moveButton.Enabled = true;
                    moveOptions.Enabled = false;
                }

            }
            else { 
                foreach (string key in matches.Keys)
                {
                    moveOptions.Enabled = true;
                    moveButton.Enabled = false;
                    folderBox.Text = null;
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();                    
                    item.Label = getFolderShortName(key);
                    item.Tag = key;
                    moveOptions.Items.Add(item);
                }
                folderBox.Text = "Click Moves";
            }
        }

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
                    //Outlook.MailItem mailItem = item as Outlook.MailItem;
                }
            }            
        }

        private Outlook.MAPIFolder GetFolder(string folderPath, Outlook.Folders folders)
        {
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

        private void moveCurrentItem(string customer)
        {
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //Outlook.MAPIFolder temp = inBox.Folders["FTA"];           

            try
            {
                Outlook.MAPIFolder destFolder = GetFolder(customer, inBox.Folders);
                this.currentItem.Move(destFolder);
                //folderBox.Text = null;
            }
            catch (Exception e)
            {
                folderBox.Text = "Bad Folder";
            }
            //allow multiple moves
            //moveButton.Enabled = false;
            //moveOptions.Enabled = false;
        }

        private void detectButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();
            folderBox.Text = null;
            moveOptions.Items.Clear();         

            Outlook.MailItem mailItem = this.currentItem;           
            string mailaddresses = "";                    

            foreach(Outlook.Recipient r in mailItem.Recipients)
            {
                mailaddresses += r.Address;
                mailaddresses += ",";
                //Debug.WriteLine(r.Address);                        
            }

            mailaddresses += mailItem.SenderEmailAddress;

            findCustomerFolder(mailItem.TaskSubject, mailaddresses);

            //testtje
            //Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //Outlook.MAPIFolder mf = GetFolder(folderBox.Text, inBox.Folders);            
        }

        private void gallery1_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDownItem item = moveOptions.SelectedItem;
            folderBox.Text = item.Label;
            folderBox.Tag = item.Tag;
            moveOptions.Enabled = false;
            //from before when moving right away
            //moveCurrentItem((string)item.Tag);
            //MessageBox.Show("click");
        }

        private Outlook.MailItem currentItem
        {
            get;
            set;
        }

        private void moveButton_Click(object sender, RibbonControlEventArgs e)
        {
            moveCurrentItem((string)folderBox.Tag);            
        }

        private void learnButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();

            Outlook.MailItem mailItem = this.currentItem;
            Outlook.MAPIFolder parentFolder = mailItem.Parent as Outlook.MAPIFolder;
            string FolderLocation = parentFolder.FolderPath;

            Hashtable suffixes = new Hashtable();

            foreach (Outlook.Recipient r in mailItem.Recipients)
            {
                if (r.Address.Contains("@"))
                {
                    string suffix = "@" + r.Address.Split('@')[1];
                    if (!suffixes.ContainsKey(suffix)){
                        suffixes.Add(suffix, "");
                    }
                }                           
            }

            foreach(string suffix in suffixes.Keys)
            {
                updateCustomers(FolderLocation, suffix, "email");
            }

            //Matching customer name in between the -
            Regex rx = new Regex(@".*[-](.*)[-].*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            MatchCollection matches = rx.Matches(mailItem.TaskSubject);

            string customer = "";

            if(matches.Count == 1)
            {
                foreach (Match match in matches)
                {
                    GroupCollection groups = match.Groups;
                    if(groups.Count == 2)
                    {
                        customer = groups[1].Value.Trim();
                        updateCustomers(FolderLocation, customer, "keyword");
                    }                    
                }
            }          

            //ShowMyDialogBox();

            //Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //Outlook.MAPIFolder temp = inBox.Folders["FTA"];
        }               
   
        private void createButton_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder FTA = inBox.Folders["FTA"];

            Outlook.MAPIFolder customFolder = null;
            try
            {
                customFolder = (Outlook.MAPIFolder)FTA.Folders.Add(folderBox.Text, Outlook.OlDefaultFolders.olFolderInbox);
                folderBox.Text = null;
            }
            catch (Exception ex)
            {
                folderBox.Text = "Error creating folder";
            }


        }

        private void searchButton_Click(object sender, RibbonControlEventArgs e)
        {
            initializeCurrentItem();
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder FTA = inBox.Folders["FTA"];
            Hashtable res = new Hashtable();
            foreach (Outlook.MAPIFolder folder in FTA.Folders)
            {
                if (folder.Name.ToLower().Contains(folderBox.Text.ToLower())) {
                    res[folder.Name] = folder.FolderPath;
                }
            }

            if (res.Keys.Count == 0)
            {
                folderBox.Text = "Folder not found";
            }
            else if (res.Keys.Count == 1)
            {
                foreach (string key in res.Keys)
                {
                    folderBox.Text = key;
                    folderBox.Tag = res[key];
                    moveButton.Enabled = true;
                    moveOptions.Enabled = false;
                }
            }
            else
            {
                foreach (string key in res.Keys)
                {
                    moveOptions.Enabled = true;
                    moveButton.Enabled = false;
                    folderBox.Text = null;
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = key;
                    item.Tag = res[key];
                    moveOptions.Items.Add(item);
                }
            }
        }

        private void focusButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)Globals.ThisAddIn.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook.MAPIFolder destFolder = GetFolder((string)folderBox.Tag, inBox.Folders);
                Globals.ThisAddIn.Application.ActiveExplorer().Activate();
                Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder = destFolder;
            }
            catch(Exception ex)
            {
                folderBox.Text = "Error focussing";
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
                folderBox.Text = "Error going inbox";
            }
        }
    }
}
