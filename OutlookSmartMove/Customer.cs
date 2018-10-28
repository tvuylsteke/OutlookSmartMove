using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookSmartMove
{
    class Customer
    {
        public Customer(string FolderName)
        {
            this.FolderName = FolderName;
            this.EmailAddresses = new List<string>();
            this.Keywords = new List<string>();
        }
        public string FolderName
        {
            get;
            set;
        }
        public List<string> EmailAddresses
        {
            get;
            set;
        }
        public List<string> Keywords
        {
            get;
            set;
        }
    }
}
