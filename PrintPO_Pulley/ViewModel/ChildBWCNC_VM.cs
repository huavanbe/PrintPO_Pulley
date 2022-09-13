using BFDataCrawler.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BFDataCrawler.ViewModel
{
    class ChildBWCNC_VM : BaseViewModel
    {
        private string passwordUnlock;
        public string PasswordUnlock
        {
            get { return passwordUnlock; }
            set { passwordUnlock = value; OnPropertyChanged(nameof(PasswordUnlock)); }
        }
    }
}
