using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibrary
{
    public class WorkbookOptions
    {
        private bool includeHidden = false;

        public bool IncludeHidden
        {
            get { return this.includeHidden; }
            set { this.includeHidden = value; }
        }
    }
}
