using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;


namespace dRevealAI
{
    public class EmailWithSummary
    {
        public Outlook.MailItem Mail { get; set; }
        public string Summary { get; set; }
    }
}
