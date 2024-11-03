using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class SearchResultData
    {
        public string Url { get;set; }

        public string TicketNumber { get;set; }

        public string TicketId 
        { 
            get
            {
                return Url.Substring(Url.Length - 32);
            }
        }

        public string StartTime { get;set; }

        public string EndTime { get;set; }

        public string CreatedAt { get;set; } 
    }
}
