using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelFingerPrint.Models
{
    public class HomeViewModel
    {
        public List<FingerPrintData> FingerPrintData { get; set; }
        public List<string> ListGuestID { get; set; }
        public List<string> ListEntryDoor { get; set; }
    }
}