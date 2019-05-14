using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace finalProject.Models
{
    public class OutwardViewModel
    {
        public decimal OutwardNumber { get; set; }
        [DataType(DataType.Date)]
        public DateTime LetterDated { get; set; }
        public string SendTo { get; set; }
        public string Subject { get; set; }
        public string ToAddress { get; set; }
        public string ToDistrict { get; set; }
        public string ToState { get; set; }
        public string ToCountry { get; set; }
        public string ReplyFor { get; set; }
        public string ReplyDetails { get; set; }
        public string BrowseFile { get; set; }
        [DataType(DataType.Date)]
        public DateTime SendOn { get; set; }
        public string DispatchType { get; set; }
        public string ReceiptFile { get; set; }
        public string DispatchDetails { get; set; }
        public Nullable<decimal> DispatchCharges { get; set; }
        public string Status { get; set; }
    }
}