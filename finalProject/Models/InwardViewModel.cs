using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace finalProject.Models
{
    public class InwardViewModel
    {
        public decimal InwardNumber { get; set; }
        [DataType(DataType.Date)]
        public DateTime ReceivedOn { get; set; }
        [DataType(DataType.Date)]
        public DateTime LetterDated { get; set; }

        [Required(ErrorMessage = "ReceivedFrom is required")]
        public string ReceivedFrom { get; set; }
       
        public string Subject { get; set; }
        public string FromAddress { get; set; }
        public string FromDistrict { get; set; }
        public string FromState { get; set; }
        public string FromCountry { get; set; }
        public string ReplyFor { get; set; }
        public string ReplyDetails { get; set; }
        public string BrowseFile { get; set; }
        public string Status { get; set; }
        public string Users { get; set; }
    }
}