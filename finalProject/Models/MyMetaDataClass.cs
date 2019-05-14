using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace finalProject.Models
{
    public class AdminLoginMetaData
    {
        [Required(ErrorMessage = "Email ID/UserName is required")]
        [Display(Name = "Email ID")]
        [DataType(DataType.EmailAddress)]
        public String UserName { set; get; }

        [Required(ErrorMessage = "Password is required")]
        [Display(Name = "Password")]
        [DataType(DataType.Password)]
        public String Password { set; get; }
    }

    [MetadataType(typeof(AdminLoginMetaData))]
    public partial class AdminLogin
    {


    }


    public class UserRegistrationMetaData
    {
        [Required(ErrorMessage = "Staff code is required")]
        [Display(Name = "Staff Code")]      
        public String StaffCode { set; get; }

        [Required(ErrorMessage = "Staff Name is required")]
        [Display(Name = "Staff Name")]
        public String StaffName { set; get; }

        [Required(ErrorMessage = "Email ID/UserName is required")]
        [Display(Name = "Email ID")]
        [DataType(DataType.EmailAddress)]
        //[RegularExpression("^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$")]
        public String EmailId { set; get; }

        

        [Required(ErrorMessage = "Contact Number is required")]
        [Display(Name = "Contact Number")]
        //[RegularExpression("^\\d{10}$", ErrorMessage = "Contact Number should consists of 10 digits")]
        [RegularExpression("^[7-9][0-9]{9}$", ErrorMessage ="Start with 7,8,9")]
        public String ContactNumber { set; get; }

        [Required(ErrorMessage = "Department is required")]
        [Display(Name = "Department")]
        public String Department { set; get; }

        [Required(ErrorMessage = "Password is required")]
        [Display(Name = "Password")]
        [DataType(DataType.Password)]
        public String Password { set; get; }

        [Required(ErrorMessage = "Confirm Password is required")]
        [Compare("Password", ErrorMessage = "Password Mismatch..")]
        [Display(Name = "Confirm Password")]
        [DataType(DataType.Password)]
        public String ConfirmPassword { set; get; }

        [Display(Name = "Role")]
        public String Role { set; get; }

    }

    [MetadataType(typeof(UserRegistrationMetaData))]
    public partial class UserRegistration
    {


    }
    public class Inward_RegisterMetaData
    {
        /* [Required(ErrorMessage ="Field is Required")]
         public string ReceivedFrom { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string Subject { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string FromAddress { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string FromDistrict { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string FromState { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string FromCountry { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string ReplyFor { get; set; }

         [Required(ErrorMessage = "Field is Required")]
         public string ReplyDetails { get; set; }
        
        [Required(ErrorMessage = "Please select file.")]
        [RegularExpression(@"([a-zA-Z0-9\s_\\.\-:])+(.png|.jpg|.gif)$", ErrorMessage = "Only Image files allowed.")]
        public string BrowseFile { get; set; }
         */

    }

    [MetadataType(typeof(Inward_RegisterMetaData))]
    public partial class Inward_Register
    {

    }


    public class Outward_Register2MetaData
    {
      /*  [Required(ErrorMessage = "Field is Required")]
        public string SendTo { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string Subject { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ToAddress { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ToDistrict { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ToState { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ToCountry { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ReplyFor { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ReplyDetails { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string BrowseFile { get; set; }

        [DataType(DataType.Date)]
        public DateTime SendOn { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string DispatchType { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string ReceiptFile { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public string DispatchDetails { get; set; }

        [Required(ErrorMessage = "Field is Required")]
        public Nullable<decimal> DispatchCharges { get; set; }
        */

    }

    [MetadataType(typeof(Outward_Register2MetaData))]
    public partial class Outward_Register2
    {

    }

}