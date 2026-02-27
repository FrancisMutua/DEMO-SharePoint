using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DEMO_SharePoint.Models
{
    public class UserModel
    {
        public int Id { get; set; }
        [Required(ErrorMessage = "Username is required")]
        [Display(Name = "User Name")]
        public string Username { get; set; }
        [Required(ErrorMessage = "Password is required")]
        [DataType(DataType.Password)]
        [Display(Name = "Password")]
        public string Password { get; set; }
        public string Email { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Phone { get; set; }
        public string Department { get; set; }

        // Search result fields (populated by AD/SP user search)
        public string DisplayName { get; set; }
        public string Login { get; set; }
    }
}