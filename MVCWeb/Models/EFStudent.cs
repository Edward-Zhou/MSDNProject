using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace MVCWeb.Models
{
    public class EFStudent
    {
        public int ID { get; set; }
        [Required]
        [StringLength(50)]
        [Display(Name="Last Name")]
        [RegularExpression(@"^[A-Z]+[a-zA-Z''-'\s]*$")]
        public string LastName { get; set; }
        [Required]
        [Display(Name="First Name")]
        [Column("FirstName")]
        [RegularExpression(@"^[A-Z]+[a-zA-Z''-'\s]*$")]
        [StringLength(50,ErrorMessage="First Name could not be longer than 50")]
        public string FirstMidName { get; set; }
        [DataType(DataType.Date),DisplayFormat(DataFormatString="{0:dd/MM/yyyy}",ApplyFormatInEditMode=true)]
        [Display(Name="Enrollment Date")]
        public DateTime EnrollmentDate { get; set; }
        [Display(Name="Full Name")]
        public string FullName
        {
            get {
                return LastName + "," + FirstMidName;
            }
        }
        //navigation property, hold other entities that are related to this entity
        //the Enrollments property of a student entity will hold all of the Enrollment entities
        //that are related to that students entiry
        public virtual ICollection<EFEnrollment> EFEnrollments { get; set; }
    }
}