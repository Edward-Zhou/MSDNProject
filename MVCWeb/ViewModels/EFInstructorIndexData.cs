using MVCWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCWeb.ViewModels
{
    public class EFInstructorIndexData
    {
        public IEnumerable<EFInstructor> EFInstructors { get; set; }
        public IEnumerable<EFCourse> EFCourses { get; set; }
        public IEnumerable<EFEnrollment> EFEnrollments { get; set; }
    }
}