using MVCWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCWeb.DAL
{
    public class SchoolInitializer :System.Data.Entity.DropCreateDatabaseIfModelChanges<SchoolContext>
    {
        //Seed method taks the database context object as an input parameter,
        protected override void Seed(SchoolContext context)
        {
            var EFstudents = new List<EFStudent>
            {
            new EFStudent{FirstMidName="Carson",LastName="Alexander",EnrollmentDate=DateTime.Parse("2005-09-01")},
            new EFStudent{FirstMidName="Meredith",LastName="Alonso",EnrollmentDate=DateTime.Parse("2002-09-01")},
            new EFStudent{FirstMidName="Arturo",LastName="Anand",EnrollmentDate=DateTime.Parse("2003-09-01")},
            new EFStudent{FirstMidName="Gytis",LastName="Barzdukas",EnrollmentDate=DateTime.Parse("2002-09-01")},
            new EFStudent{FirstMidName="Yan",LastName="Li",EnrollmentDate=DateTime.Parse("2002-09-01")},
            new EFStudent{FirstMidName="Peggy",LastName="Justice",EnrollmentDate=DateTime.Parse("2001-09-01")},
            new EFStudent{FirstMidName="Laura",LastName="Norman",EnrollmentDate=DateTime.Parse("2003-09-01")},
            new EFStudent{FirstMidName="Nino",LastName="Olivetto",EnrollmentDate=DateTime.Parse("2005-09-01")}
            };
            EFstudents.ForEach(s=>context.EFStudents.Add(s));
            context.SaveChanges();
            var EFcourses = new List<EFCourse>
            {
            new EFCourse{EFCourseID=1050,Title="Chemistry",Credits=3,},
            new EFCourse{EFCourseID=4022,Title="Microeconomics",Credits=3,},
            new EFCourse{EFCourseID=4041,Title="Macroeconomics",Credits=3,},
            new EFCourse{EFCourseID=1045,Title="Calculus",Credits=4,},
            new EFCourse{EFCourseID=3141,Title="Trigonometry",Credits=4,},
            new EFCourse{EFCourseID=2021,Title="Composition",Credits=3,},
            new EFCourse{EFCourseID=2042,Title="Literature",Credits=4,}
            };
            EFcourses.ForEach(s=>context.EFCourses.Add(s));
            context.SaveChanges();
            var EFenrollments = new List<EFEnrollment>
            {
            new EFEnrollment{EFStudentID=1,EFCourseID=1050,Grade=Grade.A},
            new EFEnrollment{EFStudentID=1,EFCourseID=4022,Grade=Grade.C},
            new EFEnrollment{EFStudentID=1,EFCourseID=4041,Grade=Grade.B},
            new EFEnrollment{EFStudentID=2,EFCourseID=1045,Grade=Grade.B},
            new EFEnrollment{EFStudentID=2,EFCourseID=3141,Grade=Grade.F},
            new EFEnrollment{EFStudentID=2,EFCourseID=2021,Grade=Grade.F},
            new EFEnrollment{EFStudentID=3,EFCourseID=1050},
            new EFEnrollment{EFStudentID=4,EFCourseID=1050,},
            new EFEnrollment{EFStudentID=4,EFCourseID=4022,Grade=Grade.F},
            new EFEnrollment{EFStudentID=5,EFCourseID=4041,Grade=Grade.C},
            new EFEnrollment{EFStudentID=6,EFCourseID=1045},
            new EFEnrollment{EFStudentID=7,EFCourseID=3141,Grade=Grade.A},
            };
            EFenrollments.ForEach(s=>context.EFEnrollments.Add(s));
            context.SaveChanges();
        }
        

        


        
    }
}