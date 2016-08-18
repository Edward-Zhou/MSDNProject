using MVCWeb.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Linq;
using System.Web;

namespace MVCWeb.DAL
{
    public class SchoolContext:DbContext
    {
        //base("SchoolContext") specify the connection string in web.config
        //if you do not specify a connection string or the name of one explicitly, EF assume that the
        //connection string name is the same as class name
        public SchoolContext()
            : base("SchoolContext")
        {
            //this.Configuration.LazyLoadingEnabled = false;
        }

        //DbSet property for each entity set. In EF, an entity set typically corresponds
        //to a database table, and an entity corresponds to a row in the table.
        public DbSet<EFStudent> EFStudents { get; set; }
        public DbSet<EFEnrollment> EFEnrollments { get; set; }
        public DbSet<EFCourse> EFCourses { get; set; }
        public DbSet<EFDepartment> EFDepartments { get; set; }

        public DbSet<EFInstructor> EFInstructors { get; set; }
        public DbSet<EFOfficeAssignment> EFOfficeAssignments { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            modelBuilder.Entity<EFCourse>()
                .HasMany(c => c.EFInstructors).WithMany(i => i.EFCourses)
                .Map(t => t.MapLeftKey("EFCourseID")
                .MapRightKey("EFInstructorID")
                .ToTable("EFCourseEFInstructor"));

        }
        public DbSet<File> Files { get; set; }

        public System.Data.Entity.DbSet<MVCWeb.Models.Person> People { get; set; }
       
    }
}