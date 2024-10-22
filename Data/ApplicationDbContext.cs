using ExcelImport.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelImport.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext>options) : base(options)
        {
            
        }

        public DbSet<Student> students { get; set; }
    }
}
