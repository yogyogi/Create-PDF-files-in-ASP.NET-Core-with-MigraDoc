using Microsoft.EntityFrameworkCore;

namespace PECTutorial.Models
{
    public class CompanyContext : DbContext
    {
        public CompanyContext(DbContextOptions<CompanyContext> options) : base(options)
        {
        }
        public DbSet<Employee> Employee { get; set; }
    }
}
