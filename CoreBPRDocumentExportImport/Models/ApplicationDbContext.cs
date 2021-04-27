using Microsoft.EntityFrameworkCore;

namespace CoreBPRDocumentExportImport.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<DcxTemplateMaster> DcxTemplateMasters { get; set; }
        public DbSet<DcxTemplateDetail> DcxTemplateDetails { get; set; }

        protected override void OnModelCreating(ModelBuilder builder)
        {
            base.OnModelCreating(builder);


        }
    }
}
