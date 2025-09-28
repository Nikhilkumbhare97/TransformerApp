using Microsoft.EntityFrameworkCore;
using InventorApp.API.Models;

namespace InventorApp.API.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<Project> Projects { get; set; }
        public DbSet<Transformer> Transformers { get; set; }
        public DbSet<TransformerConfiguration> TransformerConfigurations { get; set; }
        public DbSet<ImageConfig> ImageConfigs { get; set; }
        public DbSet<User> Users { get; set; }
    }
} 