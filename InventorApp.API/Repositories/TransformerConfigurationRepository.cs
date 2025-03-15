using System.Threading.Tasks;
using InventorApp.API.Data;
using InventorApp.API.Models;
using Microsoft.EntityFrameworkCore;

namespace InventorApp.API.Repositories
{
    public class TransformerConfigurationRepository : ITransformerConfigurationRepository
    {
        private readonly ApplicationDbContext _context;

        public TransformerConfigurationRepository(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<TransformerConfiguration?> GetByIdAsync(long projectUniqueId)
        {
            return await _context.TransformerConfigurations.FindAsync(projectUniqueId);
        }

        public async Task<TransformerConfiguration> CreateAsync(TransformerConfiguration configuration)
        {
            _context.TransformerConfigurations.Add(configuration);
            await _context.SaveChangesAsync();
            return configuration;
        }

        public async Task<TransformerConfiguration> UpdateAsync(TransformerConfiguration configuration)
        {
            _context.Entry(configuration).State = EntityState.Modified;
            await _context.SaveChangesAsync();
            return configuration;
        }
    }
} 