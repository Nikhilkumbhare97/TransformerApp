using System;
using System.Threading.Tasks;
using InventorApp.API.Models;
using InventorApp.API.Data;
using Microsoft.EntityFrameworkCore;

namespace InventorApp.API.Repositories
{
    public class TransformerRepository : ITransformerRepository
    {
        private readonly ApplicationDbContext _context;

        public TransformerRepository(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<Transformer?> GetByIdAsync(long projectUniqueId)
        {
            return await _context.Transformers.FindAsync(projectUniqueId);
        }

        public async Task<Transformer> CreateAsync(Transformer transformer)
        {
            _context.Transformers.Add(transformer);
            await _context.SaveChangesAsync();
            return transformer;
        }

        public async Task<Transformer> UpdateAsync(Transformer transformer)
        {
            _context.Entry(transformer).State = EntityState.Modified;
            await _context.SaveChangesAsync();
            return transformer;
        }
    }
} 