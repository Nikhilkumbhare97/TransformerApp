using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using InventorApp.API.Data;
using InventorApp.API.Models;
using Microsoft.EntityFrameworkCore;

namespace InventorApp.API.Repositories
{
    public class ImageConfigRepository : IImageConfigRepository
    {
        private readonly ApplicationDbContext _context;

        public ImageConfigRepository(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<ImageConfig?> GetByImageNameAsync(string imageName)
        {
            return await _context.ImageConfigs.FindAsync(imageName);
        }

        public async Task<ImageConfig> SaveAsync(ImageConfig imageConfig)
        {
            var existing = await _context.ImageConfigs.FindAsync(imageConfig.ImageName);
            if (existing != null)
            {
                _context.Entry(existing).CurrentValues.SetValues(imageConfig);
            }
            else
            {
                await _context.ImageConfigs.AddAsync(imageConfig);
            }
            await _context.SaveChangesAsync();
            return imageConfig;
        }

        public async Task<List<string>> GetAllImageNamesAsync()
        {
            return await _context.ImageConfigs
                .Select(i => i.ImageName)
                .ToListAsync();
        }
    }
} 