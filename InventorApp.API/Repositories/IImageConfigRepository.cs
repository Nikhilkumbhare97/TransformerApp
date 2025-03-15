using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Models;

namespace InventorApp.API.Repositories
{
    public interface IImageConfigRepository
    {
        Task<ImageConfig?> GetByImageNameAsync(string imageName);
        Task<ImageConfig> SaveAsync(ImageConfig imageConfig);
        Task<List<string>> GetAllImageNamesAsync();
    }
} 