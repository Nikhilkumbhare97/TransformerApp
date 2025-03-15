using System.Threading.Tasks;
using InventorApp.API.Models;

namespace InventorApp.API.Repositories
{
    public interface ITransformerRepository
    {
        Task<Transformer?> GetByIdAsync(long projectUniqueId);
        Task<Transformer> CreateAsync(Transformer transformer);
        Task<Transformer> UpdateAsync(Transformer transformer);
    }
} 