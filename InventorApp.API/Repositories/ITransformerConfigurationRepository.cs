using System.Threading.Tasks;
using InventorApp.API.Models;

namespace InventorApp.API.Repositories
{
    public interface ITransformerConfigurationRepository
    {
        Task<TransformerConfiguration?> GetByIdAsync(long projectUniqueId);
        Task<TransformerConfiguration> CreateAsync(TransformerConfiguration configuration);
        Task<TransformerConfiguration> UpdateAsync(TransformerConfiguration configuration);
    }
} 