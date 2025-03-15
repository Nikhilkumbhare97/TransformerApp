using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Models;

namespace InventorApp.API.Repositories
{
    public interface IProjectRepository
    {
        Task<IEnumerable<Project>> GetAllAsync();
        Task<Project> GetByIdAsync(long projectUniqueId);
        Task<Project> CreateAsync(Project project);
        Task<Project> UpdateAsync(Project project);
        Task UpdateProjectStatusAsync(long projectUniqueId, string status);
        Task<bool> ExistsByProjectNumberAsync(string projectNumber);
        Task<IEnumerable<Project>> FindByStatusInAsync(IEnumerable<string> statuses);
    }
} 