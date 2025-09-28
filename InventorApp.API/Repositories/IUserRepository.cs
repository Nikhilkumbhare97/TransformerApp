using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Models;

namespace InventorApp.API.Repositories
{
    public interface IUserRepository
    {
        Task<User> CreateAsync(User user);
        Task<User> UpdateAsync(User user);
        Task<User?> GetByIdAsync(long userUniqueId);
        Task<User?> GetByEmployeeIdAsync(string employeeId);
        Task<User?> GetByEmailAsync(string email);
        Task<IEnumerable<User>> GetAllAsync();
        Task DeleteAsync(long userUniqueId);
    }
}

