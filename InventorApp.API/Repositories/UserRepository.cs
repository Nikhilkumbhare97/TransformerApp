using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using InventorApp.API.Data;
using InventorApp.API.Models;
using Microsoft.EntityFrameworkCore;

namespace InventorApp.API.Repositories
{
    public class UserRepository : IUserRepository
    {
        private readonly ApplicationDbContext _context;

        public UserRepository(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<User> CreateAsync(User user)
        {
            _context.Users.Add(user);
            await _context.SaveChangesAsync();
            return user;
        }

        public async Task<User> UpdateAsync(User user)
        {
            _context.Entry(user).State = EntityState.Modified;
            await _context.SaveChangesAsync();
            return user;
        }

        public async Task<User?> GetByIdAsync(long userUniqueId)
        {
            return await _context.Users.FindAsync(userUniqueId);
        }

        public async Task<User?> GetByEmployeeIdAsync(string employeeId)
        {
            return await _context.Users.FirstOrDefaultAsync(u => u.EmployeeId == employeeId);
        }

        public async Task<User?> GetByEmailAsync(string email)
        {
            return await _context.Users.FirstOrDefaultAsync(u => u.Email == email);
        }

        public async Task<IEnumerable<User>> GetAllAsync()
        {
            return await _context.Users.ToListAsync();
        }

        public async Task DeleteAsync(long userUniqueId)
        {
            var user = await _context.Users.FindAsync(userUniqueId);
            if (user != null)
            {
                _context.Users.Remove(user);
                await _context.SaveChangesAsync();
            }
        }
    }
}

