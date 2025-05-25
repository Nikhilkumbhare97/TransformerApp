using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using InventorApp.API.Models;
using Microsoft.EntityFrameworkCore;
using InventorApp.API.Data;
using System;

namespace InventorApp.API.Repositories
{
    public class ProjectRepository : IProjectRepository
    {
        private readonly ApplicationDbContext _context;

        public ProjectRepository(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IEnumerable<Project>> GetAllAsync()
        {
            return await _context.Projects.ToListAsync();
        }

        public async Task<Project> GetByIdAsync(long projectUniqueId)
        {
            var project = await _context.Projects.FindAsync(projectUniqueId);
            return project ?? throw new Exception($"Project not found with id: {projectUniqueId}");
        }

        public async Task<Project> CreateAsync(Project project)
        {
            _context.Projects.Add(project);
            await _context.SaveChangesAsync();
            return project;
        }

        public async Task<Project> UpdateAsync(Project project)
        {
            _context.Entry(project).State = EntityState.Modified;
            await _context.SaveChangesAsync();
            return project;
        }

        public async Task UpdateProjectStatusAsync(long projectUniqueId, string status)
        {
            var project = await _context.Projects.FindAsync(projectUniqueId);
            if (project != null)
            {
                project.Status = status;
                await _context.SaveChangesAsync();
            }
        }

        public async Task<bool> ExistsByProjectNumberAsync(string projectNumber)
        {
            return await _context.Projects.AnyAsync(p => p.ProjectNumber == projectNumber);
        }

        public async Task<IEnumerable<Project>> FindByStatusInAsync(IEnumerable<string> statuses)
        {
            return await _context.Projects.Where(p => statuses.Contains(p.Status)).ToListAsync();
        }

        public IQueryable<Project> Query() // New method
        {
            return _context.Projects;
        }
    }
}