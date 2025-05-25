using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Models;
using InventorApp.API.Repositories;
using Microsoft.EntityFrameworkCore;

namespace InventorApp.API.Services
{
    public class ProjectService
    {
        private readonly IProjectRepository _projectRepository;

        public ProjectService(IProjectRepository projectRepository)
        {
            _projectRepository = projectRepository;
        }

        public async Task<Project> CreateProject(Project project)
        {
            if (string.IsNullOrEmpty(project.Status))
            {
                project.Status = "Active";
            }
            return await _projectRepository.CreateAsync(project);
        }

        public async Task DeleteProject(long projectUniqueId)
        {
            await _projectRepository.UpdateProjectStatusAsync(projectUniqueId, "Deleted");
        }

        public async Task<Project> UpdateProject(long projectUniqueId, Project partialProject)
        {
            var existingProject = await _projectRepository.GetByIdAsync(projectUniqueId);
            if (existingProject == null)
            {
                throw new Exception($"Project not found with id: {projectUniqueId}");
            }

            // Update only non-null fields
            if (!string.IsNullOrEmpty(partialProject.ProjectName))
                existingProject.ProjectName = partialProject.ProjectName;
            if (!string.IsNullOrEmpty(partialProject.ProjectNumber))
                existingProject.ProjectNumber = partialProject.ProjectNumber;
            if (!string.IsNullOrEmpty(partialProject.ProjectId))
                existingProject.ProjectId = partialProject.ProjectId;
            if (!string.IsNullOrEmpty(partialProject.ClientName))
                existingProject.ClientName = partialProject.ClientName;
            if (!string.IsNullOrEmpty(partialProject.CreatedBy))
                existingProject.CreatedBy = partialProject.CreatedBy;
            if (!string.IsNullOrEmpty(partialProject.PreparedBy))
                existingProject.PreparedBy = partialProject.PreparedBy;
            if (!string.IsNullOrEmpty(partialProject.CheckedBy))
                existingProject.CheckedBy = partialProject.CheckedBy;
            if (partialProject.Date != DateTime.MinValue)
                existingProject.Date = partialProject.Date;

            return await _projectRepository.UpdateAsync(existingProject);
        }

        public async Task<IEnumerable<Project>> GetAllProjects()
        {
            return await _projectRepository.GetAllAsync();
        }

        public async Task<IEnumerable<Project>> GetAllProjectsByStatus(IEnumerable<string> status)
        {
            return await _projectRepository.FindByStatusInAsync(status);
        }

        public async Task<Project> GetProjectById(long projectUniqueId)
        {
            return await _projectRepository.GetByIdAsync(projectUniqueId);
        }

        public async Task<bool> IsProjectNumberExists(string projectNumber)
        {
            return await _projectRepository.ExistsByProjectNumberAsync(projectNumber);
        }

        public async Task<(IEnumerable<Project> Projects, int TotalCount)> GetPagedProjects(
            int page, int pageSize, string search, string sortBy, string sortDirection, IEnumerable<string> status)
        {
            var query = _projectRepository.Query(); // Expose IQueryable from repository

            // Filter by status
            if (status != null && status.Any())
                query = query.Where(p => status.Contains(p.Status));

            // Search
            if (!string.IsNullOrEmpty(search))
            {
                query = query.Where(p =>
                    p.ProjectName.Contains(search) ||
                    p.ProjectNumber.Contains(search) ||
                    p.ProjectId.Contains(search) ||
                    p.ClientName.Contains(search) ||
                    p.Status.Contains(search)
                );
            }

            // Sorting
            if (!string.IsNullOrEmpty(sortBy))
            {
                bool desc = sortDirection?.ToLower() == "desc";
                query = sortBy switch
                {
                    "projectName" => desc ? query.OrderByDescending(p => p.ProjectName) : query.OrderBy(p => p.ProjectName),
                    "projectNumber" => desc ? query.OrderByDescending(p => p.ProjectNumber) : query.OrderBy(p => p.ProjectNumber),
                    "projectId" => desc ? query.OrderByDescending(p => p.ProjectId) : query.OrderBy(p => p.ProjectId),
                    "clientName" => desc ? query.OrderByDescending(p => p.ClientName) : query.OrderBy(p => p.ClientName),
                    "status" => desc ? query.OrderByDescending(p => p.Status) : query.OrderBy(p => p.Status),
                    _ => query.OrderBy(p => p.ProjectName)
                };
            }

            int totalCount = await query.CountAsync();
            var projects = await query.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();

            return (projects, totalCount);
        }
    }
}