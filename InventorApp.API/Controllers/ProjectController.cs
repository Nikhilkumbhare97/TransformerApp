using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using InventorApp.API.Models;
using InventorApp.API.Services;
using Microsoft.AspNetCore.Cors;

namespace InventorApp.API.Controllers
{
    [Route("api/projects")]
    [ApiController]
    [EnableCors("AllowSpecificOrigin")]
    public class ProjectController : ControllerBase
    {
        private readonly ProjectService _projectService;

        public ProjectController(ProjectService projectService)
        {
            _projectService = projectService;
        }

        [HttpPost]
        public async Task<ActionResult<Project>> CreateProject([FromBody] Project project)
        {
            var createdProject = await _projectService.CreateProject(project);
            return CreatedAtAction(nameof(GetProjectById), new { projectUniqueId = createdProject.ProjectUniqueId }, createdProject);
        }

        [HttpDelete("{projectUniqueId}")]
        public async Task<IActionResult> DeleteProject(long projectUniqueId)
        {
            await _projectService.DeleteProject(projectUniqueId);
            return NoContent();
        }

        [HttpPut("{projectUniqueId}")]
        public async Task<ActionResult<Project>> UpdateProject(long projectUniqueId, [FromBody] Project partialProject)
        {
            var updatedProject = await _projectService.UpdateProject(projectUniqueId, partialProject);
            return Ok(updatedProject);
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Project>>> GetAllProjects()
        {
            var projects = await _projectService.GetAllProjects();
            return Ok(projects);
        }

        [HttpGet("{projectUniqueId}")]
        public async Task<ActionResult<Project>> GetProjectById(long projectUniqueId)
        {
            var project = await _projectService.GetProjectById(projectUniqueId);
            if (project == null)
            {
                return NotFound();
            }
            return Ok(project);
        }

        [HttpGet("exists/{projectNumber}")]
        public async Task<ActionResult<bool>> CheckProjectNumberExists(string projectNumber)
        {
            var exists = await _projectService.IsProjectNumberExists(projectNumber);
            return Ok(exists);
        }

        [HttpGet("status")]
        public async Task<ActionResult<IEnumerable<Project>>> GetProjectsByStatus([FromQuery] List<string> status)
        {
            var projects = await _projectService.GetAllProjectsByStatus(status);
            return Ok(projects);
        }

        [HttpGet("paged")]
        public async Task<ActionResult<object>> GetPagedProjects(
    [FromQuery] int page = 1,
    [FromQuery] int pageSize = 10,
    [FromQuery] string search = "",
    [FromQuery] string sortBy = "projectName",
    [FromQuery] string sortDirection = "asc",
    [FromQuery] List<string>? status = null)
        {
#pragma warning disable CS8604 // Possible null reference argument.
            var (projects, totalCount) = await _projectService.GetPagedProjects(page, pageSize, search, sortBy, sortDirection, status);
#pragma warning restore CS8604 // Possible null reference argument.
            return Ok(new { projects, totalCount });
        }
    }
}