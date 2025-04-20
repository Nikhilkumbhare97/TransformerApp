using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Services;
using System.Text.Json.Serialization;
using InventorApp.API.Models;

namespace InventorAPI.Controllers
{
    [ApiController]
    [Route("api/inventor")]
    public class AssemblyController : ControllerBase
    {
        private readonly AssemblyService _assemblyService;

        public AssemblyController(AssemblyService assemblyService)
        {
            _assemblyService = assemblyService;
        }

        [HttpPost("open-assembly")]
        public IActionResult OpenAssembly([FromBody] AssemblyRequest request)
        {
            _assemblyService.OpenAssembly(request.AssemblyPath);
            return Ok(new { message = "Assembly opened successfully." });
        }

        [HttpPost("close-assembly")]
        public IActionResult CloseAssembly()
        {
            _assemblyService.CloseAssembly();
            return Ok(new { message = "Assembly closed successfully." });
        }

        [HttpPost("change-parameters")]
        public IActionResult ChangeParameters([FromBody] ParameterRequest request)
        {
            _assemblyService.ChangeParameters(request.PartFilePath, request.Parameters);
            return Ok(new { message = "Parameters updated successfully." });
        }

        [HttpPost("suppress-component")]
        public IActionResult SuppressComponent([FromBody] SuppressComponentRequest request)
        {
            _assemblyService.SuppressComponent(request.AssemblyFilePath, request.ComponentName, request.Suppress);
            return Ok(new { message = $"Component {request.ComponentName} {(request.Suppress ? "suppressed" : "unsuppressed")} successfully." });
        }

        [HttpGet("assembly-status")]
        public IActionResult GetAssemblyStatus()
        {
            return Ok(new { isAssemblyOpen = _assemblyService.IsAssemblyOpen });
        }

        [HttpPost("suppress-multiple-components")]
        public IActionResult SuppressMultipleComponents([FromBody] SuppressMultipleRequest request)
        {
            if (request.SuppressActions == null || request.SuppressActions.Count == 0)
            {
                return BadRequest("Invalid request: 'suppressActions' must be a non-empty list.");
            }

            _assemblyService.SuppressMultipleComponents(request.SuppressActions);
            return Ok(new { message = "Multiple components updated successfully." });
        }


        [HttpPost("update-all-properties")]
        public IActionResult UpdateAllProperties([FromBody] UpdateAllPropertiesRequest request)
        {
            try
            {
                if (string.IsNullOrEmpty(request.DirectoryPath) || request.IProperties == null || request.IProperties.Count == 0)
                {
                    return BadRequest(new { message = "Invalid request: directoryPath and iProperties are required." });
                }

                bool success = _assemblyService.UpdateIPropertiesForAllFiles(request.DirectoryPath, request.IProperties);
                return success
                    ? Ok(new { message = "iProperties updated successfully for all assemblies and parts." })
                    : StatusCode(500, new { message = "Failed to update iProperties. Check the application logs for details." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"An error occurred: {ex.Message}" });
            }
        }

        [HttpPost("update-multiple-iparts-iassemblies")]
        public IActionResult UpdateIpartsAndIassemblies([FromBody] UpdateIpartsRequest request)
        {
            if (request.AssemblyUpdates == null || request.AssemblyUpdates.Count == 0)
            {
                return BadRequest("Invalid request: assemblyUpdates cannot be empty.");
            }

            bool success = _assemblyService.UpdateIpartsAndIassemblies(request.AssemblyUpdates);
            return Ok(new { message = "Iparts Iassemblies updated successfully." });
        }

        [HttpPost("update-model-state-and-representations")]
        public IActionResult UpdateModelStateAndRepresentations([FromBody] ModelStateUpdateRequest request)
        {
            if (request.AssemblyUpdates == null || request.AssemblyUpdates.Count == 0)
            {
                return BadRequest("Invalid request: assemblyUpdates cannot be empty.");
            }

            bool success = _assemblyService.UpdateModelStateAndRepresentations(request.AssemblyUpdates);
            return success
                ? Ok(new { message = "Model states and representations updated successfully." })
                : StatusCode(500, "Failed to update model states and representations.");
        }
    }

    public class AssemblyRequest { public string AssemblyPath { get; set; } = "C:\\path\\to\\your\\assembly.iam"; }
    public class ParameterRequest { public string PartFilePath { get; set; } = "C:\\path\\to\\your\\part.ipt"; public List<Dictionary<string, object>> Parameters { get; set; } = new(); }
    public class SuppressComponentRequest { public string AssemblyFilePath { get; set; } = "C:\\path\\to\\your\\assembly.iam"; public string ComponentName { get; set; } = "ComponentName"; public bool Suppress { get; set; } = true; }
    public class SuppressMultipleRequest
    {
        [JsonPropertyName("suppressActions")]
        public List<SuppressAction> SuppressActions { get; set; } = new();
    }

    public class UpdateIpartsRequest
    {
        [JsonPropertyName("assemblyUpdates")]
        public List<AssemblyUpdate> AssemblyUpdates { get; set; } = new();
    }
    public class UpdateAllPropertiesRequest
    {
        [JsonPropertyName("drawingspath")]
        public string DirectoryPath { get; set; } = "";

        [JsonPropertyName("ipropertiesdetails")]
        public Dictionary<string, string> IProperties { get; set; } = new();
    }
    public class UpdatePropertiesRequest { public List<Dictionary<string, object>> AssemblyUpdates { get; set; } = new(); }

    public class ModelStateUpdateRequest
    {
        [JsonPropertyName("assemblyUpdates")]
        public List<ModelStateUpdate> AssemblyUpdates { get; set; } = new();
    }
}
