using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Services;
using System.Text.Json.Serialization;
using InventorApp.API.Models;
using System.ComponentModel.DataAnnotations;

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


        [HttpPost("design-assist-recursive-rename-with-prefix")]
        public IActionResult DesignAssistRecursiveRenameWithPrefix([FromBody] RecursiveRenameWithPrefixRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.ModelPath))
                return BadRequest(new { message = "modelPath is required." });

            if (string.IsNullOrWhiteSpace(request.Prefix))
                return BadRequest(new { message = "prefix is required." });

            // Validate path format
            try
            {
                var fullPath = Path.GetFullPath(request.ModelPath);
                if (!Directory.Exists(fullPath))
                    return BadRequest(new
                    {
                        message = $"Directory not found: {fullPath}",
                        providedPath = request.ModelPath,
                        resolvedPath = fullPath
                    });
            }
            catch (Exception ex)
            {
                return BadRequest(new
                {
                    message = $"Invalid path format: {ex.Message}",
                    providedPath = request.ModelPath
                });
            }

            try
            {
                // Step 1: Perform the recursive rename
                var filesToDelete = _assemblyService.RenameAssemblyRecursivelyWithPrefix(request.ModelPath, request.Prefix);

                // Step 2: If rename was successful and there are files to delete, call the delete API
                if (filesToDelete != null && filesToDelete.Count > 0)
                {
                    var deleteResult = _assemblyService.DeleteFiles(filesToDelete);
                    return Ok(new
                    {
                        message = "Recursive rename with prefix completed and old files cleaned up.",
                        filesToDelete = filesToDelete,
                        deleteResult = deleteResult
                    });
                }
                else
                {
                    return Ok(new
                    {
                        message = "Recursive rename with prefix completed. No files to delete.",
                        filesToDelete = filesToDelete
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error during recursive rename with prefix: {ex.Message}" });
            }
        }

        [HttpPost("design-assist-recursive-rename-with-prefix-and-drawings")]
        public IActionResult DesignAssistRecursiveRenameWithPrefixAndDrawings([FromBody] RecursiveRenameWithPrefixAndDrawingsRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.ModelPath))
                return BadRequest(new { message = "modelPath is required." });

            if (string.IsNullOrWhiteSpace(request.DrawingsPath))
                return BadRequest(new { message = "drawingspath is required." });

            if (string.IsNullOrWhiteSpace(request.OldPrefix))
                return BadRequest(new { message = "oldPrefix is required." });

            if (string.IsNullOrWhiteSpace(request.NewPrefix))
                return BadRequest(new { message = "newPrefix is required." });

            // Validate path formats
            try
            {
                var modelFullPath = Path.GetFullPath(request.ModelPath);
                var drawingsFullPath = Path.GetFullPath(request.DrawingsPath);

                if (!Directory.Exists(modelFullPath))
                    return BadRequest(new
                    {
                        message = $"Model directory not found: {modelFullPath}",
                        providedPath = request.ModelPath,
                        resolvedPath = modelFullPath
                    });

                if (!Directory.Exists(drawingsFullPath))
                    return BadRequest(new
                    {
                        message = $"Drawings directory not found: {drawingsFullPath}",
                        providedPath = request.DrawingsPath,
                        resolvedPath = drawingsFullPath
                    });
            }
            catch (Exception ex)
            {
                return BadRequest(new
                {
                    message = $"Invalid path format: {ex.Message}",
                    modelPath = request.ModelPath,
                    drawingsPath = request.DrawingsPath
                });
            }

            try
            {
                // Perform the enhanced recursive rename with drawing updates
                var result = _assemblyService.RenameAssemblyRecursivelyWithPrefixAndUpdateDrawings(
                    request.ModelPath,
                    request.DrawingsPath,
                    request.ProjectPath,
                    request.OldPrefix,
                    request.NewPrefix);

                // Step 2: If rename was successful and there are files to delete, call the delete API
                if (result.FilesToDelete != null && result.FilesToDelete.Count > 0)
                {
                    var deleteResult = _assemblyService.DeleteFiles(result.FilesToDelete);
                    return Ok(new
                    {
                        message = "Recursive rename with prefix and drawing updates completed and old files cleaned up.",
                        result = result,
                        deleteResult = deleteResult
                    });
                }
                else
                {
                    return Ok(new
                    {
                        message = "Recursive rename with prefix and drawing updates completed. No files to delete.",
                        result = result
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error during recursive rename with prefix and drawing updates: {ex.Message}" });
            }
        }

        [HttpPost("delete-files")]
        public IActionResult DeleteFiles([FromBody] DeleteFilesRequest request)
        {
            if (request.FilePaths == null || request.FilePaths.Count == 0)
            {
                return BadRequest(new { message = "filePaths is required and cannot be empty." });
            }

            try
            {
                var result = _assemblyService.DeleteFiles(request.FilePaths);
                return Ok(new
                {
                    message = "File deletion completed.",
                    result = result
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error during file deletion: {ex.Message}" });
            }
        }

        public class RecursiveRenameWithPrefixRequest
        {
            public string ModelPath { get; set; } = "";
            public string Prefix { get; set; } = "";
        }

        public class RecursiveRenameWithPrefixAndDrawingsRequest
        {
            [JsonPropertyName("drawingspath")]
            public string DrawingsPath { get; set; } = "";

            [JsonPropertyName("modelPath")]
            public string ModelPath { get; set; } = "";

            [JsonPropertyName("projectpath")]
            public string ProjectPath { get; set; } = "";

            [JsonPropertyName("oldPrefix")]
            public string OldPrefix { get; set; } = "";

            [JsonPropertyName("newPrefix")]
            public string NewPrefix { get; set; } = "";
        }

        public class DeleteFilesRequest
        {
            [JsonPropertyName("filePaths")]
            [Required(ErrorMessage = "filePaths is required")]
            public List<string> FilePaths { get; set; } = new();
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