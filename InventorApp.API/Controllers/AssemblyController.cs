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

        [HttpPost("design-assist-rename")]
        public IActionResult DesignAssistRename([FromBody] DesignAssistRenameRequest request)
        {
            // Enhanced validation with more specific error messages
            if (string.IsNullOrWhiteSpace(request.DrawingsPath))
                return BadRequest(new { message = "drawingsPath is required and cannot be empty or whitespace." });

            if (string.IsNullOrWhiteSpace(request.PartPrefix))
                return BadRequest(new { message = "partPrefix is required and cannot be empty or whitespace." });

            // Validate path format
            try
            {
                var fullPath = Path.GetFullPath(request.DrawingsPath);
                if (!Directory.Exists(fullPath))
                    return BadRequest(new
                    {
                        message = $"Directory not found: {fullPath}",
                        providedPath = request.DrawingsPath,
                        resolvedPath = fullPath
                    });
            }
            catch (Exception ex)
            {
                return BadRequest(new
                {
                    message = $"Invalid path format: {ex.Message}",
                    providedPath = request.DrawingsPath
                });
            }

            try
            {
                // Log the operation start
                Console.WriteLine($"=== Design Assistant Rename Operation Started ===");
                Console.WriteLine($"Path: {request.DrawingsPath}");
                Console.WriteLine($"Prefix: {request.PartPrefix}");
                Console.WriteLine($"Assembly List Provided: {request.AssemblyList?.Count ?? 0} assemblies");

                // Pass the assemblyList (can be null/empty for auto-discovery)
                bool result = _assemblyService.DesignAssistRename(
                    request.DrawingsPath,
                    request.PartPrefix,
                    request.AssemblyList?.Count > 0 ? request.AssemblyList : null
                );

                if (result)
                {
                    var response = new
                    {
                        message = "Design Assistant renaming completed successfully.",
                        processedPath = request.DrawingsPath,
                        prefix = request.PartPrefix,
                        autoDiscovered = request.AssemblyList?.Count == 0,
                        timestamp = DateTime.UtcNow,
                        status = "success"
                    };

                    Console.WriteLine($"=== Design Assistant Rename Operation Completed Successfully ===");
                    return Ok(response);
                }
                else
                {
                    var response = new
                    {
                        message = "Design Assistant renaming failed. Check the application logs for details.",
                        processedPath = request.DrawingsPath,
                        prefix = request.PartPrefix,
                        autoDiscovered = request.AssemblyList?.Count == 0,
                        timestamp = DateTime.UtcNow,
                        status = "failed"
                    };

                    Console.WriteLine($"=== Design Assistant Rename Operation Failed ===");
                    return StatusCode(500, response);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                var response = new
                {
                    message = $"Access denied to directory: {ex.Message}",
                    processedPath = request.DrawingsPath,
                    prefix = request.PartPrefix,
                    timestamp = DateTime.UtcNow,
                    status = "access_denied"
                };
                return StatusCode(403, response);
            }
            catch (Exception ex)
            {
                var response = new
                {
                    message = $"An unexpected error occurred: {ex.Message}",
                    processedPath = request.DrawingsPath,
                    prefix = request.PartPrefix,
                    timestamp = DateTime.UtcNow,
                    status = "error",
                    errorType = ex.GetType().Name
                };
                return StatusCode(500, response);
            }
        }

        /// <summary>
        /// Analyzes what files would be renamed without performing the actual rename operation
        /// </summary>
        [HttpPost("design-assist-analyze")]
        public IActionResult DesignAssistAnalyze([FromBody] DesignAssistRenameRequest request)
        {
            // Enhanced validation with more specific error messages
            if (string.IsNullOrWhiteSpace(request.DrawingsPath))
                return BadRequest(new { message = "drawingsPath is required and cannot be empty or whitespace." });

            if (string.IsNullOrWhiteSpace(request.PartPrefix))
                return BadRequest(new { message = "partPrefix is required and cannot be empty or whitespace." });

            // Validate path format
            try
            {
                var fullPath = Path.GetFullPath(request.DrawingsPath);
                if (!Directory.Exists(fullPath))
                    return BadRequest(new
                    {
                        message = $"Directory not found: {fullPath}",
                        providedPath = request.DrawingsPath,
                        resolvedPath = fullPath
                    });
            }
            catch (Exception ex)
            {
                return BadRequest(new
                {
                    message = $"Invalid path format: {ex.Message}",
                    providedPath = request.DrawingsPath
                });
            }

            try
            {
                // Analyze what would be renamed without actually doing it
                var analysis = _assemblyService.AnalyzeDesignAssistRename(
                    request.DrawingsPath,
                    request.PartPrefix,
                    request.AssemblyList?.Count > 0 ? request.AssemblyList : null
                );

                var response = new
                {
                    message = "Analysis completed successfully.",
                    processedPath = request.DrawingsPath,
                    prefix = request.PartPrefix,
                    autoDiscovered = request.AssemblyList?.Count == 0,
                    timestamp = DateTime.UtcNow,
                    status = "analysis_complete",
                    analysis = analysis
                };

                return Ok(response);
            }
            catch (UnauthorizedAccessException ex)
            {
                var response = new
                {
                    message = $"Access denied to directory: {ex.Message}",
                    processedPath = request.DrawingsPath,
                    prefix = request.PartPrefix,
                    timestamp = DateTime.UtcNow,
                    status = "access_denied"
                };
                return StatusCode(403, response);
            }
            catch (Exception ex)
            {
                var response = new
                {
                    message = $"An unexpected error occurred: {ex.Message}",
                    processedPath = request.DrawingsPath,
                    prefix = request.PartPrefix,
                    timestamp = DateTime.UtcNow,
                    status = "error",
                    errorType = ex.GetType().Name
                };
                return StatusCode(500, response);
            }
        }

        [HttpPost("design-assist-recursive-rename")]
        public IActionResult DesignAssistRecursiveRename([FromBody] RecursiveRenameRequest request)
        {
            if (request == null || request.AssemblyDocumentNames == null || request.AssemblyDocumentNames.Count == 0 || request.FileNames == null || request.FileNames.Count == 0)
            {
                return BadRequest(new { message = "assemblyDocumentNames and fileNames are required." });
            }
            try
            {
                var result = _assemblyService.RenameAssemblyRecursively(request.AssemblyDocumentNames, request.FileNames);
                return Ok(new { message = "Recursive rename completed.", filesToDelete = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error during recursive rename: {ex.Message}" });
            }
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
                    return Ok(new { 
                        message = "Recursive rename with prefix completed and old files cleaned up.", 
                        filesToDelete = filesToDelete,
                        deleteResult = deleteResult
                    });
                }
                else
                {
                    return Ok(new { 
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
                return Ok(new { 
                    message = "File deletion completed.", 
                    result = result 
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error during file deletion: {ex.Message}" });
            }
        }

        [HttpPost("update-drawing-references")]
        public IActionResult UpdateDrawingReferences([FromBody] UpdateDrawingReferencesRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.DrawingsPath))
                return BadRequest(new { message = "drawingsPath is required." });

            if (string.IsNullOrWhiteSpace(request.ModelPath))
                return BadRequest(new { message = "modelPath is required." });

            if (string.IsNullOrWhiteSpace(request.OldPrefix))
                return BadRequest(new { message = "oldPrefix is required." });

            if (string.IsNullOrWhiteSpace(request.NewPrefix))
                return BadRequest(new { message = "newPrefix is required." });

            // Validate paths
            try
            {
                var drawingsFullPath = Path.GetFullPath(request.DrawingsPath);
                var modelFullPath = Path.GetFullPath(request.ModelPath);
                
                if (!Directory.Exists(drawingsFullPath))
                    return BadRequest(new
                    {
                        message = $"Drawings directory not found: {drawingsFullPath}",
                        providedPath = request.DrawingsPath,
                        resolvedPath = drawingsFullPath
                    });

                if (!Directory.Exists(modelFullPath))
                    return BadRequest(new
                    {
                        message = $"Model directory not found: {modelFullPath}",
                        providedPath = request.ModelPath,
                        resolvedPath = modelFullPath
                    });
            }
            catch (Exception ex)
            {
                return BadRequest(new
                {
                    message = $"Invalid path format: {ex.Message}",
                    drawingsPath = request.DrawingsPath,
                    modelPath = request.ModelPath
                });
            }

            try
            {
                var result = _assemblyService.UpdateDrawingReferences(request.DrawingsPath, request.ModelPath, request.ProjectPath, request.OldPrefix, request.NewPrefix);
                return Ok(new { 
                    message = "Drawing and project references updated and files renamed successfully.", 
                    processedDrawings = result.ProcessedDrawings,
                    updatedReferences = result.UpdatedReferences,
                    failedDrawings = result.FailedDrawings,
                    renamedDrawings = result.RenamedDrawings,
                    failedRenames = result.FailedRenames,
                    renamedProjects = result.RenamedProjects,
                    failedProjectRenames = result.FailedProjectRenames
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = $"Error updating drawing references: {ex.Message}" });
            }
        }

        public class RecursiveRenameRequest
        {
            public List<string> AssemblyDocumentNames { get; set; } = new();
            public Dictionary<string, string> FileNames { get; set; } = new();
        }

        public class RecursiveRenameWithPrefixRequest
        {
            public string ModelPath { get; set; } = "";
            public string Prefix { get; set; } = "";
        }

        public class UpdateDrawingReferencesRequest
        {
            public string DrawingsPath { get; set; } = "";
            public string ModelPath { get; set; } = "";
            public string ProjectPath { get; set; } = "";
            public string OldPrefix { get; set; } = "";
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

    public class DesignAssistRenameRequest
    {
        /// <summary>
        /// The directory path containing the CAD files to be renamed
        /// </summary>
        [JsonPropertyName("drawingspath")]
        [Required(ErrorMessage = "drawingsPath is required")]
        [StringLength(500, MinimumLength = 1, ErrorMessage = "drawingsPath must be between 1 and 500 characters")]
        public string DrawingsPath { get; set; } = "";

        /// <summary>
        /// Optional list of specific assembly files to process. If null or empty, auto-discovery will be used.
        /// </summary>
        [JsonPropertyName("assemblyList")]
        public List<string>? AssemblyList { get; set; } = null;

        /// <summary>
        /// The new prefix to apply to matching components
        /// </summary>
        [JsonPropertyName("partPrefix")]
        [Required(ErrorMessage = "partPrefix is required")]
        [StringLength(50, MinimumLength = 1, ErrorMessage = "partPrefix must be between 1 and 50 characters")]
        [RegularExpression(@"^[A-Za-z0-9_-]+$", ErrorMessage = "partPrefix can only contain letters, numbers, underscores, and hyphens")]
        public string PartPrefix { get; set; } = "";
    }
}
