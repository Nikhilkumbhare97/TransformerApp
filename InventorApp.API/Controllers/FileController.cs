using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace InventorApp.API.Controllers
{
    [ApiController]
    [Route("api/files")]
    public class FileController : ControllerBase
    {
        private readonly string _sourcePath;
        private readonly string _destinationPath;

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
        public FileController(IConfiguration configuration)
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
        {
#pragma warning disable CS8601 // Possible null reference assignment.
            _sourcePath = configuration["FilePaths:SourcePath"];
#pragma warning restore CS8601 // Possible null reference assignment.
#pragma warning disable CS8601 // Possible null reference assignment.
            _destinationPath = configuration["FilePaths:DestinationPath"];
#pragma warning restore CS8601 // Possible null reference assignment.
        }

        /// <summary>
        /// API to Copy Folder to Destination
        /// </summary>
        [HttpPost("copy-folder/{folderName}")]
        public IActionResult CopyFolder(string folderName)
        {
            string sourceDir = Path.Combine(_sourcePath, folderName);
            string destDir = Path.Combine(_destinationPath, folderName);

            if (!Directory.Exists(sourceDir))
            {
                return BadRequest("Source folder does not exist.");
            }

            try
            {
                // Create all directories
                foreach (string dirPath in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
                {
                    Directory.CreateDirectory(dirPath.Replace(sourceDir, destDir));
                }

                // Copy all files
                foreach (string filePath in Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories))
                {
                    string destFilePath = filePath.Replace(sourceDir, destDir);
                    System.IO.File.Copy(filePath, destFilePath, true);
                }

                return Ok($"Folder copied successfully to {destDir}");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error copying folder: {ex.Message}");
            }
        }

        /// <summary>
        /// API to Rename Folder in Destination
        /// </summary>
        [HttpPost("rename-folder/{oldFolderName}/{newFolderName}")]
        public IActionResult RenameFolder(string oldFolderName, string newFolderName)
        {
            string oldPath = Path.Combine(_destinationPath, oldFolderName);
            string newPath = Path.Combine(_destinationPath, newFolderName);

            if (!Directory.Exists(oldPath))
            {
                return BadRequest("Folder to rename does not exist.");
            }

            try
            {
                Directory.Move(oldPath, newPath);
                return Ok($"Folder renamed successfully from {oldFolderName} to {newFolderName}");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error renaming folder: {ex.Message}");
            }
        }
    }
}