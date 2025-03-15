using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using InventorApp.API.Models;
using InventorApp.API.Services;
using Microsoft.AspNetCore.Cors;

namespace InventorApp.API.Controllers
{
    [Route("api/config")]
    [ApiController]
    [EnableCors("AllowSpecificOrigin")]
    public class ImageConfigController : ControllerBase
    {
        private readonly ImageConfigService _service;

        public ImageConfigController(ImageConfigService service)
        {
            _service = service;
        }

        [HttpGet("{imageName}")]
        public async Task<ActionResult<ImageConfig>> GetImageConfig(string imageName)
        {
            var imageConfig = await _service.GetImageConfig(imageName);
            if (imageConfig == null)
            {
                return NotFound();
            }
            return Ok(imageConfig);
        }

        [HttpGet("images")]
        public async Task<ActionResult<List<string>>> GetAllImageNames()
        {
            var imageNames = await _service.GetAllImageNames();
            return Ok(imageNames);
        }

        [HttpPut("{imageName}")]
        public async Task<ActionResult<ImageConfig>> UpdateImageConfig(string imageName, [FromBody] ImageConfig imageConfig)
        {
            if (imageName != imageConfig.ImageName)
            {
                return BadRequest();
            }

            var updatedConfig = await _service.UpdateImageConfig(imageConfig);
            return Ok(updatedConfig);
        }
    }
} 