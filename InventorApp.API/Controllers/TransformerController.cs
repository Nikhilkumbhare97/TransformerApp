using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using InventorApp.API.Models;
using InventorApp.API.Services;
using Microsoft.AspNetCore.Cors;

namespace InventorApp.API.Controllers
{
    [Route("api/transformer")]
    [ApiController]
    [EnableCors("AllowSpecificOrigin")]
    public class TransformerController : ControllerBase
    {
        private readonly TransformerService _transformerService;

        public TransformerController(TransformerService transformerService)
        {
            _transformerService = transformerService;
        }

        [HttpPost]
        public async Task<ActionResult<Transformer>> SaveTransformerDetails([FromBody] Transformer transformer)
        {
            var savedTransformer = await _transformerService.SaveTransformerDetails(transformer);
            return CreatedAtAction(nameof(GetTransformerDetailsById), new { projectUniqueId = savedTransformer.ProjectUniqueId }, savedTransformer);
        }

        [HttpPut("{projectUniqueId}")]
        public async Task<ActionResult<Transformer>> UpdateTransformerDetails(long projectUniqueId, [FromBody] Transformer transformer)
        {
            var updatedTransformer = await _transformerService.UpdateTransformerDetails(projectUniqueId, transformer);
            return Ok(updatedTransformer);
        }

        [HttpGet("{projectUniqueId}")]
        public async Task<ActionResult<Transformer>> GetTransformerDetailsById(long projectUniqueId)
        {
            var transformer = await _transformerService.GetTransformerDetailsById(projectUniqueId);
            if (transformer == null)
            {
                return NotFound();
            }
            return Ok(transformer);
        }
    }
} 