using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using InventorApp.API.Models;
using InventorApp.API.Services;
using Microsoft.AspNetCore.Cors;

namespace InventorApp.API.Controllers
{
    [Route("api/transformer-config")]
    [ApiController]
    [EnableCors("AllowSpecificOrigin")]
    public class TransformerConfigurationController : ControllerBase
    {
        private readonly TransformerConfigurationService _service;

        public TransformerConfigurationController(TransformerConfigurationService service)
        {
            _service = service;
        }

        [HttpPost]
        public async Task<ActionResult<TransformerConfiguration>> SaveTransformerConfigDetails([FromBody] TransformerConfiguration configuration)
        {
            var savedConfig = await _service.SaveTransformerConfigDetails(configuration);
            return CreatedAtAction(nameof(GetTransformerDetailsById), new { projectUniqueId = savedConfig.ProjectUniqueId }, savedConfig);
        }

        [HttpPut("{projectUniqueId}")]
        public async Task<ActionResult<TransformerConfiguration>> UpdateTransformerDetails(long projectUniqueId, [FromBody] TransformerConfiguration configuration)
        {
            try
            {
                var updatedConfig = await _service.UpdateTransformerConfigDetails(projectUniqueId, configuration);
                return Ok(updatedConfig);
            }
            catch (Exception ex)
            {
                return NotFound(ex.Message);
            }
        }

        [HttpGet("{projectUniqueId}")]
        public async Task<ActionResult<TransformerConfiguration>> GetTransformerDetailsById(long projectUniqueId)
        {
            var config = await _service.GetTransformerConfigDetailsById(projectUniqueId);
            if (config == null)
            {
                return NotFound();
            }
            return Ok(config);
        }
    }
}