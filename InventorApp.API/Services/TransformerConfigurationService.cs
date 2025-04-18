using System;
using System.Threading.Tasks;
using InventorApp.API.Models;
using InventorApp.API.Repositories;

namespace InventorApp.API.Services
{
    public class TransformerConfigurationService
    {
        private readonly ITransformerConfigurationRepository _repository;

        public TransformerConfigurationService(ITransformerConfigurationRepository repository)
        {
            _repository = repository;
        }

        public async Task<TransformerConfiguration> SaveTransformerConfigDetails(TransformerConfiguration configuration)
        {
            return await _repository.CreateAsync(configuration);
        }

        public async Task<TransformerConfiguration> UpdateTransformerConfigDetails(long projectUniqueId, TransformerConfiguration configuration)
        {
            // First get the existing configuration
            var existingConfig = await _repository.GetByIdAsync(projectUniqueId);
            if (existingConfig == null)
            {
                throw new Exception($"Transformer configuration with ID {projectUniqueId} not found.");
            }

            // Update all fields
            existingConfig.TankDetails = configuration.TankDetails;
            existingConfig.LvTurretDetails = configuration.LvTurretDetails;
            existingConfig.TopCoverDetails = configuration.TopCoverDetails;
            existingConfig.HvTurretDetails = configuration.HvTurretDetails;
            existingConfig.Piping = configuration.Piping;
            existingConfig.LvTrunkingDetails = configuration.LvTrunkingDetails;

            return await _repository.UpdateAsync(existingConfig);
        }

        public async Task<TransformerConfiguration?> GetTransformerConfigDetailsById(long projectUniqueId)
        {
            return await _repository.GetByIdAsync(projectUniqueId);
        }
    }
}