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
            configuration.ProjectUniqueId = projectUniqueId;
            return await _repository.UpdateAsync(configuration);
        }

        public async Task<TransformerConfiguration?> GetTransformerConfigDetailsById(long projectUniqueId)
        {
            return await _repository.GetByIdAsync(projectUniqueId);
        }
    }
} 