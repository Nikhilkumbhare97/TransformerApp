using System;
using System.Threading.Tasks;
using InventorApp.API.Models;
using InventorApp.API.Repositories;

namespace InventorApp.API.Services
{
    public class TransformerService
    {
        private readonly ITransformerRepository _transformerRepository;

        public TransformerService(ITransformerRepository transformerRepository)
        {
            _transformerRepository = transformerRepository;
        }

        public async Task<Transformer> SaveTransformerDetails(Transformer transformer)
        {
            return await _transformerRepository.CreateAsync(transformer);
        }

        public async Task<Transformer> UpdateTransformerDetails(long projectUniqueId, Transformer transformer)
        {
            var existingTransformer = await _transformerRepository.GetByIdAsync(projectUniqueId);
            if (existingTransformer == null)
            {
                throw new Exception($"Transformer details not found with projectUniqueId: {projectUniqueId}");
            }

            if (!string.IsNullOrEmpty(transformer.TransformerType))
                existingTransformer.TransformerType = transformer.TransformerType;
            if (!string.IsNullOrEmpty(transformer.DesignType))
                existingTransformer.DesignType = transformer.DesignType;

            return await _transformerRepository.UpdateAsync(existingTransformer);
        }

        public async Task<Transformer?> GetTransformerDetailsById(long projectUniqueId)
        {
            return await _transformerRepository.GetByIdAsync(projectUniqueId);
        }
    }
} 