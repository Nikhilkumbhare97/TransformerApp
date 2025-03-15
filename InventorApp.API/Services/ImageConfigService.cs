using System.Collections.Generic;
using System.Threading.Tasks;
using InventorApp.API.Models;
using InventorApp.API.Repositories;

namespace InventorApp.API.Services
{
    public class ImageConfigService
    {
        private readonly IImageConfigRepository _repository;

        public ImageConfigService(IImageConfigRepository repository)
        {
            _repository = repository;
        }

        public async Task<ImageConfig?> GetImageConfig(string imageName)
        {
            return await _repository.GetByImageNameAsync(imageName);
        }

        public async Task<List<string>> GetAllImageNames()
        {
            return await _repository.GetAllImageNamesAsync();
        }

        public async Task<ImageConfig> UpdateImageConfig(ImageConfig imageConfig)
        {
            return await _repository.SaveAsync(imageConfig);
        }
    }
} 