using System.Text.Json.Serialization;

namespace InventorApp.API.Models
{
    public class ModelStateUpdate
    {
        [JsonPropertyName("assemblyFilePath")]
        public string AssemblyFilePath { get; set; } = "";

        [JsonPropertyName("model-state")]
        public string ModelState { get; set; } = "";

        [JsonPropertyName("representations")]
        public string Representations { get; set; } = "";
    }
}