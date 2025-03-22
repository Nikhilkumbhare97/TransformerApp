using System.Text.Json.Serialization;

namespace InventorApp.API.Models
{
    public class SuppressAction
    {
        [JsonPropertyName("assemblyFilePath")]
        public string AssemblyFilePath { get; set; } = string.Empty;

        [JsonPropertyName("components")]
        public List<string> Components { get; set; } = new();

        [JsonPropertyName("suppress")]
        public bool Suppress { get; set; }
    }
}