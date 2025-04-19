using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace InventorApp.API.Models;
public class AssemblyUpdate
{
    [JsonPropertyName("assemblyFilePath")]
    public string AssemblyFilePath { get; set; } = "";

    [JsonPropertyName("iparts-iassemblies")]
    public Dictionary<string, string> IpartsIassemblies { get; set; } = new();
}