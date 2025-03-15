using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace InventorApp.API.Models
{
    [Table("image_config")]
    public class ImageConfig
    {
        [Key]
        [Column("image_name")]
        public string ImageName { get; set; }

        [Column("config", TypeName = "json")]
        public string Config { get; set; }

        public ImageConfig()
        {
            ImageName = string.Empty;
            Config = string.Empty;
        }

        public ImageConfig(string imageName, string config)
        {
            ImageName = imageName;
            Config = config;
        }
    }
} 