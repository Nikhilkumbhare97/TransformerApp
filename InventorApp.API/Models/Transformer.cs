using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace InventorApp.API.Models
{
    [Table("transformer_details", Schema = "transformer")]
    public class Transformer
    {
        [Key]
        [Column("project_unique_id")]
        public long ProjectUniqueId { get; set; }

        [Required]
        [Column("transformer_type")]
        [StringLength(100)]
        public string TransformerType { get; set; }

        [Required]
        [Column("design_type")]
        [StringLength(50)]
        public string DesignType { get; set; }

        public Transformer()
        {
            TransformerType = string.Empty;
            DesignType = string.Empty;
        }
    }
} 