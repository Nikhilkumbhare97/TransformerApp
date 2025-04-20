using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace InventorApp.API.Models
{
    [Table("transformer_configuration")]
    public class TransformerConfiguration
    {
        [Key]
        [Column("project_unique_id")]
        public long ProjectUniqueId { get; set; }

        [Column("tank_details", TypeName = "json")]
        public string? TankDetails { get; set; }

        [Column("lv_turret_details", TypeName = "json")]
        public string? LvTurretDetails { get; set; }

        [Column("top_cover_details", TypeName = "json")]
        public string? TopCoverDetails { get; set; }

        [Column("hv_turret_details", TypeName = "json")]
        public string? HvTurretDetails { get; set; }

        [Column("piping", TypeName = "json")]
        public string? Piping { get; set; }

        [Column("lv_trunking_details", TypeName = "json")]
        public string? LvTrunkingDetails { get; set; }

        [Column("lv_hv_turret_details", TypeName = "json")]
        public string? LvHvTurretDetails { get; set; }

        [Column("conservator_details", TypeName = "json")]
        public string? ConservatorDetails { get; set; }

        [Column("conservator_support_details", TypeName = "json")]
        public string? ConservatorSupportDetails { get; set; }

    }
}