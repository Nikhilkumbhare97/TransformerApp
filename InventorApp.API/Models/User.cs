using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace InventorApp.API.Models
{
    [Table("users", Schema = "transformer")]
    public class User
    {
        [Key]
        [Column("user_unique_id")]
        public long UserUniqueId { get; set; }

        [Required]
        [Column("name")]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;

        [Required]
        [Column("employee_id")]
        [StringLength(200)]
        public string EmployeeId { get; set; } = string.Empty;

        [Required]
        [EmailAddress]
        [Column("email")]
        [StringLength(320)]
        public string Email { get; set; } = string.Empty;

        // Store encrypted password text (base64) – never plain
        [Required]
        [Column("password_encrypted")]
        [JsonPropertyName("password")] // bind/serialize as "password"
        public string PasswordEncrypted { get; set; } = string.Empty;

        [Required]
        [Column("role")]
        [StringLength(50)]
        public string Role { get; set; } = "User";

        [Required]
        [Column("status")]
        [StringLength(20)]
        public string Status { get; set; } = "Active";
    }
}

