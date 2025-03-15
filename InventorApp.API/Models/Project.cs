using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace InventorApp.API.Models
{
    [Table("projects")]
    public class Project
    {
        [Key]
        [Column("project_unique_id")]
        public long ProjectUniqueId { get; set; }

        [Required]
        [Column("project_id")]
        [StringLength(200)]
        public string ProjectId { get; set; }

        [Required]
        [Column("project_name")]
        [StringLength(200)]
        public string ProjectName { get; set; }

        [Required]
        [Column("project_number")]
        [StringLength(100)]
        public string ProjectNumber { get; set; }

        [Required]
        [Column("client_name")]
        [StringLength(200)]
        public string ClientName { get; set; }

        [Required]
        [Column("created_by")]
        [StringLength(200)]
        public string CreatedBy { get; set; }

        [Required]
        [Column("prepared_by")]
        [StringLength(200)]
        public string PreparedBy { get; set; }

        [Required]
        [Column("checked_by")]
        [StringLength(200)]
        public string CheckedBy { get; set; }

        [Required]
        [Column("date")]
        public DateTime Date { get; set; }

        [Required]
        [Column("status")]
        [StringLength(20)]
        public string Status { get; set; }

        public Project()
        {
            ProjectId = string.Empty;
            ProjectName = string.Empty;
            ProjectNumber = string.Empty;
            ClientName = string.Empty;
            CreatedBy = string.Empty;
            PreparedBy = string.Empty;
            CheckedBy = string.Empty;
            Status = "Active";
            Date = DateTime.Now;
        }
    }
}