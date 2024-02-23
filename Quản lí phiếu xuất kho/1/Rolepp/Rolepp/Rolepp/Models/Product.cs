using Microsoft.EntityFrameworkCore;
using Rolepp.Data.Migrations;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Rolepp.Models
{
    public class Product
    {
        [Key]
        public int ProductID { get; set; }

        [Required]
        [StringLength(50)]
        public string ProductName { get; set; }

        [Required]
        [Column(TypeName = "decimal(18,2)")]
        public decimal Price { get; set; }

        [Required]
        public int Quantity { get; set; }

        [Required]
        [StringLength(20)]
        public string ProductCode { get; set; }
        

        // Change to WarehouseId to match the foreign key convention
        public int? WarehouseId { get; set; }

        [ForeignKey("WarehouseId")]
        public Warehouse Warehouse { get; set; }
        public ICollection<NoteProduct> NoteProducts { get; set; }


    }

}
