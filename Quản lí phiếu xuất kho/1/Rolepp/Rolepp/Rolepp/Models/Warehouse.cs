using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Rolepp.Models
{
    public class Warehouse
    {
       
        [Key]
        public int WarehouseId { get; set; }
        public string WarehouseName { get; set; }
        public string Address { get; set; }
    }
}
