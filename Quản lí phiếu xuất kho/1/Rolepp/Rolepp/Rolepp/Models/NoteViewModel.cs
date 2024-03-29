﻿using System.ComponentModel.DataAnnotations;

namespace Rolepp.Models
{
    public class NoteViewModel
    {
        [Required]
        public string NoteCode { get; set; }
        [Required]
        public string CreateName { get; set; }
        [Required]
        public string Customer { get; set; }
        [Required]
        public string AddressCustomer { get; set; }
        [Required]
        public string Reason { get; set; }


        public List<NoteProductViewModel> Products { get; set; }

    }

    public class NoteProductViewModel
    {
        public int ProductID { get; set; }
        public int StockOut { get; set; }
    }
}
