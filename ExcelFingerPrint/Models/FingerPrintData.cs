namespace ExcelFingerPrint.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FingerPrintData")]
    public partial class FingerPrintData
    {
        [StringLength(255)]
        public string Id { get; set; }

        [StringLength(50)]
        public string GuestID { get; set; }

        [StringLength(50)]
        public string CardNo { get; set; }

        [StringLength(255)]
        public string GuestName { get; set; }

        [StringLength(500)]
        public string Department { get; set; }

        [StringLength(50)]
        public string Date { get; set; }

        public DateTime? Time { get; set; }

        [StringLength(50)]
        public string EntryDoor { get; set; }

        [StringLength(50)]
        public string EventDescription { get; set; }

        [StringLength(50)]
        public string VerificationSource { get; set; }
    }
}
