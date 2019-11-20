using System.ComponentModel.DataAnnotations;

namespace EpPlusPoC
{
    public class ExcelResourceDto
    {
        [Column(2)]
        [Required]
        public string Title { get; set; }

        [Column(3)]
        [Required]
        public string SearchTags { get; set; }
    }
}
