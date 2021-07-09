using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace ImportExportExcel.ViewModels
{
    public class ImportarUfViewModel
    {
        [Required]
        public IFormFile Arquivo { get; set; }

        public ImportarUfViewModel()
        {

        }
    }
}
