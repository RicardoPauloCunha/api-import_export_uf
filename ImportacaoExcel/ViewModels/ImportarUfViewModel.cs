using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ImportacaoExcel.ViewModels
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
