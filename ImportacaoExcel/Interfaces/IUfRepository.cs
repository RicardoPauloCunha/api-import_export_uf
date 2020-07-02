using ImportacaoExcel.Domains;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ImportacaoExcel.Interfaces
{
    public interface IUfRepository
    {
        List<UfDomain> ImportarExcel(IFormFile formFile);
        Task ExportExcel();
    }
}
