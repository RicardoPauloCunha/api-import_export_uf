using ImportExportExcel.Domains;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ImportExportExcel.Interfaces
{
    public interface IUfRepository
    {
        List<UfDomain> ImportarExcel(IFormFile formFile);
        Task ExportExcel();
    }
}
