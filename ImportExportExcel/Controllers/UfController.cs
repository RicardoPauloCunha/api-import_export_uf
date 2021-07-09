using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using ImportExportExcel.Domains;
using ImportExportExcel.Interfaces;
using ImportExportExcel.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Hosting;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.AspNetCore.Authorization;

namespace ImportExportExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UfController : ControllerBase
    {
        private readonly IUfRepository _ufRepository;
        private readonly IHostingEnvironment _hostingEnvironment;

        public UfController(IUfRepository ufRepository, IHostingEnvironment hostingEnvironment)
        {
            _ufRepository = ufRepository;
            _hostingEnvironment = hostingEnvironment;
        }

        [HttpPost]
        public IActionResult ImportarUfs([FromForm]ImportarUfViewModel input)
        {
            try
            {
                List<UfDomain> ufDomains = _ufRepository.ImportarExcel(input.Arquivo);

                return Ok(ufDomains);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        [AllowAnonymous]
        public async Task<IActionResult> ExportUf()
        {
            try
            {
                string webRootPath = _hostingEnvironment.WebRootPath;
                string fileName = "export_ufs.xlsx";

                List<UfDomain> ufs = new List<UfDomain>()
                {
                    new UfDomain(1, "Uf1", "Sigla1"),
                    new UfDomain(2, "Uf2", "Sigla2"),
                    new UfDomain(3, "Uf3", "Sigla3"),
                    new UfDomain(4, "Uf4", "Sigla4"),
                    new UfDomain(5, "Uf5", "Sigla5"),
                };

                DataTable dataTable = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(ufs), (typeof(DataTable)));

                var memoryStream = new MemoryStream();

                using (var fs = new FileStream(Path.Combine(webRootPath, fileName), FileMode.Create))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("Ufs");

                    List<string> columns = new List<string>();
                    IRow row = sheet.CreateRow(0);
                    int columnIndex = 0;
                    int rowIndex = 1;

                    foreach (DataColumn colum in dataTable.Columns)
                    {
                        columns.Add(colum.ColumnName);
                        row.CreateCell(columnIndex).SetCellValue(colum.ColumnName);
                        columnIndex++;
                    }

                    foreach (DataRow drow in dataTable.Rows)
                    {
                        int cellIndex = 0;

                        row = sheet.CreateRow(rowIndex);

                        foreach (string column in columns)
                        {
                            row.CreateCell(cellIndex).SetCellValue(drow[column].ToString());
                            cellIndex++;
                        }

                        rowIndex++;
                    }

                    workbook.Write(fs);

                }

                using (var fileStream = new FileStream(Path.Combine(webRootPath, fileName), FileMode.Open))
                {
                    await fileStream.CopyToAsync(memoryStream);
                }

                memoryStream.Position = 0;

                return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}