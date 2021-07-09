using ImportExportExcel.Domains;
using ImportExportExcel.Interfaces;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ImportExportExcel.Repositories
{
    public class UfRepository : IUfRepository
    {
        public Task ExportExcel()
        {
            return Task.CompletedTask;
        }

        public List<UfDomain> ImportarExcel(IFormFile formFile)
        {
            // Gera um objeto com todas as informações do excel importado
            IWorkbook excel = WorkbookFactory.Create(formFile.OpenReadStream());

            // Pega a planilha com a tabela de registros das fus
            ISheet tabela = excel.GetSheet("AR_BR_UF_2018");

            // Caso não encontre a planilha
            if (tabela == null)
            {
                // Lança uma exception
                throw new Exception("A planilha não foi encontrada");
            }

            // Pega o index da primeira linha (Titulos) dessa tabela
            int primeiraLinhaIndexTitulos = tabela.FirstRowNum;

            // Pega o index da seguna linha (1° Registro) dessa tabela
            int segundaLinhaIndex = (tabela.FirstRowNum + 1);

            // Pega o index da ultima linha (Ultimo Registro) dessa tabela
            int ultimaLinhaIndex = tabela.LastRowNum;

            // Declara a variavel que armazena a celula com o titulo da coluna de id
            ICell idTituloCelula = null;

            // Declara a variavel que armazena a celula com o titulo da coluna do nome
            ICell nomeTituloCelula = null;

            // Declara a variavel que armazena a celula com o titulo da coluna da sigla
            ICell siglaTituloCelula = null;

            // Declara uma lista de ufs
            List<UfDomain> ufsImportadas = new List<UfDomain>();

            // Pega a linha dos titulos
            IRow linhasTitulos = tabela.GetRow(primeiraLinhaIndexTitulos);

            // Passa por todas as celulas (colunas) dessa linha de titulos
            foreach (ICell celula in linhasTitulos.Cells)
            {
                // Recupera os titulos atraves do que está escrito dentro da celula
                switch(celula.StringCellValue)
                {
                    // Caso seja a celula do ID
                    case "ID":
                        idTituloCelula = celula;
                        break;

                    // Caso seja a celula do Nome
                    case "NM_UF":
                        nomeTituloCelula = celula;
                        break;

                    // Caso seja a celula do Sigla
                    case "NM_UF_SIGLA":
                        siglaTituloCelula = celula;
                        break;

                    // Nenhum Caso
                    default:
                        break;
                }
            }

            // Passa por todas as linhas da tabela
            for (int rowNum = segundaLinhaIndex; rowNum < ultimaLinhaIndex; rowNum++)
            {
                // Recupera a linha atual
                IRow row = tabela.GetRow(rowNum);

                // Declara a variavel que armazena o id da uf
                int idUfCelula = 0;

                // Declara a variavel que armazena o nome da uf
                string nomeufCelula = "";

                // Declara a variavel que armazena a sigla da uf
                string siglaUfCelula = "";

                // Passa por todas as celulas da linha(Colunas)
                foreach (ICell cell in row.Cells)
                {
                    // Caso a coluna do celula seja a mesma do id
                    if (cell.ColumnIndex == idTituloCelula.ColumnIndex)
                    {
                        // Atribui ao id o valor da celula
                        idUfCelula = Convert.ToInt32(cell.NumericCellValue);
                    }

                    // Caso a coluna do celula seja a mesma do nome
                    if (cell.ColumnIndex == nomeTituloCelula.ColumnIndex)
                    {
                        // Atribui o nome o valor da celula
                        nomeufCelula = cell.StringCellValue;
                    }

                    // Caso a coluna do celula seja a mesma da sigla
                    if (cell.ColumnIndex == siglaTituloCelula.ColumnIndex)
                    {
                        // Atribui a sigla o valor da celula
                        siglaUfCelula = cell.StringCellValue;
                    }
                }

                // Intancia a ufDomain
                UfDomain ufDomain = new UfDomain(idUfCelula, nomeufCelula, siglaUfCelula);

                // Adiciona a uf na lista
                ufsImportadas.Add(ufDomain);
            }

            // Retorna a lista de ufs
            return ufsImportadas;
        }
    }
}
