using ClosedXML.Excel;
using ExportsJuntos.Models;
using ExportsJuntos.Repositories;

namespace ExportsJuntos.Infra.Files.Writer
{
    public partial class PortfolioFileRepository : IPortfolioFileRepository
    {
        public Stream Write(IEnumerable<Portfolio> portfolios)
        {
            var stream = new MemoryStream();
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet 1");
            FillWorksheet(worksheet, portfolios);
            workbook.SaveAs(stream);
            stream.Position = 0;
            return stream;
        }

        private static void FillWorksheet(IXLWorksheet worksheet, IEnumerable<Portfolio> portfolios)
        {
            var headers = GetPortfolioHeaders();
            var column = 1;
            foreach (var header in headers)
            {
                worksheet.Cell(row: 1, column++)
                    .SetValue(header)
                    .Style
                    .Font.SetBold(true)
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }

            var row = 2;
            foreach (var portfolio in portfolios)
            {
                column = 1;
                SetCell(portfolio.PortfolioName);
                SetCell(portfolio.Description);
                SetCell(portfolio.CurrencyId);
                SetCell(portfolio.AssetName);
                SetCell(portfolio.Date);
                SetCell(portfolio.PortfolioTypeId);
                SetCell(portfolio.TenantId);
                SetCell(portfolio.ManagerId);
                SetCell(portfolio.CountryId);
                SetCell(portfolio.Name);
                SetCell(portfolio.Nickname);
                SetCell(portfolio.BirthDate);
                SetCell(portfolio.Age);
                SetCell(portfolio.Document);
                row++;

                void SetCell(object value)
                {
                    worksheet
                        .Cell(row, column++)
                        .SetValue(value)
                        .Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                }
            }
        }

        private static string[] GetPortfolioHeaders() => new[]
        {
            "Nome Carteira",
            "Descrição",
            "Id da Moeda",
            "Nome do Ativo",
            "Data",
            "Tipo de Carteira",
            "TenantId",
            "ManagerId",
            "CountryId",
            "Name",
            "Nickname",
            "BirthDate",
            "Age",
            "Document"
        };
    }
}