using ExportsJuntos.Models;
using ExportsJuntos.Repositories;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportsJuntos.Infra.Files.Writer;
public class EPPlus : IEPPlusLibRepository
{
    public void CriarPlanilha(IEnumerable<Portfolio> portfolios, string caminhoPlanilha)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelPackage excel = new ExcelPackage();

        var worksheet = excel.Workbook.Worksheets.Add("Teste");

        var headers = GetPortfolioHeaders();
        var column = 1;
        var row = 1;
        foreach (var header in headers)
        {
            worksheet.Cells[row, column++].Value = header;
        }

        for(int i = 1; i < column; i++)
        {
            worksheet.Cells[1, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        //Define propriedades da planilha
        worksheet.TabColor = System.Drawing.Color.Black;
        worksheet.DefaultRowHeight = 12;

        //Define propriedades da primeira linha
        worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        worksheet.Row(1).Style.Font.Bold = true;

        row = 2;
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
        }

        void SetCell(object value)
        {
            worksheet.Cells[row, column++].Value = value;
        }

        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        worksheet.Cells[worksheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[worksheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[worksheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[worksheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

        if (File.Exists(caminhoPlanilha))
            File.Delete(caminhoPlanilha);

        FileStream objFileStrm = File.Create(caminhoPlanilha);
        objFileStrm.Close();

        File.WriteAllBytes(caminhoPlanilha, excel.GetAsByteArray());

        excel.Dispose();
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

