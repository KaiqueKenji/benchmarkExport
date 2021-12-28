using ExportsJuntos.Fakes;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExportsJuntos.Infra.Files.Writer;

public static class NpoiLib
{
    public static async Task<byte[]> CreateExcelFileAsync()
    {
        var portfolios = PortfolioGenerator.GetRandom(1000);
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Teste");

        int rowNumber = 0;

        var row = sheet.CreateRow(rowNumber);

        XSSFCellStyle cellStyleHeader = (XSSFCellStyle)workbook.CreateCellStyle();
        cellStyleHeader.Alignment = HorizontalAlignment.Center;
        cellStyleHeader.BorderLeft = BorderStyle.Thin;
        cellStyleHeader.BorderTop = BorderStyle.Thin;
        cellStyleHeader.BorderRight = BorderStyle.Thin;
        cellStyleHeader.BorderBottom = BorderStyle.Thin;

        XSSFCellStyle cellStyleBody = (XSSFCellStyle)workbook.CreateCellStyle();
        cellStyleBody.Alignment = HorizontalAlignment.Center;
        cellStyleBody.BorderLeft = BorderStyle.Thin;
        cellStyleBody.BorderTop = BorderStyle.Thin;
        cellStyleBody.BorderRight = BorderStyle.Thin;
        cellStyleBody.BorderBottom = BorderStyle.Thin;

        XSSFFont font = (XSSFFont)workbook.CreateFont();
        font.FontHeight = 220;
        font.IsBold = true;
        cellStyleHeader.SetFont(font);

        ICell cell;

        var headers = GetPortfolioHeaders();
        for (int i = 0; i < headers.Length; i++)
        {
            cell = row.CreateCell(i);
            cell.SetCellValue(headers[i]);
            cell.CellStyle = cellStyleHeader;
        }
        var column = 0;
        rowNumber++;
        foreach (var portfolio in portfolios)
        {
            column = 0;
            row = sheet.CreateRow(rowNumber);
            SetCell(portfolio.PortfolioName);
            SetCell(portfolio.Description);
            SetCell(portfolio.CurrencyId.ToString());
            SetCell(portfolio.AssetName);
            SetCell(portfolio.Date.ToString());
            SetCell(portfolio.PortfolioTypeId.ToString());
            SetCell(portfolio.TenantId.ToString());
            SetCell(portfolio.ManagerId.ToString());
            SetCell(portfolio.CountryId.ToString());
            SetCell(portfolio.Name);
            SetCell(portfolio.Nickname);
            SetCell(portfolio.BirthDate.ToString());
            SetCell(portfolio.Age.ToString());
            SetCell(portfolio.Document);
            rowNumber++;

            void SetCell(string value)
            {
                var cells = row.CreateCell(column);
                cells.SetCellValue(value);
                cells.CellStyle = cellStyleBody;
                sheet.SetColumnWidth(column, 5500);
                column++;
            }
        }

        byte[] byteArray;
        using (var stream = new MemoryStream())
        {
            workbook.Write(stream);
            byteArray = stream.ToArray();
        }

        return await Task.FromResult(byteArray);
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