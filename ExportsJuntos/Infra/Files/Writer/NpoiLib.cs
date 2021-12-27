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

        var styleHeader = workbook.CreateCellStyle();
        var font = workbook.CreateFont();
        font.Color = HSSFColor.Blue.Index2;
        styleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
        styleHeader.FillPattern = FillPattern.SolidForeground;

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

        XSSFFont ffont = (XSSFFont)workbook.CreateFont();
        ffont.FontHeight = 220;
        ffont.IsBold = true;
        cellStyleHeader.SetFont(ffont);

        ICell cell;

        var headers = GetPortfolioHeaders();
        for (int i = 0; i < headers.Length; i++)
        {
            cell = row.CreateCell(i);
            cell.SetCellValue(headers[i]);
            cell.CellStyle = styleHeader;
            cell.CellStyle = cellStyleHeader;
        }
        var column = 0;
        rowNumber++;
        foreach (var portfolio in portfolios)
        {
            column = 0;
            CreateNewRow(rowNumber);
            SetCell(portfolio.PortfolioName, 0);
            SetCell(portfolio.Description, 1);
            SetCell(portfolio.CurrencyId.ToString(), 2);
            SetCell(portfolio.AssetName, 3);
            SetCell(portfolio.Date.ToString(), 4);
            SetCell(portfolio.PortfolioTypeId.ToString(), 5);
            SetCell(portfolio.TenantId.ToString(), 6);
            SetCell(portfolio.ManagerId.ToString(), 7);
            SetCell(portfolio.CountryId.ToString(), 8);
            SetCell(portfolio.Name, 9);
            SetCell(portfolio.Nickname, 10);
            SetCell(portfolio.BirthDate.ToString(), 11);
            SetCell(portfolio.Age.ToString(), 12);
            SetCell(portfolio.Document, 13);
            rowNumber++;

            void SetCell(string value, int columns)
            {
                var cells = row.CreateCell(columns);
                cells.SetCellValue(value);
                cells.CellStyle = cellStyleBody;
                //setCellStyle(workbook, cells);
                sheet.SetColumnWidth(columns, 5500);
            }

            void CreateNewRow(int rowNumbers)
            {
                row = sheet.CreateRow(rowNumbers);
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