using ExportsJuntos.Models;
using ExportsJuntos.Repositories;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportsJuntos.Infra.Files.Writer;
public class Interop : IInteropLibRepository
{

    public void CriarPlanilha(IEnumerable<Portfolio> portfolios, string filePath)
    {
        Excel.Application xlApp = new Excel.Application();

        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        xlWorkBook = xlApp.Workbooks.Add(misValue);
        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

        var headers = GetPortfolioHeaders();
        var column = 1;
        var row = 1;
        foreach (var header in headers)
        {
            xlWorkSheet.Cells[row, column++] = header;
        }

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
            SetCell(portfolio.TenantId.ToString());
            SetCell(portfolio.ManagerId.ToString());
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
            xlWorkSheet.Cells[row, column++] = value;
        }

        xlWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        xlWorkBook.Close(true, misValue, misValue);
        xlApp.Quit();

        Marshal.ReleaseComObject(xlWorkSheet);
        Marshal.ReleaseComObject(xlWorkBook);
        Marshal.ReleaseComObject(xlApp);
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