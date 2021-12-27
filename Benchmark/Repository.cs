﻿using ClosedXML.Excel;
using ExportsJuntos.Models;
using Ganss.Excel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Benchmark;

public class Repository
{
    public void CriarPlanilhaEPPlus(IEnumerable<Portfolio> portfolios, string caminhoPlanilha)
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

        for (int i = 1; i < column; i++)
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

    public void CriarPlanilhaInterop(IEnumerable<Portfolio> portfolios, string filePath)
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

        if (File.Exists(filePath))
            File.Delete(filePath);

        xlWorkBook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        xlWorkBook.Close(true, misValue, misValue);
        xlApp.Quit();

        Marshal.ReleaseComObject(xlWorkSheet);
        Marshal.ReleaseComObject(xlWorkBook);
        Marshal.ReleaseComObject(xlApp);

        byte[] byteArray;
        using (var stream = new MemoryStream())
        {
            byteArray = stream.ToArray();
        }

        File.WriteAllBytes(@"C:\dev\Excel\TesteInterop.xlsx", byteArray);
    }

    public void CriarPlanilhaMapper(IEnumerable<Portfolio> portfolios, string filePath)
    {
        ExcelMapper excel = new ExcelMapper();
        excel.Save(filePath, portfolios, "TESTE", true);

        byte[] byteArray;
        using (var stream = new MemoryStream())
        {
            byteArray = stream.ToArray();
        }
        File.WriteAllBytes(@"C:\dev\Excel\TesteMapper.xls", byteArray);
    }

    public void Write(IEnumerable<Portfolio> portfolios)
    {
        var stream = new MemoryStream();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet 1");
        FillWorksheet(worksheet, portfolios);
        byte[] byteArray;
        workbook.SaveAs(stream);
        byteArray = stream.ToArray();
        File.WriteAllBytes(@"C:\dev\Excel\TesteClosed.xlsx", byteArray);
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

    public void CriarPlanilhaNpoi(IEnumerable<Portfolio> portfolios)
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Teste");

        int rowNumber = 0;

        var row = sheet.CreateRow(rowNumber);

        var styleHeader = workbook.CreateCellStyle();
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
        File.WriteAllBytes(@"C:\dev\Excel\TesteNpoi.xlsx", byteArray);
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