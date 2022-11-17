using System.Drawing;
using OddsBoots.Helper;
using OddsBoots.Modal;
using OfficeOpenXml;
using OfficeOpenXml.Style;

Console.Write("Please input Company csv Path:");
var companyPath = Console.ReadLine();

var companyFileInfo = new FileInfo(companyPath ?? string.Empty);

var isOk = FileHelper.Get(companyFileInfo.FullName, out Company[] companyList, out var message);
if (!isOk)
{
    Console.Write($"Company Csv Data Error:{message}");
    return;
}

Console.Write("Please Input Account Csv Path:");
var accountPath = Console.ReadLine();

var accountFileInfo = new FileInfo(accountPath ?? string.Empty);
var isOk1 = FileHelper.Get(accountFileInfo.FullName, out Account[] accountList, out var message1);
if (!isOk1)
{
    Console.WriteLine($"Account Csv Data Error:{message1}");
    return;
}

Console.WriteLine("Please Input Excel FilePath(Include name):");
var excelFileName = Console.ReadLine();

if (accountList.Any())
{
    using var package = new ExcelPackage();
    foreach (var item in accountList.Select(s => s.WebId).Distinct())
    {
        var companyName = companyList.FirstOrDefault(s => s.WebId.Equals(item))?.BrandName;

        var accountData = accountList.Where(s => s.WebId.Equals(item)).ToList();


        var License_Eur = accountData.Select(s => s.License_EUR).Sum();
      
        var YYYYMM = excelFileName?.Split('_')[1].Replace(new FileInfo(excelFileName).Extension, "");
        var year = int.Parse(YYYYMM?.Substring(0, 4) ?? string.Empty);
        var month = int.Parse(YYYYMM?.Substring(4).PadLeft(2, '0') ?? string.Empty);
        var workSheet = package.Workbook.Worksheets.Add(companyName);
        workSheet.SetValue(1, 1,
            $"{companyName}-Promotion-odds-Boosts REPORT({year}-{month}-01~{year}-{month}-{DateTime.DaysInMonth(year, month)})");
        workSheet.SetValue(1, 2, "");
        workSheet.SetValue(1, 3, "");
        workSheet.SetValue(1, 4, "");
        workSheet.SetValue(1, 5, "");
        workSheet.SetValue(1, 6, "");
        workSheet.Cells["A1:F1"].Merge = true;
        workSheet.SetValue(2, 1, "Currency");
        workSheet.SetValue(2, 2, "CompanyWinLoss");
        workSheet.SetValue(2, 3, "%");
        workSheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        workSheet.SetValue(2, 4, "Profit Sharing(Currency)");
        workSheet.SetValue(2, 5, "Licenes(Currency)");
        workSheet.SetValue(2, 6, "Licenes(EUR)");
        workSheet.Cells["A2:F2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells["A2:F2"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
        workSheet.Cells["A2:F2"].Style.Font.Color.SetColor(Color.White);

     
        for (var i = 0; i < accountData.Count; i++)
        {
            int index = i + 2;
            workSheet.SetValue(1 + index, 1, accountData[i].Currency);
            workSheet.Cells[1 + index, 1].Style.Font.Bold = true;
            workSheet.SetValue(1 + index, 2, $"({accountData[i].CompanyWinLoss})");
            workSheet.Cells[1 + index, 2].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells[1 + index, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            workSheet.SetValue(1 + index, 3, accountData[i].Percent);
            workSheet.Cells[1 + index, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            workSheet.SetValue(1 + index, 4, $"({accountData[i].ProfitSharing})");
            workSheet.Cells[1 + index, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            workSheet.Cells[1 + index, 4].Style.Font.Color.SetColor(Color.Red);
            workSheet.SetValue(1 + index, 5, $"({accountData[i].Licenes_currency})");
            workSheet.Cells[1 + index, 5].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells[1 + index, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            workSheet.SetValue(1 + index, 6, $"({accountData[i].License_EUR})");
            workSheet.Cells[1 + index, 6].Style.Font.Color.SetColor(Color.Red);
            workSheet.Cells[1 + index, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        }

        var _ = 3 + accountData.Count;
        workSheet.SetValue(_, 1, "Total EUR");
        workSheet.Cells[_, 1].Style.Font.Bold = true;
        workSheet.SetValue(_, 6, $"({License_Eur})");
        workSheet.Cells[_, 6].Style.Font.Color.SetColor(Color.Red);
        workSheet.Cells[_, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells.Style.Font.Size = 16;
        workSheet.Cells.AutoFitColumns();
    }

    package.SaveAs(excelFileName);
    package.Dispose();
    Console.WriteLine($"Excel Created Done,Path:{excelFileName}");
    Console.ReadKey();
}