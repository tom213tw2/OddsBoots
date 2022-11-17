using System.Drawing;
using OddsBoots.Helper;
using OddsBoots.Modal;
using OfficeOpenXml;
using OfficeOpenXml.Style;

Console.WriteLine("Please input Company csv Path:");
var companyPath = Console.ReadLine();

var companyFileInfo = new FileInfo(companyPath ?? string.Empty);

var isOk = FileHelper.Get(companyFileInfo.FullName, out Company[] companyList, out var message);
if (!isOk)
{
    Console.WriteLine($"Company Csv Data Error:{message}");
    return;
}

Console.WriteLine("Please Input Account Csv Path:");
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
    var workSheet = package.Workbook.Worksheets.Add("All");
    var index = 0;
    var yyyymm = excelFileName?.Split('_')[1].Replace(new FileInfo(excelFileName).Extension, "");
    var year = int.Parse(yyyymm?.Substring(0, 4));
    var month = int.Parse(yyyymm.Substring(4).PadLeft(2, '0'));
    foreach (var item in accountList.Select(s => s.WebId).Distinct())
    {
        var companyName = companyList.FirstOrDefault(s => s.WebId.Equals(item))?.BrandName;

        var accountData = accountList.Where(s => s.WebId.Equals(item)).ToList();


        var licenseEur = accountData.Select(s => s.License_EUR).Sum();


        index += 1;
        workSheet.SetValue(index, 1,
            $"{companyName}-Promotion-odds-Boosts REPORT({year}-{month}-01~{year}-{month}-{DateTime.DaysInMonth(year, month)})");
        workSheet.SetValue(index, 2, "");
        workSheet.SetValue(index, 3, "");
        workSheet.SetValue(index, 4, "");
        workSheet.SetValue(index, 5, "");
        workSheet.SetValue(index, 6, "");
        workSheet.Cells[$"A{index}:F{index}"].Merge = true;
        index += 1;
        workSheet.SetValue(index, 1, "Currency");
        workSheet.SetValue(index, 2, "CompanyWinLoss");
        workSheet.SetValue(index, 3, "%");
        workSheet.Cells[index, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        workSheet.SetValue(index, 4, "Profit Sharing(Currency)");
        workSheet.SetValue(index, 5, "Licenes(Currency)");
        workSheet.SetValue(index, 6, "Licenes(EUR)");
        workSheet.Cells[$"A{index}:F{index}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        workSheet.Cells[$"A{index}:F{index}"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
        workSheet.Cells[$"A{index}:F{index}"].Style.Font.Color.SetColor(Color.White);


        for (var i = 0; i < accountData.Count; i++)
        {
            index += i;
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

        index += 2;
        workSheet.SetValue(index, 1, "Total EUR");
        workSheet.Cells[index, 1].Style.Font.Bold = true;
        workSheet.SetValue(index, 6, $"({licenseEur})");
        workSheet.Cells[index, 6].Style.Font.Color.SetColor(Color.Red);
        workSheet.Cells[index, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        workSheet.Cells.Style.Font.Size = 16;
        workSheet.Cells.AutoFitColumns();
        index += 1;
    }

    package.SaveAs(excelFileName);
    package.Dispose();
    Console.WriteLine($"Excel Created FilePath:{excelFileName}");
    Console.ReadKey();
}