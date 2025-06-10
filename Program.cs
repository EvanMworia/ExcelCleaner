using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

class Person
{
    public string Name { get; set; }
    public string Phone { get; set; }
    public string Address { get; set; }
}

class Program
{
  

    static void Main()
    {
        //using EPPlus 8+, no longer uses ExcelPackage.LicenseContext.
        // Instead, it expects the license to be set via a static property:
        ExcelPackage.License = new OfficeOpenXml.LicenseProvider.License
        {
            LicenseType = OfficeOpenXml.LicenseType.NonCommercial
        };

        var inputPath = "input.xlsx";
        var outputPath = "cleaned.xlsx";

        var people = ReadExcel(inputPath);
        var cleaned = CleanData(people);
        WriteExcel(cleaned, outputPath);

        Console.WriteLine("✅ Cleaning complete. Check cleaned.xlsx");
    }


    static List<Person> ReadExcel(string path)
    {
        var list = new List<Person>();

        using var package = new ExcelPackage(new FileInfo(path));
        var worksheet = package.Workbook.Worksheets[0];
        var rowCount = worksheet.Dimension.Rows;

        for (int row = 2; row <= rowCount; row++)
        {
            var person = new Person
            {
                Name = worksheet.Cells[row, 1].Text,
                Phone = worksheet.Cells[row, 2].Text,
                Address = worksheet.Cells[row, 3].Text
            };
            list.Add(person);
        }

        return list;
    }

    static List<Person> CleanData(List<Person> data)
    {
        var cleaned = new List<Person>();

        foreach (var p in data)
        {
            bool isAddressNyeri = !string.IsNullOrWhiteSpace(p.Address) &&
                                  p.Address.Trim().ToLower() == "nyeri";

            bool hasPhone = !string.IsNullOrWhiteSpace(p.Phone);

            if (isAddressNyeri && hasPhone)
            {
                cleaned.Add(p);
            }
        }

        return cleaned;
    }

    static void WriteExcel(List<Person> people, string path)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Cleaned");

        worksheet.Cells[1, 1].Value = "Name";
        worksheet.Cells[1, 2].Value = "Phone";
        worksheet.Cells[1, 3].Value = "Address";

        for (int i = 0; i < people.Count; i++)
        {
            worksheet.Cells[i + 2, 1].Value = people[i].Name;
            worksheet.Cells[i + 2, 2].Value = people[i].Phone;
            worksheet.Cells[i + 2, 3].Value = people[i].Address;
        }

        package.SaveAs(new FileInfo(path));
    }
}
