// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
Console.WriteLine("Hello, World!");
// path to your excel file
string path = "C:/Temp/Firstbook.xlsx";
FileInfo fileInfo = new FileInfo(path);

ExcelPackage package = new ExcelPackage(fileInfo);
ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

// get number of rows in the sheet
int rows = worksheet.Dimension.Rows; // 10
Console.WriteLine(rows);
int columns = worksheet.Dimension.Columns; // 10
Console.WriteLine(columns);
worksheet.Cells["D2"].Value = "ABCD"; // присвоить ячейке D2 значение
Console.WriteLine(worksheet.Cells["D2"].Value); // вывести значение ячейки
worksheet.Cells["D4"].Value = "ABC"; // присвоить ячейке D2 значение
Console.WriteLine(worksheet.Cells["D3"].Value);
worksheet = package.Workbook.Worksheets[1]; // делаем активным 2 лист (нумерация идёт с нуля)
int rows2 = worksheet.Dimension.Rows; // 10
Console.WriteLine(rows2);
int columns2 = worksheet.Dimension.Columns; // 10
Console.WriteLine(columns2);

worksheet.Cells["D4"].Value = "ABC"; // присвоить ячейке D4 второго листа значение
Console.WriteLine(worksheet.Cells["D4"].Value);

worksheet = package.Workbook.Worksheets[2];
worksheet.Cells["D4"].Value = "ABC"; // присвоить ячейке D4 третьего листа значение

worksheet = package.Workbook.Worksheets["Second"]; // делаем активным 2 лист (нумерация идёт с нуля)
worksheet.Cells["B2"].Value = "Cool"; // делаю активным второй лист по его имени

var perenos = worksheet.Cells["E3"].Value;
// Console.WriteLine(perenos);

worksheet = package.Workbook.Worksheets["First"];
worksheet.Cells["D4"].Value = perenos;
Console.WriteLine(worksheet.Cells["D4"].Value);

worksheet = package.Workbook.Worksheets["Second"];
Console.WriteLine(rows2);

object [] array1 = new object[rows2];

for (int i = 0; i < rows2; i++)
{
   
    array1[i] = worksheet.Cells[i+1,5].Value;  // Cells[x,y] x номер строки, y номер столбца
    Console.WriteLine(array1[i]);
 //   Console.WriteLine(worksheet.Cells["Ei"].Value);
}
worksheet = package.Workbook.Worksheets["First"];

for (int i = 0; i < rows2; i++)
{

    worksheet.Cells[i + 1, 5].Value = array1[i];  // Cells[x,y] x номер строки, y номер столбца
    
    }
// save changes
package.Save();