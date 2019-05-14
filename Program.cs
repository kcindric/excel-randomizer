using System;
using System.Linq;
using LinqToExcel;
using ClosedXML.Excel;

//Creating a class with property names which match the collumn names
//in the Excel document
public class Obj
{
    public double Area { get; set; }
    public double Length { get; set; }
}


namespace LinqToExcel.Screencast
{
    class Program
    {
        static void Main()
        {
            //Creating a connection with a excel file (LinqToExcel)
            var excel = new ExcelQueryFactory();
            excel.FileName = "X:\\filepath\\file.xlsx";

            Random rnd = new Random();

            //Randomizing the order of the query
            //using random numbers for their indexes (LinqToExcel + Linq)
            var listExcel = from x in excel.Worksheet<Obj>().ToList().OrderBy(r => rnd.Next())
                            select x;

            //Creating a new workbook (ClosedXML)
            var wb = new XLWorkbook();

            //Adding a worksheet (ClosedXML)
            var ws = wb.Worksheets.Add("Objects");

            //Adding headers (ClosedXML)
            ws.Cell("A1").Value = "Area";
            ws.Cell("B1").Value = "Length";

            int cell = 2;

            //Adding randomized values to the worksheet(ClosedXML)
            foreach (var u in listExcel)
            {
                ws.Cell($"A{cell}").Value = u.Area;
                ws.Cell($"B{cell}").Value = u.Length;
                cell++;
            }

            //Saving the workbook
            wb.SaveAs("X:\\filepath\\new_file.xlsx");
        }
    }
}
