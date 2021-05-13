using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadWriteConsole
{
    class Program
    {
        
        static List<Employee> employees;
        static void Main(string[] args)
        {
            Console.WriteLine("---- Reading Data ----");
            employees = new List<Employee>();


            //Create COM Objects. Create a COM object for everything that is referenced
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //change below path according to your system the employee file is zipped along with
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\employee.csv");
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            Employee emp;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (i != 1 & i <= 17)
                    {
                        emp = new Employee();
                        emp.Lastname = xlRange.Cells[i, j].Value2;
                        emp.Firstname = xlRange.Cells[i, j + 1].Value2;
                        emp.DateOfBirth = DateTime.FromOADate(xlRange.Cells[i, j + 2].Value2);
                        var today = DateTime.Now;
                        var m = (today.Year * 100 + today.Month) * 100 + emp.DateOfBirth.Day;
                        var y = (emp.DateOfBirth.Year * 100 + emp.DateOfBirth.Month) * 100 + emp.DateOfBirth.Day;
                        emp.age = (m - y) / 10000;
                        emp.Group = xlRange.Cells[i, j + 3].Value2;
                        emp.Money = double.Parse(xlRange.Cells[i, j + 4].Value2.ToString());
                        employees.Add(emp);


                        break;
                        //employees.Add(new Employee { Lastname = (xlRange.Cells[i, j]).Value2,Firstname= (xlRange.Cells[i, j]).Value2,DateOfBirth= });
                    }
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("Read Success");
            StringBuilder sb = new StringBuilder();
            int num_employeA = 0, num_employeB = 0, num_employeC = 0, num_employeD = 0;
            double a, b, c, d;

            a = employees.Where(item => item.Group == "A").Sum(item => item.Money);
            b = employees.Where(item => item.Group == "B").Sum(item => item.Money);
            c = employees.Where(item => item.Group == "C").Sum(item => item.Money);
            d = employees.Where(item => item.Group == "D").Sum(item => item.Money);

            num_employeA = employees.Where(item => item.Group == "A").Count();
            num_employeB = employees.Where(item => item.Group == "B").Count();
            num_employeC = employees.Where(item => item.Group == "C").Count();
            num_employeD = employees.Where(item => item.Group == "D").Count();
            var line = string.Format("{0},{1},{2},{3},{4},{5}", "Group", "num_employee", "min_age", "max_age", "ave_age", "sum_money");
            sb.AppendLine(line);
            line = string.Format("{0},{1},{2},{3},{4},{5}", "A", num_employeA, employees.Where(item => item.Group == "A").Min(item => item.age), employees.Where(item => item.Group == "A").Max(item => item.age), employees.Where(item => item.Group == "A").Average(item => item.age), a);
            sb.AppendLine(line);
            line = string.Format("{0},{1},{2},{3},{4},{5}", "B", num_employeB, employees.Where(item => item.Group == "B").Min(item => item.age), employees.Where(item => item.Group == "B").Max(item => item.age), employees.Where(item => item.Group == "B").Average(item => item.age), b);
            sb.AppendLine(line);
            line = string.Format("{0},{1},{2},{3},{4},{5}", "C", num_employeC, employees.Where(item => item.Group == "C").Min(item => item.age), employees.Where(item => item.Group == "C").Max(item => item.age), employees.Where(item => item.Group == "C").Average(item => item.age), c);
            sb.AppendLine(line);
            line = string.Format("{0},{1},{2},{3},{4},{5}", "D", num_employeD, employees.Where(item => item.Group == "D").Min(item => item.age), employees.Where(item => item.Group == "D").Max(item => item.age), employees.Where(item => item.Group == "D").Average(item => item.age), d);
            sb.AppendLine(line);
            line = string.Format("{0},{1},{2},{3},{4},{5}", "Total", num_employeA + num_employeB + num_employeC + num_employeD,
                (employees.Where(item => item.Group == "A").Min(item => item.age) + employees.Where(item => item.Group == "B").Min(item => item.age) + employees.Where(item => item.Group == "C").Min(item => item.age) + employees.Where(item => item.Group == "D").Min(item => item.age)),
                (employees.Where(item => item.Group == "A").Max(item => item.age) + employees.Where(item => item.Group == "B").Max(item => item.age) + employees.Where(item => item.Group == "C").Max(item => item.age) + employees.Where(item => item.Group == "D").Max(item => item.age)),
                (employees.Where(item => item.Group == "A").Average(item => item.age) + employees.Where(item => item.Group == "B").Average(item => item.age) + employees.Where(item => item.Group == "C").Average(item => item.age) + employees.Where(item => item.Group == "D").Average(item => item.age)),
                (a / num_employeA) + (b / num_employeB) + (c / num_employeC) + (d / num_employeD));
            sb.AppendLine(line);
            File.WriteAllText("group.csv", sb.ToString());
            Console.WriteLine("Write Success");
            Console.ReadKey();
        }
    }
}
