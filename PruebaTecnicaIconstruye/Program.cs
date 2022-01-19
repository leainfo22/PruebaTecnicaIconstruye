using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MyApp // Note: actual namespace depends on the project name.
{
    public class Program
    {
        private static CultureInfo enUS = new CultureInfo("en-US");
        public static void Main(string[] args)
        {
            if (!string.IsNullOrEmpty(args[0]))
            {
                bool run = true;
                while (run) 
                {
                    DateTime date;
                    if (DateTime.TryParseExact(args[0], "yyyy", enUS, DateTimeStyles.None, out date))
                    {
                        Console.WriteLine("Ingrese la ruta de guardado del archivo: \n");
                        var path = Directory.GetCurrentDirectory(); //Console.ReadLine();
                        if (path != null)
                        {
                            if (!string.IsNullOrEmpty(path.Trim())) 
                            {
                                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                                using (ExcelPackage excelPackage = new ExcelPackage())
                                {
                                    //Set some properties of the Excel document
                                    excelPackage.Workbook.Properties.Author = "Leandro Vasquez";
                                    excelPackage.Workbook.Properties.Title = "Prueba tecnica Iconstruye";
                                    excelPackage.Workbook.Properties.Created = DateTime.Now;
                                    string numberMonth = DateTime.Now.ToString("MM");
                                    DateTime month = DateTime.Now;
                                    DateTime day = DateTime.Now;


                                    if (DateTime.Now.ToString("MM") != "01")
                                    {
                                        for (int i = 1; i < 12; i++)
                                        {
                                            if (DateTime.Now.AddMonths(i).ToString("MM") == "01")
                                            {
                                                month = DateTime.Now.AddMonths(i);
                                                break;
                                            }
                                        }
                                    }

                                    if (DateTime.Now.ToString("dddd") != "lunes")
                                    {
                                        for (int i = 1; i < 7; i++)
                                        {
                                            if (DateTime.Now.AddMonths(i).ToString("dddd") == "lunes")
                                            {
                                                day = DateTime.Now.AddMonths(i);
                                                break;
                                            }
                                        }
                                    }

                                    for (int i = 0; i < 12; i++)
                                    {
                                        var a = DateTime.Now.AddMonths(i).ToString("MMMM");
                                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(month.AddMonths(i).ToString("MMMM"));
                                        //excelPackage.Workbook.Worksheets.Add(DateTime.Now.AddMonths(i).ToString("MMMM"));
                                        for (int j = 1; j < 7; j++)
                                        {
                                            worksheet.Cells["A"+j].Value = day.ToString("dddd");
                                            worksheet.Cells["A"+j].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                                        }

                                    }
                                    //Create the WorkSheet

                                    //Add some text to cell A1
                                    //worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
                                    //You could also use [line, column] notation:
                                    //worksheet.Cells[1, 2].Value = "This is cell B1!";

                                    //Save your file
                                    FileInfo fi = new FileInfo(path + "\\excelTest.xlsx");
                                    excelPackage.SaveAs(fi);
                                    run = false;
                                }
                            }
                            else
                                Console.WriteLine("Ingrese una ruta de guardado del archivo: \n");
                        }
                        else
                            Console.WriteLine("Ingrese una ruta de guardado del archivo: \n");

                    }
                    else
                        Console.WriteLine("Ingrese un año con el formato 'yyyy'. Ejemplo: 2022 ");
                }
                
            }
            else
                Console.WriteLine("Ingrese un año.");
        }
    }
}