using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;

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
                    try 
                    {
                        DateTime year;
                        if (DateTime.TryParseExact(args[0], "yyyy", enUS, DateTimeStyles.None, out year))
                        {
                            Console.WriteLine("Ingrese la ruta de guardado del archivo: Ejemplo: C:\\Directorio\\excel \n");
                            var path = Directory.GetCurrentDirectory();//Console.ReadLine();
                            if (path != null)
                            {
                                if (!string.IsNullOrEmpty(path.Trim()))
                                {
                                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                                    using (ExcelPackage excelPackage = new ExcelPackage())
                                    {
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
                                                if (DateTime.Now.AddDays(i).ToString("dddd") == "lunes")
                                                {
                                                    day = DateTime.Now.AddDays(i);
                                                    break;
                                                }
                                            }
                                        }
                                        for (int i = 0; i < 12; i++)
                                        {
                                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(month.AddMonths(i).ToString("MMMM"));
                                            for (int j = 1; j <= 7; j++)
                                            {
                                                worksheet.Cells[1, j].Value = day.AddDays(j - 1).ToString("dddd");
                                                worksheet.Cells[1, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                worksheet.Cells[1, j].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                                                worksheet.Cells[1, j].Style.Font.Color.SetColor(Color.White);
                                            }
                                            string firstDay = new DateTime(year.Year, month.AddMonths(i).Month, 1).ToString("dddd");
                                            
                                            int intDay = 0;
                                            int intDay2 = 0;
                                            switch (firstDay)
                                            {
                                                case "lunes":
                                                    intDay = 0;
                                                    intDay2 = 7;
                                                    break;
                                                case "martes":
                                                    intDay = -1;
                                                    intDay2 = 6;
                                                    break;
                                                case "miércoles":
                                                    intDay = -2;
                                                    intDay2 = 5;
                                                    break;
                                                case "jueves":
                                                    intDay = -3;
                                                    intDay2 = 4;
                                                    break;
                                                case "viernes":
                                                    intDay = -4;
                                                    intDay2 = 3;
                                                    break;
                                                case "sábado":
                                                    intDay = -5;
                                                    intDay2 = 2;
                                                    break;
                                                case "domingo":
                                                    intDay = -6;
                                                    intDay2 = 1;
                                                    break;
                                            }
                                            for (int j = 1; j <= 7; j++)
                                            {
                                                if (intDay < 0)
                                                {
                                                    worksheet.Cells[2, j].Value = year.AddMonths(i).AddDays(intDay).ToString("dd");
                                                    worksheet.Cells[2, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    worksheet.Cells[2, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                                    worksheet.Cells[2, j].Style.Font.Color.SetColor(Color.Black);
                                                }
                                                else
                                                {
                                                    worksheet.Cells[2, j].Value = year.AddMonths(i).AddDays(intDay).ToString("dd");
                                                    worksheet.Cells[2, j].Style.Font.Color.SetColor(Color.Black);
                                                }
                                                intDay++;
                                            }
                                            string lastDayOfMonth = month.AddMonths(i + 1).AddDays(-month.AddMonths(i).Day).ToString("dd");
                                            bool flag = false;
                                            for (int z = 3; z < 8; z++)
                                            {
                                                for (int j = 1; j <= 7; j++)
                                                {

                                                    if (!flag)
                                                    {
                                                        worksheet.Cells[z, j].Value = year.AddMonths(i).AddDays(intDay2).ToString("dd");
                                                        worksheet.Cells[z, j].Style.Font.Color.SetColor(Color.Black);
                                                    }
                                                    else
                                                    {
                                                        worksheet.Cells[z, j].Value = year.AddMonths(i).AddDays(intDay2).ToString("dd");
                                                        worksheet.Cells[z, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                        worksheet.Cells[z, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                                        worksheet.Cells[z, j].Style.Font.Color.SetColor(Color.Black);
                                                    }
                                                    if (lastDayOfMonth == year.AddMonths(i).AddDays(intDay2).ToString("dd"))
                                                        flag = true;

                                                    intDay2++;
                                                }
                                            }  
                                        }
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
                    catch (Exception ex)
                    { 
                        run = false; 
                        Console.WriteLine(ex.Message);
                    }
                    
                }
                
            }
            else
                Console.WriteLine("Ingrese un año.");
        }
    }
}