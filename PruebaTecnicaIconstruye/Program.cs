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
                        var path = Console.ReadLine();
                        if (path != null)
                        {
                            if (string.IsNullOrEmpty(path.Trim())) 
                            {
                            
                            }
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