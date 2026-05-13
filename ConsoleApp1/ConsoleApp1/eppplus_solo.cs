using ExcelDataReader;
using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Concurrent;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics;
using System.Text;
using System.Text;

namespace ConsoleApp1
{
    public class eppplus_solo
    {
        public static void f_eppplus_solo() {


            ExcelPackage.License.SetNonCommercialPersonal("Tu Nombre o Proyecto");

            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));

            // Combinamos con tu carpeta de archivos
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");


            var archivos = Directory.GetFiles(carpeta, "*.xlsx");

            Stopwatch timerGlobal = Stopwatch.StartNew();

            foreach (var ruta in archivos)
            {
                Stopwatch timerArchivo = Stopwatch.StartNew();

                try
                {
                    using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var package = new ExcelPackage(stream);

                    var hojas = package.Workbook.Worksheets;
                    int totalHojas = hojas.Count;

                    for (int i = 0; i < totalHojas; i++)
                    {
                        var hoja = hojas[i];
                        var celda = hoja.Cells[3, 15];

                        if (celda.Value != null)
                        {
                            string datoImportante = celda.Text.Replace(" 00:00", "").Trim();
                            Console.WriteLine($"Hoja: {hoja.Name} | Valor: {datoImportante}");
                        }
                    }

                    timerArchivo.Stop();
                    Console.WriteLine($"Archivo: {Path.GetFileName(ruta)} - Tiempo: {timerArchivo.ElapsedMilliseconds} ms");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error en {Path.GetFileName(ruta)}: {ex.Message}");
                }
            }

            timerGlobal.Stop();

            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO");
            Console.WriteLine($"Archivos procesados: {archivos.Length}");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));
        }
    }
}
