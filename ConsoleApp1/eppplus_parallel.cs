
using OfficeOpenXml;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace ConsoleApp1
{
    public class eppplus_parallel
    {
        public static void f_eppplus_parallel() {
 
            ExcelPackage.License.SetNonCommercialPersonal("Tu Proyecto");

            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));

            // Combinamos con tu carpeta de archivos
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");


            if (!Directory.Exists(carpeta)) return;

            var archivos = Directory.GetFiles(carpeta, "*.xlsx");
            var resultados = new ConcurrentBag<string>(); // Almacenamiento seguro para hilos
            Stopwatch timerGlobal = Stopwatch.StartNew();

            // Procesamiento en paralelo para usar todos los núcleos del CPU
            Parallel.ForEach(archivos, (ruta) =>
            {
                try
                {
                    using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var package = new ExcelPackage(stream);

                    foreach (var hoja in package.Workbook.Worksheets)
                    {
                        // Acceso directo a la celda O3
                        var valor = hoja.Cells["O3"].Text;

                        if (valor != null)
                        {
                            string dato = valor.ToString().Replace(" 00:00", "").Trim();
                            resultados.Add($"Archivo: {Path.GetFileName(ruta)} | Hoja: {hoja.Name} | Valor: {dato}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    resultados.Add($"ERROR en {Path.GetFileName(ruta)}: {ex.Message}");
                }
            });

            timerGlobal.Stop();

            // Salida de resultados
            foreach (var res in resultados) Console.WriteLine(res);

            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));

        }
    }
}
