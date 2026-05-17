using ExcelDataReader;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;

namespace ConsoleApp1
{
    public class excel_data_reader_parallel
    {
        public static void f_excel_data_reader_parallel() {


            // Configuración necesaria para manejar codificaciones de Excel (solo una vez)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));

            // Combinamos con tu carpeta de archivos
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");


            if (!Directory.Exists(carpeta))
            {
                Console.WriteLine("La carpeta no existe.");
                return;
            }

            var archivos = Directory.GetFiles(carpeta, "*.xlsx");
            var resultados = new ConcurrentBag<string>();
            Stopwatch timerGlobal = Stopwatch.StartNew();

            // Procesamiento en paralelo real
            Parallel.ForEach(archivos, (ruta) =>
            {
                try
                {
                    using var stream = File.Open(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var reader = ExcelReaderFactory.CreateReader(stream);

                    // --- BUCLE PARA RECORRER TODAS LAS HOJAS ---
                    do
                    {
                        string nombreHoja = reader.Name;
                        int filaActual = 1;

                        // --- BUCLE PARA RECORRER LAS FILAS DE LA HOJA ACTUAL ---
                        while (reader.Read())
                        {
                            if (filaActual == 3)
                            {
                                var valor = reader.GetValue(14); // Columna O

                                if (valor != null)
                                {
                                    string dato = (valor is DateTime dt)
                                        ? dt.ToString("dd/MM/yyyy")
                                        : valor.ToString().Replace(" 00:00:00", "").Trim();

                                    resultados.Add($"Archivo: {Path.GetFileName(ruta)} | Hoja: {nombreHoja} | Valor: {dato}");
                                }

                                // IMPORTANTE: Ya encontramos la fila 3, dejamos de leer ESTA hoja
                                // y saltamos a la siguiente con NextResult()
                            
                            }
                            if (filaActual >= 7 && filaActual <= 10)
                            {
                                var valor = reader.GetValue(15); // Columna P (índice 15)

                                if (valor != null)
                                {
                                    string dato = valor.ToString().Replace(" 00:00:00", "").Trim();
                                    resultados.Add($"Archivo: {Path.GetFileName(ruta)} | Hoja: {nombreHoja} | Fila: {filaActual} | Valor: {dato}");
                                }
                                var valors = reader.GetValue(14); // Columna P (índice 15)

                                if (valors != null)
                                {
                                    string dato = valors.ToString().Replace(" 00:00:00", "").Trim();
                                    resultados.Add($"Archivo: {Path.GetFileName(ruta)} | Hoja: {nombreHoja} | Fila: {filaActual} | Valor: {dato}");
                                }
                            }

                            // Si ya pasamos la fila 10, no tiene sentido seguir leyendo esta hoja
                            if (filaActual > 10) break;
                            filaActual++;
                        }
                    } while (reader.NextResult()); // Mueve el lector a la siguiente hoja
                }
                catch (Exception ex)
                {
                    resultados.Add($"Error en {Path.GetFileName(ruta)}: {ex.Message}");
                }
            });
            timerGlobal.Stop();

            foreach (var res in resultados) 
                Console.WriteLine(res);

            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO");
            Console.WriteLine($"Archivos procesados: {archivos.Length}");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));
        }
    }
}
