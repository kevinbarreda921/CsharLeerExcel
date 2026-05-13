using ExcelDataReader;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Diagnostics;
using System.Collections.Concurrent;
using System.Text;
using ExcelDataReader;

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
            Parallel.ForEach(archivos, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, (ruta) =>
            {
                try
                {
                    // Abrimos el archivo en modo compartido y solo lectura para máxima velocidad de I/O
                    using var stream = File.Open(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                    // Creamos el lector optimizado para OpenXml (.xlsx)
                    using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    do
                    {
                        // Nombre de la hoja actual
                        string nombreHoja = reader.Name;
                        int filaActual = 0;

                        // Leemos fila por fila (Forward-only)
                        while (reader.Read())
                        {
                            filaActual++;

                            // Buscamos específicamente la Fila 3
                            if (filaActual == 3)
                            {
                                // La columna 15 (O) es el índice 14 en base 0
                                var valor = reader.GetValue(14);

                                if (valor != null)
                                {
                                    string dato = valor.ToString()
                                                      .Replace(" 00:00", "")
                                                      .Trim();

                                    resultados.Add($"Hoja: {nombreHoja} | Valor: {dato} | Archivo: {Path.GetFileName(ruta)}");
                                }
                                // Una vez encontrada la fila 3 de esta hoja, saltamos al siguiente ciclo/hoja
                                break;
                            }
                        }
                    } while (reader.NextResult()); // Pasa a la siguiente pestaña del Excel
                }
                catch (Exception ex)
                {
                    resultados.Add($"ERROR en {Path.GetFileName(ruta)}: {ex.Message}");
                }
            });

            timerGlobal.Stop();

            // Impresión de resultados masiva
            Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = false });
            foreach (var res in resultados) Console.WriteLine(res);
            Console.Out.Flush();

            var standardOutput = new StreamWriter(Console.OpenStandardOutput());
            standardOutput.AutoFlush = true;
            Console.SetOut(standardOutput);

            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO");
            Console.WriteLine($"Archivos procesados: {archivos.Length}");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));
        }
    }
}
