using ExcelDataReader;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;

namespace ConsoleApp1
{
    public class excel_data_reader_parallel_claude
    {
        static excel_data_reader_parallel_claude()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        public static void f_excel_data_reader_parallel_claude() {

            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

            if (!Directory.Exists(carpeta))
            {
                Console.WriteLine("La carpeta no existe.");
                return;
            }

            // Usar array directo, más rápido que IEnumerable para Parallel.ForEach
            string[] archivos = Directory.GetFiles(carpeta, "*.xlsx");

            if (archivos.Length == 0)
            {
                Console.WriteLine("No se encontraron archivos .xlsx.");
                return;
            }

            // Pre-alocar capacidad en el bag equivalente (usamos List con lock es más rápido)
            // pero ConcurrentBag está bien para escrituras concurrentes
            var resultados = new ConcurrentBag<string>();

            // Calcular grado óptimo: I/O bound = más threads que cores ayuda
            // Para lectura de disco SSD: 2x cores; HDD: igual a cores
            int gradoParalelismo = Environment.ProcessorCount * 2;

            // Opciones de ExcelDataReader preconfiguradas (reutilizables)
            var excelConfig = new ExcelReaderConfiguration
            {
                // Desactiva conversión de tipos automática → más rápido
                FallbackEncoding = Encoding.UTF8,
            };

            Stopwatch timerGlobal = Stopwatch.StartNew();

            Parallel.ForEach(
                archivos,
                new ParallelOptions { MaxDegreeOfParallelism = gradoParalelismo },
                (ruta) =>
                {
                    try
                    {
                        // SequentialScan: hint al OS para leer el archivo secuencialmente
                        // Optimiza el prefetch del disco
                        using var stream = new FileStream(
                            ruta,
                            FileMode.Open,
                            FileAccess.Read,
                            FileShare.Read,
                            bufferSize: 65536,           // 64KB buffer → reduce llamadas al OS
                            FileOptions.SequentialScan    // Hint de lectura secuencial
                        );

                        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream, excelConfig);

                        string nombreArchivo = Path.GetFileName(ruta); // Cache del nombre

                        do
                        {
                            string nombreHoja = reader.Name;
                            int filaActual = 0;

                            while (reader.Read())
                            {
                                filaActual++;

                                if (filaActual == 3)
                                {
                                    var valor = reader.GetValue(14);

                                    if (valor != null)
                                    {
                                        // Evitar ToString() + Replace() en cadena
                                        // Usar ReadOnlySpan cuando sea posible
                                        string raw = valor.ToString()!;
                                        string dato = raw.Contains(" 00:00", StringComparison.Ordinal)
                                            ? raw.Replace(" 00:00", "", StringComparison.Ordinal).Trim()
                                            : raw.Trim();

                                        // String interpolation con variables ya cacheadas
                                        resultados.Add($"Hoja: {nombreHoja} | Valor: {dato} | Archivo: {nombreArchivo}");
                                    }
                                    break; // Salir del while inmediatamente
                                }

                                // Optimización: si ya pasamos la fila 3, no seguir leyendo
                                if (filaActual > 3) break;
                            }

                        } while (reader.NextResult());
                    }
                    catch (Exception ex)
                    {
                        resultados.Add($"ERROR en {Path.GetFileName(ruta)}: {ex.Message}");
                    }
                });

            timerGlobal.Stop();

            // Escribir todos los resultados de una sola vez usando StringBuilder
            // Evita múltiples llamadas a Console.WriteLine (cada una hace flush)
            var sb = new StringBuilder(resultados.Count * 80); // Pre-alocar memoria estimada
            foreach (var res in resultados)
                sb.AppendLine(res);

            sb.AppendLine();
            sb.AppendLine(new string('=', 30));
            sb.AppendLine("PROCESO FINALIZADO");
            sb.AppendLine($"Archivos procesados: {archivos.Length}");
            sb.AppendLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            sb.AppendLine(new string('=', 30));

            Console.Write(sb.ToString()); // Un solo write al final

        }
    }
}
