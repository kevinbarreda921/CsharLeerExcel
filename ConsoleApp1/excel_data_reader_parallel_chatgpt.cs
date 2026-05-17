using ExcelDataReader;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;

namespace ConsoleApp1;

public static class excel_data_reader_parallel_chatgpt
{
    public static void Ejecutar()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string carpeta = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            @"..\..\..\ArchivosExcel\Parte_diario");

        if (!Directory.Exists(carpeta))
        {
            Console.WriteLine("La carpeta no existe.");
            return;
        }

        string[] archivos = Directory.GetFiles(
            carpeta,
            "*.xlsx",
            SearchOption.TopDirectoryOnly);

        Stopwatch sw = Stopwatch.StartNew();

        // Menos contención que ConcurrentBag
        ConcurrentQueue<string> resultados = new();

        // Ideal para I/O intensivo
        ParallelOptions options = new()
        {
            // SSD → 2x CPUs suele rendir mejor
            MaxDegreeOfParallelism = Environment.ProcessorCount * 2
        };

        Parallel.ForEach(archivos, options, ruta =>
        {
            ProcesarArchivo(ruta, resultados);
        });

        sw.Stop();

        // StringBuilder gigante evita miles de WriteLine
        StringBuilder sb = new(1024 * 64);

        while (resultados.TryDequeue(out string? r))
        {
            sb.AppendLine(r);
        }

        sb.AppendLine();
        sb.AppendLine(new string('=', 40));
        sb.AppendLine("PROCESO FINALIZADO");
        sb.AppendLine($"Archivos procesados: {archivos.Length}");
        sb.AppendLine($"Tiempo total: {sw.Elapsed.TotalSeconds:F2} segundos");
        sb.AppendLine(new string('=', 40));

        Console.WriteLine(sb.ToString());
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static void ProcesarArchivo(
        string ruta,
        ConcurrentQueue<string> resultados)
    {
        try
        {
            // FileOptions.SequentialScan mejora mucho lectura masiva
            using FileStream stream = new(
                ruta,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite,
                bufferSize: 1024 * 64,
                options: FileOptions.SequentialScan);

            using IExcelDataReader reader =
                ExcelReaderFactory.CreateOpenXmlReader(stream);

            do
            {
                string hoja = reader.Name;

                // SOLO leer hasta fila 3
                for (int fila = 0; fila < 3; fila++)
                {
                    if (!reader.Read())
                        break;

                    // Fila 3
                    if (fila == 2)
                    {
                        object? valor = reader.GetValue(14);

                        if (valor is null)
                            break;

                        string dato = LimpiarTexto(valor);

                        resultados.Enqueue(
                            $"Hoja: {hoja} | Valor: {dato} | Archivo: {Path.GetFileName(ruta)}");

                        break;
                    }
                }

            } while (reader.NextResult());
        }
        catch (Exception ex)
        {
            resultados.Enqueue(
                $"ERROR en {Path.GetFileName(ruta)}: {ex.Message}");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static string LimpiarTexto(object valor)
    {
        string s = valor.ToString()!;

        // Más rápido que Replace
        int idx = s.IndexOf(" 00:00", StringComparison.Ordinal);

        if (idx >= 0)
            s = s[..idx];

        return s.Trim();
    }
}