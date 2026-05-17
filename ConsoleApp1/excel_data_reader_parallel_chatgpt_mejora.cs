using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;

namespace ConsoleApp1;

public static class excel_data_reader_parallel_chatgpt_mejora
{
    public static void Ejecutar()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string carpeta = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            @"..\..\..\ArchivosExcel\Parte_diario");

        if (!Directory.Exists(carpeta))
        {
            Console.WriteLine("Carpeta no encontrada");
            return;
        }

        string[] archivos = Directory.GetFiles(carpeta, "*.xlsx");

        Stopwatch sw = Stopwatch.StartNew();

        ConcurrentQueue<string> resultados = new();

        ParallelOptions options = new()
        {
            // Ajustar según SSD/HDD
            MaxDegreeOfParallelism = Environment.ProcessorCount * 2
        };

        Parallel.ForEach(archivos, options, archivo =>
        {
            ProcesarArchivo(archivo, resultados);
        });

        sw.Stop();

        StringBuilder sb = new(1024 * 64);

        while (resultados.TryDequeue(out string? r))
        {
            sb.AppendLine(r);
        }

        sb.AppendLine();

        sb.AppendLine($"Tiempo total: {sw.Elapsed.TotalSeconds:F2} segundos");
        sb.AppendLine($"Archivos: {archivos.Length}");

        Console.WriteLine(sb.ToString());
    }

    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private static void ProcesarArchivo(
        string ruta,
        ConcurrentQueue<string> resultados)
    {
        try
        {
            using SpreadsheetDocument doc =
                SpreadsheetDocument.Open(ruta, false);

            WorkbookPart workbookPart = doc.WorkbookPart!;

            SharedStringTable? sharedStrings =
                workbookPart.SharedStringTablePart?.SharedStringTable;

            foreach (Sheet sheet in workbookPart.Workbook.Sheets!)
            {
                WorksheetPart wsPart =
                    (WorksheetPart)workbookPart.GetPartById(sheet.Id!);

                using OpenXmlReader reader =
                    OpenXmlReader.Create(wsPart);

                int filaActual = 0;

                while (reader.Read())
                {
                    // SOLO Rows
                    if (reader.ElementType != typeof(Row))
                        continue;

                    Row row = (Row)reader.LoadCurrentElement();

                    filaActual++;

                    // SOLO fila 3
                    if (filaActual != 3)
                        continue;

                    int columnaActual = 0;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        // Columna O = índice 14
                        if (columnaActual == 14)
                        {
                            string valor =
                                ObtenerValorCelda(cell, sharedStrings);

                            if (!string.IsNullOrWhiteSpace(valor))
                            {
                                int idx = valor.IndexOf(
                                    " 00:00",
                                    StringComparison.Ordinal);

                                if (idx >= 0)
                                    valor = valor[..idx];

                                resultados.Enqueue(
                                    $"Hoja: {sheet.Name} | Valor: {valor.Trim()} | Archivo: {Path.GetFileName(ruta)}");
                            }

                            break;
                        }

                        columnaActual++;
                    }

                    // YA encontramos fila 3
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            resultados.Enqueue(
                $"ERROR {Path.GetFileName(ruta)}: {ex.Message}");
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static string ObtenerValorCelda(
        Cell cell,
        SharedStringTable? sharedStrings)
    {
        if (cell.CellValue == null)
            return string.Empty;

        string valor = cell.CellValue.InnerText;

        // SharedString optimization
        if (cell.DataType != null &&
            cell.DataType == CellValues.SharedString)
        {
            if (sharedStrings == null)
                return valor;

            return sharedStrings
                .ElementAt(int.Parse(valor))
                .InnerText;
        }

        return valor;
    }
}