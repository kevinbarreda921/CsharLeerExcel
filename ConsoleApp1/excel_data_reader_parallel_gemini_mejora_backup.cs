using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ConsoleApp1
{
    public class excel_data_reader_parallel_gemini_mejora_backup
    {
        public static void f_excel_data_reader_parallel_gemini_mejora_backup()
        {

            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

            var archivos = Directory.GetFiles(carpeta, "*.xlsx");

            var resultados = new ConcurrentBag<string>();
            Stopwatch timerGlobal = Stopwatch.StartNew();

            Parallel.ForEach(Partitioner.Create(archivos, loadBalance: true),
                new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                (ruta) =>
                {
                    try
                    {
                        // 1. Abrimos el XLSX como si fuera un ZIP normal (súper rápido)
                        using var fileStream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.SequentialScan);
                        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Read);

                        // 2. Buscamos directamente el XML de la primera hoja
                        var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
                        if (sheetEntry != null)
                        {
                            using var sheetStream = sheetEntry.Open();

                            // 3. Usamos XmlReader (la forma más rápida y ligera de leer XML en C#)
                            using var xmlReader = XmlReader.Create(sheetStream, new XmlReaderSettings { IgnoreWhitespace = true });

                            bool found = false;
                            while (xmlReader.Read() && !found)
                            {
                                // Buscamos las etiquetas <row>
                                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "row")
                                {
                                    string numeroFila = xmlReader.GetAttribute("r");

                                    // Si llegamos a la fila 3
                                    if (numeroFila == "3")
                                    {
                                        // Leemos el interior de la fila buscando la celda "O3" (Columna 15)
                                        while (xmlReader.Read())
                                        {
                                            if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "c")
                                            {
                                                string referenciaCelda = xmlReader.GetAttribute("r");

                                                if (referenciaCelda == "O3")
                                                {
                                                    // Leemos el valor de la etiqueta <v> (Value) dentro de <c> (Cell)
                                                    xmlReader.ReadToDescendant("v");
                                                    string valorCrudo = xmlReader.ReadElementContentAsString();

                                                    // IMPORTANTE: En OpenXML, si el valor es texto, 'valorCrudo' será un número de índice.
                                                    // Ese índice apunta al archivo "xl/sharedStrings.xml" en el ZIP.
                                                    // Si es un número o fecha, estará aquí directamente.

                                                    resultados.Add($"Archivo: {Path.GetFileName(ruta)} | Valor Crudo: {valorCrudo}");
                                                    found = true;
                                                    break;
                                                }
                                            }
                                            // Si salimos de la fila 3, detenemos la búsqueda
                                            if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "row") break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        resultados.Add($"ERROR en {Path.GetFileName(ruta)}: {ex.Message}");
                    }
                });
            timerGlobal.Stop();

            // Optimización 5: Escribir a la consola en bloque. Console.WriteLine dentro de un ciclo es extremadamente lento.
            StringBuilder salidaConsola = new StringBuilder();
            foreach (var res in resultados)
            {
                salidaConsola.AppendLine(res);
            }

            salidaConsola.AppendLine(new string('=', 30));
            salidaConsola.AppendLine("PROCESO FINALIZADO");
            salidaConsola.AppendLine($"Archivos procesados: {archivos.Length}");
            salidaConsola.AppendLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            salidaConsola.AppendLine(new string('=', 30));

            Console.Write(salidaConsola.ToString());

        }
    }
}
