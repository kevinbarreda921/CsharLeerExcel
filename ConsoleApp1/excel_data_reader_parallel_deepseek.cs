using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ConsoleApp1
{
    public class excel_data_reader_parallel_deepseek
    {
        static excel_data_reader_parallel_deepseek()
        {
            ThreadPool.SetMinThreads(Environment.ProcessorCount, Environment.ProcessorCount);
        }

        public static void f_excel_data_reader_parallel_deepseek()
        {
            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

            if (!Directory.Exists(carpeta))
            {
                Console.WriteLine("La carpeta no existe.");
                return;
            }

            var archivos = Directory.GetFiles(carpeta, "*.xlsx");
            var resultados = new ConcurrentBag<string>();
            Stopwatch timerGlobal = Stopwatch.StartNew();

            // Procesamiento directo del ZIP (Excel es un ZIP de XMLs)
            Parallel.ForEach(archivos,
                new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount * 2 },
                ruta =>
                {
                    try
                    {
                        ProcesarExcelComoZip(ruta, resultados);
                    }
                    catch (Exception ex)
                    {
                        resultados.Add($"ERROR en {Path.GetFileName(ruta)}: {ex.Message}");
                    }
                });

            timerGlobal.Stop();

            foreach (var res in resultados)
                Console.WriteLine(res);

            Console.WriteLine("\n" + new string('=', 30));
            Console.WriteLine($"PROCESO FINALIZADO - VELOCIDAD EXTREMA");
            Console.WriteLine($"Archivos procesados: {archivos.Length}");
            Console.WriteLine($"Tiempo total: {timerGlobal.Elapsed.TotalSeconds:F2} segundos");
            Console.WriteLine(new string('=', 30));
        }

        private static void ProcesarExcelComoZip(string ruta, ConcurrentBag<string> resultados)
        {
            string fileName = Path.GetFileName(ruta);

            // Leer archivo completo en memoria (más rápido para archivos < 100MB)
            byte[] zipBytes = File.ReadAllBytes(ruta);

            using var memoryStream = new MemoryStream(zipBytes);
            using var archive = new ZipArchive(memoryStream, ZipArchiveMode.Read);

            // Leer shared strings primero
            var sharedStrings = LeerSharedStrings(archive);

            // Procesar cada hoja
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("xl/worksheets/sheet") &&
                    entry.FullName.EndsWith(".xml"))
                {
                    string sheetName = ObtenerNombreHoja(archive, entry.FullName);

                    using var stream = entry.Open();
                    using var reader = XmlReader.Create(stream,
                        new XmlReaderSettings { IgnoreWhitespace = true });

                    while (reader.Read())
                    {
                        // Buscar fila 3 directamente
                        if (reader.IsStartElement("row") &&
                            reader.GetAttribute("r") == "3")
                        {
                            LeerCeldaO3(reader, sharedStrings, sheetName, fileName, resultados);
                            break;
                        }
                    }
                }
            }
        }

        private static Dictionary<int, string> LeerSharedStrings(ZipArchive archive)
        {
            var sharedStrings = new Dictionary<int, string>();
            var sstEntry = archive.GetEntry("xl/sharedStrings.xml");

            if (sstEntry == null) return sharedStrings;

            using var stream = sstEntry.Open();
            using var reader = XmlReader.Create(stream,
                new XmlReaderSettings { IgnoreWhitespace = true });

            int index = 0;
            while (reader.Read())
            {
                if (reader.IsStartElement("t"))
                {
                    sharedStrings[index++] = reader.ReadElementContentAsString();
                }
            }

            return sharedStrings;
        }

        private static string ObtenerNombreHoja(ZipArchive archive, string sheetPath)
        {
            // Extraer número de hoja del nombre del archivo
            string sheetNumber = Path.GetFileNameWithoutExtension(sheetPath)
                .Replace("sheet", "");

            // Leer workbook.xml para mapear número a nombre
            var wbEntry = archive.GetEntry("xl/workbook.xml");
            if (wbEntry == null) return sheetNumber;

            using var stream = wbEntry.Open();
            using var reader = XmlReader.Create(stream,
                new XmlReaderSettings { IgnoreWhitespace = true });

            int count = 0;
            while (reader.Read())
            {
                if (reader.IsStartElement("sheet"))
                {
                    count++;
                    if (count.ToString() == sheetNumber)
                    {
                        return reader.GetAttribute("name") ?? sheetNumber;
                    }
                }
            }

            return sheetNumber;
        }

        private static void LeerCeldaO3(XmlReader reader,
            Dictionary<int, string> sharedStrings,
            string sheetName,
            string fileName,
            ConcurrentBag<string> resultados)
        {
            while (reader.Read())
            {
                if (reader.IsStartElement("c") &&
                    reader.GetAttribute("r") == "O3")
                {
                    string cellType = reader.GetAttribute("t");
                    string value;

                    if (cellType == "s") // Shared string
                    {
                        reader.ReadToDescendant("v");
                        if (int.TryParse(reader.ReadElementContentAsString(), out int index) &&
                            sharedStrings.TryGetValue(index, out string ssValue))
                        {
                            value = ssValue;
                        }
                        else
                        {
                            value = string.Empty;
                        }
                    }
                    else
                    {
                        reader.ReadToDescendant("v");
                        value = reader.ReadElementContentAsString();
                    }

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Limpiar fecha
                        string dato = value.Contains(" 00:00")
                            ? value.Replace(" 00:00", "").Trim()
                            : value.Trim();

                        resultados.Add($"Hoja: {sheetName} | Valor: {dato} | Archivo: {fileName}");
                    }

                    break;
                }

                // Si llegamos al final de la fila, salir
                if (reader.IsStartElement("row") && reader.GetAttribute("r") != "3")
                {
                    break;
                }
            }
        }
    }
}