using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class excel_data_reader_parallel_gemini_mejora
    {
        public static void f_excel_data_reader_parallel_gemini_mejora()
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
                        using var fileStream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 65536, FileOptions.SequentialScan);
                        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Read);

                        // 1. CARGAR FORMATOS DE NÚMERO (Para saber si es fecha, moneda, etc.)
                        // Diccionario de: Index de Estilo -> ID de Formato de Número
                        var cellXfs = new List<int>();
                        var stylesEntry = archive.GetEntry("xl/styles.xml");
                        if (stylesEntry != null)
                        {
                            using var sStream = stylesEntry.Open();
                            using var sReader = XmlReader.Create(sStream);
                            while (sReader.Read())
                            {
                                if (sReader.NodeType == XmlNodeType.Element && sReader.Name == "xf")
                                {
                                    string numFmtId = sReader.GetAttribute("numFmtId");
                                    if (numFmtId != null) cellXfs.Add(int.Parse(numFmtId));
                                }
                            }
                        }

                        // 2. CARGAR SHARED STRINGS (Para textos)
                        List<string> sharedStrings = new List<string>();
                        var ssEntry = archive.GetEntry("xl/sharedStrings.xml");
                        if (ssEntry != null)
                        {
                            using var ssStream = ssEntry.Open();
                            using var ssReader = XmlReader.Create(ssStream);
                            while (ssReader.Read())
                            {
                                if (ssReader.NodeType == XmlNodeType.Element && ssReader.Name == "t")
                                    sharedStrings.Add(ssReader.ReadElementContentAsString());
                            }
                        }

                        // 3. LEER LA HOJA
                        var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml");
                        if (sheetEntry != null)
                        {
                            using var sheetStream = sheetEntry.Open();
                            using var xmlReader = XmlReader.Create(sheetStream);

                            bool found = false;
                            while (xmlReader.Read() && !found)
                            {
                                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "row" && xmlReader.GetAttribute("r") == "3")
                                {
                                    while (xmlReader.Read())
                                    {
                                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "c")
                                        {
                                            if (xmlReader.GetAttribute("r") == "O3")
                                            {
                                                string tipoDato = xmlReader.GetAttribute("t");
                                                string styleIndexStr = xmlReader.GetAttribute("s"); // Índice de estilo
                                                
                                                xmlReader.ReadToDescendant("v");
                                                string valorCrudo = xmlReader.ReadElementContentAsString();
                                                if (double.TryParse(valorCrudo, out double d))
                                                {
                                                    // Esto es lo que haría el GetValue() internamente
                                                    string dato = DateTime.FromOADate(d).ToShortDateString();
                                                    resultados.Add($"Valor: {dato} | Archivo: {Path.GetFileName(ruta)}");
                                                }
                                                found = true; break;
                                            }
                                        }
                                        if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "row") break;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex) { resultados.Add($"ERROR {Path.GetFileName(ruta)}: {ex.Message}"); }
                });

            timerGlobal.Stop();
            StringBuilder sb = new StringBuilder();
            foreach (var res in resultados.OrderBy(x => x)) sb.AppendLine(res);
            Console.Write(sb.ToString());
            Console.WriteLine($"Tiempo: {timerGlobal.Elapsed.TotalSeconds:F2}s");
        }
    }
}