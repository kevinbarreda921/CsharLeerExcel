using System.Buffers;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Threading.Channels;
using System.Xml;

namespace ConsoleApp1
{
    public class ExcelUltraFastReader
    {
        private const int BUFFER_SIZE = 65536;
        private const int CHANNEL_CAPACITY = 128;
        private const int TARGET_ROW = 3;
        private static readonly string TARGET_COL_LETTER = "O";

        private static readonly XmlReaderSettings XmlSettings = new()
        {
            IgnoreWhitespace = true,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            DtdProcessing = DtdProcessing.Ignore,
            ValidationType = ValidationType.None,
            CheckCharacters = false
        };

        public static async Task RunAsync()
        {
            string rutaBase = AppDomain.CurrentDomain.BaseDirectory;
            string rutaProyecto = Path.GetFullPath(Path.Combine(rutaBase, @"..\..\..\"));
            string carpeta = Path.Combine(rutaProyecto, "ArchivosExcel", "Parte_diario");

            if (!Directory.Exists(carpeta))
            {
                Console.WriteLine("La carpeta no existe.");
                return;
            }

            string[] archivos = Directory.GetFiles(carpeta, "*.xlsx");
            if (archivos.Length == 0)
            {
                Console.WriteLine("No se encontraron archivos .xlsx.");
                return;
            }

            int grado = Math.Min(archivos.Length, Environment.ProcessorCount * 2);

            var channel = Channel.CreateBounded<ResultadoArchivo>(
                new BoundedChannelOptions(CHANNEL_CAPACITY)
                {
                    SingleReader = true,
                    SingleWriter = false,
                    FullMode = BoundedChannelFullMode.Wait
                });

            var resultados = new List<string>(archivos.Length);
            var sw = Stopwatch.StartNew();

            var tareaConsumo = Task.Run(async () =>
            {
                await foreach (var r in channel.Reader.ReadAllAsync())
                {
                    resultados.Add(r.Exito
                        ? $"Hoja: {r.NombreHoja} | Valor: {r.Valor} | Archivo: {r.NombreArchivo}"
                        : $"ERROR en {r.NombreArchivo}: {r.MensajeError}");
                }
            });

            await Parallel.ForEachAsync(
                archivos,
                new ParallelOptions { MaxDegreeOfParallelism = grado },
                async (ruta, ct) =>
                {
                    var resultado = ProcesarArchivo(ruta);
                    await channel.Writer.WriteAsync(resultado, ct);
                });

            channel.Writer.Complete();
            await tareaConsumo;
            sw.Stop();

            var sb = new StringBuilder(resultados.Count * 80);
            foreach (var r in resultados) sb.AppendLine(r);
            sb.AppendLine().AppendLine(new string('=', 40));
            sb.AppendLine("PROCESO FINALIZADO");
            sb.AppendLine($"Archivos procesados : {archivos.Length}");
            sb.AppendLine($"Tiempo total        : {sw.Elapsed.TotalSeconds:F2} segundos");
            sb.AppendLine($"Throughput          : {archivos.Length / sw.Elapsed.TotalSeconds:F2} archivos/seg");
            sb.AppendLine(new string('=', 40));
            Console.Write(sb);
        }

        private static ResultadoArchivo ProcesarArchivo(string ruta)
        {
            string nombreArchivo = Path.GetFileName(ruta);
            byte[] bufferRentado = ArrayPool<byte>.Shared.Rent(BUFFER_SIZE);

            try
            {
                using var fs = new FileStream(
                    ruta,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    bufferSize: 0,
                    FileOptions.SequentialScan);

                using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);

                string[]? sharedStrings = CargarSharedStrings(zip);

                ZipArchiveEntry? sheetEntry =
                    zip.GetEntry("xl/worksheets/sheet1.xml") ??
                    zip.GetEntry("xl/worksheets/Sheet1.xml") ??
                    BuscarPrimeraHoja(zip);

                if (sheetEntry == null)
                    return ResultadoArchivo.Fallo(nombreArchivo, "No se encontró sheet1.xml");

                string nombreHoja = ObtenerNombrePrimeraHoja(zip);

                var (valor, ok) = LeerCeldaObjetivo(sheetEntry, sharedStrings);

                return ok
                    ? ResultadoArchivo.Ok(nombreArchivo, nombreHoja, valor!)
                    : ResultadoArchivo.Fallo(nombreArchivo, "Celda O3 vacía o no encontrada");
            }
            catch (Exception ex)
            {
                return ResultadoArchivo.Fallo(nombreArchivo, ex.Message);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(bufferRentado);
            }
        }

        private static (string? valor, bool encontrado) LeerCeldaObjetivo(
            ZipArchiveEntry sheetEntry,
            string[]? sharedStrings)
        {
            using var stream = sheetEntry.Open();
            using var xml = XmlReader.Create(stream, XmlSettings);

            int filaActual = 0;
            bool enCeldaO = false;
            string tipoCelda = "";

            while (xml.Read())
            {
                if (xml.NodeType != XmlNodeType.Element) continue;

                if (xml.LocalName == "row")
                {
                    var rowAttr = xml.GetAttribute("r");
                    filaActual = rowAttr != null
                        ? int.Parse(rowAttr)
                        : filaActual + 1;

                    if (filaActual > TARGET_ROW) break;
                }
                else if (xml.LocalName == "c" && filaActual == TARGET_ROW)
                {
                    string cRef = xml.GetAttribute("r") ?? "";
                    enCeldaO = cRef.StartsWith(TARGET_COL_LETTER, StringComparison.OrdinalIgnoreCase)
                                   && cRef.Length >= 2;
                    tipoCelda = xml.GetAttribute("t") ?? "";
                }
                else if (xml.LocalName == "v" && enCeldaO)
                {
                    string raw = xml.ReadElementContentAsString();

                    if (tipoCelda == "s" && sharedStrings != null)
                    {
                        if (int.TryParse(raw, out int idx) && (uint)idx < (uint)sharedStrings.Length)
                            raw = sharedStrings[idx];
                    }

                    string dato = raw.Contains(" 00:00", StringComparison.Ordinal)
                        ? raw.Replace(" 00:00", "", StringComparison.Ordinal).Trim()
                        : raw.Trim();

                    return (dato, true);
                }
            }

            return (null, false);
        }

        private static string[]? CargarSharedStrings(ZipArchive zip)
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return null;

            var lista = new List<string>(512);

            using var stream = entry.Open();
            using var xml = XmlReader.Create(stream, XmlSettings);

            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "t")
                {
                    lista.Add(xml.ReadElementContentAsString());
                }
            }

            return lista.ToArray();
        }

        private static ZipArchiveEntry? BuscarPrimeraHoja(ZipArchive zip)
        {
            foreach (var entry in zip.Entries)
            {
                if (entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                    && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    return entry;
            }
            return null;
        }

        private static string ObtenerNombrePrimeraHoja(ZipArchive zip)
        {
            var wb = zip.GetEntry("xl/workbook.xml");
            if (wb == null) return "Hoja1";

            using var stream = wb.Open();
            using var xml = XmlReader.Create(stream, XmlSettings);

            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "sheet")
                    return xml.GetAttribute("name") ?? "Hoja1";
            }

            return "Hoja1";
        }

        private readonly struct ResultadoArchivo
        {
            public readonly string NombreArchivo;
            public readonly string NombreHoja;
            public readonly string? Valor;
            public readonly string? MensajeError;
            public readonly bool Exito;

            private ResultadoArchivo(
                string archivo, string hoja,
                string? valor, string? error, bool exito)
            {
                NombreArchivo = archivo;
                NombreHoja = hoja;
                Valor = valor;
                MensajeError = error;
                Exito = exito;
            }

            public static ResultadoArchivo Ok(string archivo, string hoja, string valor)
                => new(archivo, hoja, valor, null, true);

            public static ResultadoArchivo Fallo(string archivo, string mensaje)
                => new(archivo, "", null, mensaje, false);
        }
    }

    public class excel_data_reader_parallel_claude_mejora
    {
        public static void f_excel_data_reader_parallel_claude_mejora()
            => ExcelUltraFastReader.RunAsync().GetAwaiter().GetResult();
    }
}