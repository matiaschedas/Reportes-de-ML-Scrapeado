
/*
 hola
 */
using HtmlAgilityPack;
using System;
using System.Net.Http;
using System.Reflection.Metadata.Ecma335;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using static Main;
using OfficeOpenXml;
using System.Reflection.PortableExecutable;
using System.Text.Json.Serialization;
using OfficeOpenXml.Drawing.Style.Fill;
using System.Net.Http.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.ComponentModel.DataAnnotations.Schema;
using System.Globalization;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Chart;
using System.IO;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Reflection.Metadata;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Polling;
using Telegram.Bot.Types.Enums;
using System.Text;

public class Globals
{
    public string ClientSecret { get; set; } = "GWR6eZHrAZu47zusWIX8MinrMnSOjhYh";
    public string AppID { get; set; } = "2398820623563499";
    public string RefreshToken { get; set; } = "TG-67eb579fb8f51f0001e4b390-63251640";
    public string Token { get; set; } = "APP_USR-2398820623563499-033123-3d0e381174c25b915eb04dbd62d46505-63251640";

}

public class Program
{
    // Punto de entrada
    public static async Task Main(string[] args)
    {
        try
        {
            Globals globales = new Globals();
            string rutaDelDirectorio = AppContext.BaseDirectory;
            rutaDelDirectorio += "token.txt";
            var mainInstance = new Main(globales);
            var pathHistoricos = AppContext.BaseDirectory + "Historicos";
            var pathReportes = AppContext.BaseDirectory + "Reportes";
            mainInstance.CrearCarpeta(pathHistoricos);
            mainInstance.CrearCarpeta(pathReportes);
            

            bool ingresoCorrecto = false;
            string busquedaUser = "";
            string nombreReporte = "";
            string anioInicioStr = "";
            string anioFinStr = "";
            int anioInicio = 0;
            int anioFin = 0;
            rutaDelDirectorio = rutaDelDirectorio.Replace("token.txt", "");
            //rutaDelDirectorio += "busqueda.txt";
            List <string> rutas = mainInstance.GetFilesByPattern(rutaDelDirectorio);
            //Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(rutaDelDirectorio);
            string reporteComparador = "";
            bool EsReporteComparador = false;
            string fechaAComparar = "";
            
            string rutaComparador = AppContext.BaseDirectory;
            rutaComparador += "Comparador.txt";
            CMP comparador = mainInstance.obtenerCMP(rutaComparador);
            

            if (comparador.EliminarAnterior == true)
            {
                mainInstance.BorrarAnteriorEjecucionDeTodosLosHistoricos(pathHistoricos);
                mainInstance.BorrarReportesAnteriorEjecucion();
                Console.WriteLine("Anterior ejecucion eliminada con exito");
                Console.WriteLine("Presiona Enter para salir...");
                Console.ReadLine();
                return;
            }
            if (comparador.Comparador == true)
            {
                EsReporteComparador = true;
                fechaAComparar = comparador.FechaAComparar;
                if(fechaAComparar.ToLower() == "ayer")
                {
                    DateTime currentDateTime = DateTime.Now;
                    DateTime currentDate = currentDateTime.Date;
                    string fechaActual = currentDate.ToString("dd-MM-yyyy");
                    string fechaAyer = currentDate.AddDays(-1).ToString("dd-MM-yyyy");
                    fechaAComparar = fechaAyer;
                }
                if(fechaAComparar.ToLower() == "anterior")
                {
                    string rutaReportes = AppContext.BaseDirectory;
                    rutaReportes += "Reportes";
                    DateTime fechaMasReciente = mainInstance.FechaMasReciente(rutaReportes);
                    if (fechaMasReciente == new DateTime(2000,1,1))
                    {
                        Console.WriteLine("Error, no se pudo obtener la fecha mas reciente de los reportes anteriores");
                        Console.WriteLine("Presiona Enter para salir...");
                        Console.ReadLine();
                        return;
                    }
                    if(fechaMasReciente == DateTime.Today)
                    {
                        Console.WriteLine("Error, no se puede generar el reporte porque ya fue generado para el dia de hoy");
                        Console.WriteLine("Presiona Enter para salir...");
                        Console.ReadLine();
                        return;
                    }
                    fechaAComparar = fechaMasReciente.ToString("dd-MM-yyyy");
                    
                }
            }
            if (!EsReporteComparador)
            {
                if (!(rutas.Count > 0))
                {
                    while (!ingresoCorrecto)
                    {
                        Console.WriteLine("Por favor, ingresa el nombre del reporte a guardar:");
                        nombreReporte = Console.ReadLine() ?? System.String.Empty;
                        if (nombreReporte.Replace(" ", "") != "")
                        {
                            ingresoCorrecto = true;
                        }
                    }
                    ingresoCorrecto = false;
                    while (!ingresoCorrecto)
                    {
                        Console.WriteLine("Por favor, ingresa tu busqueda de ML:");
                        busquedaUser = Console.ReadLine() ?? System.String.Empty;
                        if (busquedaUser.Replace(" ", "") != "")
                        {
                            ingresoCorrecto = true;
                            busquedaUser = busquedaUser.Replace(" ", "%20");
                        }
                    }
                    ingresoCorrecto = false;
                    while (!ingresoCorrecto)
                    {
                        Console.WriteLine("Ingrese año de inicio");
                        anioInicioStr = Console.ReadLine() ?? System.String.Empty;
                        if (!int.TryParse(anioInicioStr, out int result))
                        {
                            ingresoCorrecto = false;
                        }
                        else
                        {
                            int.TryParse(anioInicioStr, out anioInicio);
                            ingresoCorrecto = true;
                        }
                    }
                    ingresoCorrecto = false;
                    while (!ingresoCorrecto)
                    {
                        Console.WriteLine("Ingrese año de fin");
                        anioFinStr = Console.ReadLine() ?? System.String.Empty;
                        if (!int.TryParse(anioFinStr, out int result))
                        {
                            ingresoCorrecto = false;
                        }
                        else
                        {
                            int.TryParse(anioFinStr, out anioFin);
                            if (anioFin >= anioInicio)
                            {
                                ingresoCorrecto = true;
                            }
                        }
                    }
                    await mainInstance.Procesamiento(rutaDelDirectorio, nombreReporte, anioInicio, anioFin, busquedaUser, new List<string>(), comparador.Cookie);
                }
                else
                {
                    var tareas = new List<Task>();
                    Console.WriteLine("Ejecutando Reporte del dia--------------------------------------------");
                    foreach (var archivo in rutas)
                    {
                        Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(archivo);
                        Task tarea = mainInstance.Procesamiento(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query, busquedaPorTxt.NoBuscar, comparador.Cookie);
                        tareas.Add(tarea);
                    }
                    await Task.WhenAll(tareas);
                }
            }
            else
            {
                if (!(rutas.Count > 0))
                {
                    Console.WriteLine("No se puede realizar el reporte porque no hay ningun archivo de busqueda");
                    Console.WriteLine("Presiona Enter para salir...");
                    Console.ReadLine();
                    return;
                }
                if (mainInstance.EsFechaValida(fechaAComparar)==false){
                    Console.WriteLine("No se puede realizar el reporte porque la fecha a comparar no es valida");
                    Console.WriteLine("Presiona Enter para salir...");
                    Console.ReadLine();
                    return;
                }
                var tareas = new List<Task>();
                Console.WriteLine("Ejecutando Reporte Comparador-----------------------------------------");
                foreach (var archivo in rutas)
                {
                    Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(archivo);
                    Task tarea =  mainInstance.ProcesamientoComparador(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query, fechaAComparar, busquedaPorTxt.NoBuscar, comparador.Cookie);
                    //Task tarea2 = mainInstance.Procesamiento(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query);
                    tareas.Add(tarea);
                    //tareas.Add(tarea2);
                }
                await Task.WhenAll(tareas);
                Console.WriteLine("Proceso Completado al: 50%");
                Console.WriteLine("Ejecutando Reporte del dia--------------------------------------------");
                tareas = new List<Task>();
                foreach (var archivo in rutas)
                {
                    Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(archivo);
                    Task tarea = mainInstance.Procesamiento(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query, busquedaPorTxt.NoBuscar, comparador.Cookie);
                    tareas.Add(tarea);
                }
                await Task.WhenAll(tareas);
            }

            List<Auto> autos = mainInstance.ObtenerAutosReportesHoy();
            mainInstance.EnviarMensajeTelegram(autos, comparador.NotificarPesos, comparador.NotificarUsd, comparador.IdTelegram);

            Console.WriteLine("Proceso Completado con Exito, Presiona Enter para salir...");
            Console.ReadLine();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }

    
}

public class Main
{
    private readonly Globals _globals;

    public Main(Globals globals)
    {
        _globals = globals;
    }
    SemaphoreSlim semaphore = new SemaphoreSlim(1, 1);
    bool ReemplazoTokenTerminado = false;
    bool esPrimerIteracion = true;
    SemaphoreSlim semaphoreProcesamiento = new SemaphoreSlim(1, 1);
    List<string> procesamientos = new List<string>();
    int porcentajeCompletado = 10;

    public void BorrarAnteriorEjecucionDelHistorico(string pathHistorico)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(pathHistorico))
        {
            // Recorre todas las hojas del archivo
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                // Encuentra la última fila con datos en la columna A
                int lastRow = worksheet.Dimension.End.Row;

                // Si la última fila tiene datos en la columna A, la elimina
                if (!string.IsNullOrEmpty(worksheet.Cells[lastRow, 1].Text))
                {
                    worksheet.DeleteRow(lastRow);
                }
            }
            // Guarda los cambios en el archivo Excel
            package.Save();
        }
    }

    public void BorrarAnteriorEjecucionDeTodosLosHistoricos(string pathHistoricos)
    {
        var files = Directory.GetFiles(pathHistoricos, "*.xlsx");
        foreach (var file in files)
        {
            BorrarAnteriorEjecucionDelHistorico(file);
        }
    }

    public void BorrarReportesAnteriorEjecucion()
    {
        string rutaDelDirectorio = AppContext.BaseDirectory;
        rutaDelDirectorio += "\\Reportes";
        DateTime fechaMasReciente = FechaMasReciente(rutaDelDirectorio);
        BorrarArchivosPorFecha(rutaDelDirectorio, fechaMasReciente);
    }

    public bool EsFechaValida(string fecha)
    {
        if (DateTime.TryParse(fecha, out DateTime fechaConvertida))
        {
            return fechaConvertida.Date < DateTime.Now.Date;
        }
        return false;
    }

    public void CrearCarpeta(string folderPath)
    {
        if (!Directory.Exists(folderPath))
        {
            Directory.CreateDirectory(folderPath);
        }
    }

    public DateTime FechaMasReciente(string ruta)
    {
        var archivos = Directory.GetFiles(ruta).Where(a => !Path.GetFileNameWithoutExtension(a).ToUpper().Contains("COMPARADOR")).ToList();
        DateTime fechaMasReciente = new DateTime(2000,1,1);
        foreach(var archivo in archivos)
        {
            string nombreArchivo = Path.GetFileNameWithoutExtension(archivo);
            Match match = Regex.Match(nombreArchivo, @"\b(\d{2}-\d{2}-\d{4})\b");
            if(match.Success && DateTime.TryParseExact(match.Value, "dd-MM-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fecha))
            {
                if (fecha > fechaMasReciente)
                {
                    fechaMasReciente = fecha;
                }
            }
        }
        return fechaMasReciente;
    }

    public void BorrarArchivosPorFecha(string path, DateTime fecha)
    {
        string datePrefix = fecha.ToString("dd-MM-yyyy");

        if (!Directory.Exists(path))
        {
            throw new DirectoryNotFoundException("El directorio no existe: " + path);
        }
        var files = Directory.GetFiles(path, "*.xlsx");

        // Recorre todos los archivos
        foreach (var file in files)
        {
            // Obtiene el nombre del archivo sin la ruta
            string fileName = Path.GetFileName(file);
            // Verifica si el nombre del archivo comienza con la fecha dada
            if (fileName.StartsWith(datePrefix))
            {
                try
                {
                    // Elimina el archivo
                    File.Delete(file);
                    Console.WriteLine($"Archivo eliminado: {fileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"No se pudo eliminar el archivo {fileName}: {ex.Message}");
                }
            }
        }
    }

    public async Task ProcesamientoComparador(string rutaDelDirectorio, string nombreReporte, int anioInicio, int anioFin, string busquedaUser, string fechaAComparar, List<string> noBuscar, string cookie)
    {
            DateTime currentDateTime = DateTime.Now;
            DateTime currentDate = currentDateTime.Date;
            string fechaActual = currentDate.ToString("dd-MM-yyyy");
            DateTime fechaAnterior = DateTime.Parse(fechaAComparar);
            string fechaAnteriorText = fechaAnterior.ToString("dd-MM-yyyy");
            string rutaDelDirectorioAnterior = rutaDelDirectorio + "Reportes\\" + fechaAnteriorText + " " + nombreReporte + ".xlsx";
            rutaDelDirectorio += "Reportes\\" + fechaActual + " " + nombreReporte + " COMPARADOR vs " + fechaAnteriorText + ".xlsx";

            if (!(File.Exists(rutaDelDirectorioAnterior)))
            {
                Console.WriteLine("No es posible hacer un reporte comparador porque no existe el reporte de la fecha: " + fechaAnteriorText);
                return;
            }
            int cantidadHojas = ObtenerCantidadHojas(rutaDelDirectorioAnterior);
            if (cantidadHojas == 0)
            {
                Console.WriteLine("No hay hojas en el reporte de ayer");
                return;
            }
            string primerHoja = ObtenerNombrePrimerHoja(rutaDelDirectorioAnterior);
            int hoja = int.Parse(primerHoja);
            for (int i = 0; i < cantidadHojas; i++)
            {
                Console.WriteLine("Realizando consulta...");
                string busqueda = busquedaUser + " " + hoja.ToString();
                busqueda = busqueda.Replace(" ", "-");
                busqueda += "_Desde_";
                List<Auto> results = new List<Auto>();
                results = await Query(busqueda, cookie);
                results = FiltrarNoBuscados(results, noBuscar);
                //results = BorrarCajaAutomatica(results);
                Console.WriteLine("Volcando resultados de " + busquedaUser + " " + hoja);

                List<List<string>> tablaHoja = new List<List<string>>();
                tablaHoja = LeerHojaReporteAnterior(rutaDelDirectorioAnterior, i);
                bool creado = CrearAbrirExcelReportes(rutaDelDirectorio, hoja.ToString(), busquedaUser.Replace("%20", " "));
                if (!creado)
                {
                    Console.WriteLine("Error al crear reporte");
                    return;
                }
                decimal oficialUSD = await ObtenerPrecioVentaDolarOficial();
                decimal blueUSD = await ObtenerPrecioVentaDolarBlue();
                results = FiltrarResultadosRepetidos(results, tablaHoja);
                List<Auto> cambiaronPrecio = ObtenerCambiaronPrecio(results, tablaHoja);
                CompletarReporte(rutaDelDirectorio, results, hoja.ToString(), oficialUSD, blueUSD, hoja);
                ModificarReporte(rutaDelDirectorioAnterior, cambiaronPrecio, hoja.ToString());
                ModificarReporte(rutaDelDirectorio, cambiaronPrecio, hoja.ToString());
                hoja++;
            }
       
    }

    public List<Auto> FiltrarNoBuscados(List<Auto> autos, List<string> noBuscar)
    {
            List<Auto> autosEliminados = autos
                .Where(auto => noBuscar.Any(palabra => auto.Descripcion.Contains(palabra, StringComparison.OrdinalIgnoreCase)))
                .ToList();
            autos.RemoveAll(auto => noBuscar.Any(palabra => auto.Descripcion.Contains(palabra, StringComparison.OrdinalIgnoreCase)));
            return autos;
    }

    public void ModificarReporte(string ruta, List<Auto> cambiaronPrecio, string hoja)
    {
        if (cambiaronPrecio.Count == 0)
        {
            return;
        }
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var paquete = new ExcelPackage(new FileInfo(ruta)))
        {

            var workbook = paquete.Workbook;
            var ws = workbook.Worksheets[hoja];
            int rowCount = ws.Dimension?.Rows ?? 0;
            // Recorremos todas las filas y limpiamos el color de fondo
            for (int row = 1; row <= rowCount; row++)
            {
                var cell = ws.Cells[row, 1];
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None; // Sin relleno
            }

            foreach(var fila in cambiaronPrecio)
            {
                for (int row = 1; row <= rowCount; row++)
                {
                    var cell = ws.Cells[row, 1];
                    if (cell.Text.Replace("-", "") == fila.ID.Replace("-", ""))
                    {
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFF00"));
                        break;
                    }
                }
            }
            paquete.Save();
        }
    }

    public List<Auto> FiltrarResultadosRepetidos(List<Auto> results, List<List<string>> tabla)
    {
        //esta funcion se queda con todos los autos que tengan o un ID nuevo (respecto al reporte anterior es decir tabla) o un precio nuevo
        var idsYPrecios = new HashSet<(string, decimal)>();
        foreach (var fila in tabla)
        {
            decimal precio = decimal.Parse(fila[3].ToString());
            idsYPrecios.Add((fila[0], precio));
        }
        results.RemoveAll(r => idsYPrecios.Contains((r.ID, r.Precio ?? decimal.Zero)));
        return results;
    }
    public List<Auto> ObtenerCambiaronPrecio(List<Auto> results, List<List<string>> tabla)
    {
        //esta funcion se queda con los autos que tengan un ID que ya existia en el reporte anterior es decir en tabla pero que ahora tiene otro precio
        var idsYPrecios = new HashSet<(string, decimal)>();
        var ids = new HashSet<string>();
        foreach(var fila in tabla)
        {
            decimal precio = decimal.Parse(fila[3].ToString());
            idsYPrecios.Add((fila[0], precio));
            ids.Add(fila[0]);
        }
        return results.Where(r => ids.Contains(r.ID) && !idsYPrecios.Contains((r.ID, r.Precio ?? decimal.Zero))).ToList();
    }

    public string ObtenerNombrePrimerHoja(string ruta)
    {
        string primerHoja = "";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var paquete = new ExcelPackage(new FileInfo(ruta)))
        {
            var workbook = paquete.Workbook;
            primerHoja = workbook.Worksheets[0].Name;
        }
        return primerHoja;
    }

    public int ObtenerCantidadHojas(string ruta)
    {
        var cantidad = 0;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var paquete = new ExcelPackage(new FileInfo(ruta)))
        {
            var workbook = paquete.Workbook;
            cantidad = workbook.Worksheets.Count;
        }
        return cantidad;
    }

    public List<List<string>> LeerHojaReporteAnterior(string ruta, int hoja)
    {
        var tabla = new List<List<string>>();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var paquete = new ExcelPackage(new FileInfo(ruta)))
        {
            var ws = paquete.Workbook.Worksheets[hoja];
            int filasTotales = ws.Dimension.End.Row;
            for(int fila = 3; fila<=filasTotales; fila++)
            {
                var datosFila = new List<string>();
                for(int columna = 1; columna <= 4; columna++)
                {
                    var celda = ws.Cells[fila, columna].Text;
                    datosFila.Add(celda);
                }
                tabla.Add(datosFila);
            }
        }
        tabla = tabla.Where(fila => fila.Any(celda => !string.IsNullOrWhiteSpace(celda))).ToList();
        return tabla;
    }

    public List<Auto> ObtenerAutosReportesHoy()
    {
        List<Auto> autos = new List<Auto>();
        string rutaDelDirectorio = AppContext.BaseDirectory;
        rutaDelDirectorio += "Reportes\\";
        DateTime currentDateTime = DateTime.Now;
        DateTime currentDate = currentDateTime.Date;
        string fechaActual = currentDate.ToString("dd-MM-yyyy");
        string patron = fechaActual + "*.xlsx";
        string [] archivos = Directory.GetFiles(rutaDelDirectorio, patron);
        var archivosFiltrados = archivos
           .Where(archivo =>
               !Path.GetFileName(archivo)
                    .Contains("COMPARADOR", StringComparison.OrdinalIgnoreCase))
           .ToList();
        foreach (var archivo in archivosFiltrados)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var paquete = new ExcelPackage(new FileInfo(archivo)))
            {
                var workbook = paquete.Workbook;
                foreach (var worksheet in workbook.Worksheets)
                {
                    int filasTotales = worksheet.Dimension.End.Row;
                    if (filasTotales != 3)
                    {
                        for (int fila = 3; fila <= filasTotales; fila++)
                        {
                            var auto = new Auto();
                            auto.ID = worksheet.Cells[fila, 1].Text;
                            auto.Descripcion = worksheet.Cells[fila, 2].Text;
                            auto.Kilometros = worksheet.Cells[fila, 3].Text;
                            auto.Precio = decimal.Parse(worksheet.Cells[fila, 4].Text);
                            auto.Moneda = worksheet.Cells[fila, 5].Text;
                            autos.Add(auto);
                        }
                    }
                }
            }
        }
        return autos;
    }

    public void EnviarMensajeTelegram(List<Auto> autos, int notificarPesos, int notificarUsd, int idTelegram)
    {
        List<Auto> autosFiltrados = autos.Where(auto => auto.Moneda == "$" && auto.Precio < notificarPesos || auto.Moneda == "US$" && auto.Precio < notificarUsd).OrderBy(auto => auto.Precio).ToList();
        if (autosFiltrados.Count == 0)
        {
            Console.WriteLine("No hay autos para enviar por Telegram");
            return;
        }

        ITelegramBotClient _botClient;
        _botClient = new TelegramBotClient("6218585144:AAGHlD78x8FCPdsZbxmgEEWOJEnKwIIkaBw");

        var me = _botClient.GetMeAsync().Result;
        Console.WriteLine($"Hi, I am {me.Id} and my name is:  {me.FirstName}");

        var receiverOptions = new ReceiverOptions
        {
            AllowedUpdates = new UpdateType[]
            {
                    UpdateType.Message,
                    UpdateType.EditedMessage,
            }
        };
        int destino = idTelegram;
        string mensaje = "";
        var sb = new StringBuilder();
        
        sb.AppendLine("ID             | KM       | Precio     ");
        foreach (var auto in autosFiltrados)
        {
            string precioFormateado = auto.Precio.HasValue
                                                            ? auto.Precio.Value.ToString("N0")
                                                            : "N/A";
            sb.AppendLine($"{auto.ID.PadRight(13)}| {auto.Kilometros.PadRight(9)}| {precioFormateado.PadLeft(10)}");
        }
        
        string mensajeCompleto = sb.ToString();

        // Cortar en bloques de hasta 4096 caracteres
        List<string> partes = DividirEnPartes(mensajeCompleto, 4000); // dejamos margen por seguridad

        foreach (string parte in partes)
        {
            _botClient.SendTextMessageAsync(destino, $"<pre>{System.Net.WebUtility.HtmlEncode(parte)}</pre>", Telegram.Bot.Types.Enums.ParseMode.Html).Wait();
        }
    }

    public List<string> DividirEnPartes(string texto, int maxLongitud)
    {
        var partes = new List<string>();
        int inicio = 0;

        while (inicio < texto.Length)
        {
            int largo = Math.Min(maxLongitud, texto.Length - inicio);
            partes.Add(texto.Substring(inicio, largo));
            inicio += largo;
        }

        return partes;
    }

    public async Task Procesamiento(string rutaDelDirectorio, string nombreReporte, int anioInicio, int anioFin, string busquedaUser, List<string> noBuscar, string cookie)
    {
        DateTime currentDateTime = DateTime.Now;
        DateTime currentDate = currentDateTime.Date;
        string fechaActual = currentDate.ToString("dd-MM-yyyy");
        rutaDelDirectorio += "Reportes\\" + fechaActual + " " + nombreReporte + ".xlsx";
        HashSet<(string anio, decimal promedio)> tuplaHistorico = new HashSet<(string anio, decimal promedio)>();

        for (; anioInicio <= anioFin; anioInicio++)
        {
            Console.WriteLine("Realizando consulta...");
            string busqueda = busquedaUser + " " + anioInicio.ToString();
            busqueda = busqueda.Replace(" ", "-");
            busqueda += "_Desde_";
            
            List<Auto> results = await Query(busqueda, cookie);
            results = FiltrarNoBuscados(results, noBuscar);
            //bool existe = verificarIdExiste(results, "MLA1478269101");
            //results = BorrarCajaAutomatica(results);
            //existe = verificarIdExiste(results, "MLA1478269101");
            //imprimirResultados(results);
            Console.WriteLine("Volcando resultados de " + busquedaUser + " " + anioInicio);

            bool creado = CrearAbrirExcelReportes(rutaDelDirectorio, anioInicio.ToString(), busquedaUser.Replace("%20", " "));
            if (!creado)
            {
                Console.WriteLine("Error al crear reporte");
            }
            else
            {
                decimal oficialUSD = await ObtenerPrecioVentaDolarOficial();
                decimal blueUSD = await ObtenerPrecioVentaDolarBlue();
                (string anio, decimal promedio) = CompletarReporte(rutaDelDirectorio, results, anioInicio.ToString(), oficialUSD, blueUSD, anioInicio);
                tuplaHistorico.Add((anio, promedio));
                //Console.WriteLine($"Terminados los resultados del {anioInicio}");
            }
        }
        string rutaHistorico = AppContext.BaseDirectory;
        rutaHistorico += "Historicos\\" + nombreReporte + " HISTORICO.xlsx";
        CompletarHistorico(rutaHistorico, tuplaHistorico);

    }

    public async Task<List<Auto>> Query(string busquedaUser, string cookie)
    {
        int offset = 0;
        string urlScrapear = "https://autos.mercadolibre.com.ar/" + busquedaUser;
        urlScrapear += offset.ToString();
        List<Auto> results = new List<Auto>();
        HtmlWeb web = new HtmlWeb();
        int LimitLoop = 50;
        int iteration = 0;

        
       
        

        while (iteration < LimitLoop)
        {
            //HtmlDocument doc = web.Load(urlScrapear);
            using var handler = new HttpClientHandler();
            using var client = new HttpClient(handler);
            var request = new HttpRequestMessage(HttpMethod.Get, urlScrapear);
            request.Headers.Add("Cookie", cookie);
            request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/114.0");
            var response = await client.SendAsync(request);
            string html = await response.Content.ReadAsStringAsync();
            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);

            var cartelLogin = doc.DocumentNode.SelectSingleNode("//p[contains(@class, 'message-card')]");
            if( cartelLogin != null)
            {
                throw new Exception("Logear ML, cambiar la cookie");
            }

            var cartelSinPublicaciones = doc.DocumentNode.SelectNodes("//div[contains(@class, 'ui-search-rescue')]");
            if (cartelSinPublicaciones != null)
            {
                return results;
            }
            var items = doc.DocumentNode.SelectNodes("//li[contains(@class, 'ui-search-layout__item')]");
            if (items == null) return results;
            List<Auto> autos = new List<Auto>();

            foreach (var item in items)
            {
                var auto = new Auto();
                var descripcionNode = item.SelectSingleNode(".//a");
                if (descripcionNode != null)
                {
                    auto.Descripcion = descripcionNode.InnerText.Trim();
                    string href = descripcionNode.GetAttributeValue("href", string.Empty).Trim();
                    //auto.ID
                    Regex rg = new Regex(@"(MLA-\d+)");
                    Match match = rg.Match(href);
                    if (match.Success)
                    {
                        auto.ID = match.Value;
                    }
                    else
                    {
                        var url = href;
                        using var handler2 = new HttpClientHandler
                        {
                            AllowAutoRedirect = true
                        };
                        using var client2 = new HttpClient(handler2);
                        var response2 = await client2.GetAsync(url);
                        string finalURL = response2.RequestMessage.RequestUri.ToString();
                        
                        Regex rg2 = new Regex(@"(MLA-\d+)");
                        Match match2 = rg2.Match(finalURL);
                            if (match2.Success)
                            {
                                auto.ID = match2.Value;
                            }
                            else
                            {
                                auto.ID = "Error";
                            }
                    }

                }
                var kilometrosNode = item.SelectSingleNode("(.//li[contains(@class, 'poly-attributes_list__item') and contains(@class, 'poly-attributes_list__separator')])[2]");
                if (kilometrosNode != null)
                {
                    auto.Kilometros = kilometrosNode.InnerText.Trim();
                }
                var precioNode = item.SelectSingleNode(".//span[contains(@class, 'andes-money-amount__fraction')]");
                if (precioNode != null)
                {
                    decimal precioDecimal;
                    if (decimal.TryParse(precioNode.InnerText.Trim(), out precioDecimal))
                    {
                        auto.Precio = precioDecimal;
                    }
                    else
                    {
                        // La conversión falló, manejar el error o asignar un valor predeterminado
                        auto.Precio = 0;
                    }
                }
                var monedaNode = item.SelectSingleNode(".//span[contains(@class, 'andes-money-amount__currency-symbol')]");
                if (monedaNode != null)
                {
                    auto.Moneda = monedaNode.InnerText.Trim();
                }
                autos.Add(auto);
            }
            results.AddRange(autos);
            if (offset == 0)
            {
                offset += 1;
            }
            offset += 48;
            urlScrapear = "https://autos.mercadolibre.com.ar/" + busquedaUser;
            urlScrapear += offset.ToString();
            iteration++;
        }
        return results;
    }

    public bool verificarIdExiste(List<Result> results, string ID)
    {
        foreach (Result result in results)
        {
            if (ID == result.Id)
                return true;
        }
        return false;
    }

    public void imprimirResultados(List<Result> results)
    {
        foreach (var result in results)
        {
            string anio = GetAttributeValue(result, "Año");
            string kms = GetAttributeValue(result, "Kilómetros");
            string transmision = GetAttributeValue(result, "Transmisión");
            Console.WriteLine($"ID: {result.Id}, Título: {result.Title}, Precio: {result.Price}, Moneda: {result.Currency_id}, Año: {anio}, Kilómetros: {kms}, Transmisión: {transmision}");
        }
        Console.WriteLine($"Cantidad de productos: {results.LongCount()}");
    }
    

    public static void ReemplazarContenidoArchivo(string rutaArchivo, string nuevoContenido)
    {
        try
        {
            // Reemplaza el contenido del archivo con el nuevo string
            File.WriteAllText(rutaArchivo, nuevoContenido);
            Console.WriteLine($"Contenido del archivo '{rutaArchivo}' reemplazado exitosamente.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al reemplazar el contenido del archivo: {ex.Message}");
        }
    }

    public static string GetAttributeValue(Result product, string attributeName)
    {
        foreach (var attribute in product.Attributes)
        {
            if (attribute.Name == attributeName)
            {
                return attribute.Value_name;
            }
        }
        return null; 
    }
    public List<Result> BorrarCajaAutomatica(List<Result> results)
    {
        List<Result> resultsFiltrado = results.ToList();
        resultsFiltrado.RemoveAll(item => GetAttributeValue(item, "Transmisión") == "Automática");
        //resultsFiltrado.RemoveAll(item => GetAttributeValue(item, "Año") != anioInicio.ToString());
        return resultsFiltrado;
    }

    public List<Result> BorrarCajaManual(List<Result> results) {
        List<Result> resultsFiltrado = results.ToList();
        resultsFiltrado.RemoveAll(item => GetAttributeValue(item, "Transmisión") == "Manual");
        //resultsFiltrado.RemoveAll(item => GetAttributeValue(item, "Año") != anioInicio.ToString());
        return resultsFiltrado;
    }

    public void CompletarHistorico(string ruta, HashSet<(string anio, decimal promedio)> tuplaHistorico)
    {
        if (!File.Exists(ruta))
        {
            string rutaTemplate = Path.GetDirectoryName(ruta).Replace("\\Historicos", "") + "\\Template HISTORICO.xlsx";
            File.Copy(rutaTemplate, ruta);
            Console.WriteLine("Se creo el historico: " + ruta);
        }
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(ruta)))
        {
            var workbook = package.Workbook;
            foreach (var tupla in tuplaHistorico)
            {
                var worksheet = workbook.Worksheets[tupla.anio];
                if (worksheet == null)
                    return;
                var tabla = worksheet.Tables.FirstOrDefault();
                if (tabla == null)
                    return;
                DateTime currentDateTime = DateTime.Now;
                DateTime currentDate = currentDateTime.Date;
                string fechaActual = currentDate.ToString("dd/MM/yyyy");
                
                int ultimaFilaTabla = tabla.Address.End.Row;
                while (ultimaFilaTabla > tabla.Address.Start.Row && string.IsNullOrEmpty(worksheet.Cells[ultimaFilaTabla, 1].Text))
                {
                    ultimaFilaTabla--;
                }
                ultimaFilaTabla++;
                worksheet.Cells[ultimaFilaTabla, 1].Value = fechaActual;
                worksheet.Cells[ultimaFilaTabla, 2].Value = tupla.promedio;
                

                string nuevaCeldaFin = worksheet.Cells[ultimaFilaTabla, tabla.Address.End.Column].Address;
                tabla.TableXml.InnerXml = tabla.TableXml.InnerXml.Replace(tabla.Address.End.Address, nuevaCeldaFin);
                var chart = worksheet.Drawings.OfType<ExcelChart>().FirstOrDefault();
                if (chart != null)
                {
                    string nuevaDireccionX = worksheet.Cells[tabla.Address.Start.Row + 1, tabla.Address.Start.Column, ultimaFilaTabla, tabla.Address.Start.Column].Address;
                    string nuevaDireccionY = worksheet.Cells[tabla.Address.Start.Row + 1, tabla.Address.Start.Column + 1, ultimaFilaTabla, tabla.Address.Start.Column + 1].Address;

                    chart.Series[0].XSeries = $"'{worksheet.Name}'!{nuevaDireccionX}";
                    chart.Series[0].Series = $"'{worksheet.Name}'!{nuevaDireccionY}";
                }
            }
            package.Save();
        }
    }
    public (string anio, decimal promedio) CompletarReporte(string ruta, List<Auto> results, string hoja, decimal dolarOficial, decimal dolarBlue, int anioInicio)
    {
        try
        {
            var itemsOrdenados = results.OrderBy(item => item.Precio).ToList();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(ruta)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[hoja]; // Selecciona la hoja
                int fila = 3;
                decimal cantUSD = 0;
                decimal cantARS = 0;
                decimal totalUSD = 0;
                decimal totalARS = 0;
                decimal totalUSDBlue = 0;
                decimal totalUSDOficial = 0;
                foreach (var result in itemsOrdenados)
                {
                    string anio = hoja;
                    string kms = result.Kilometros;
                    decimal precio = result.Precio ?? decimal.Zero;
                    string id = result.ID;
                    string moneda = result.Moneda;
                    if (moneda == "US$")
                    {
                        cantUSD++;
                        totalUSD += precio;
                    }
                    else
                    {
                        cantARS++;
                        totalARS += precio;
                        totalUSDBlue += precio / dolarBlue;
                        totalUSDOficial += precio / dolarOficial;
                    }
                    string descripcion = result.Descripcion;
                    worksheet.Cells[fila, 1].Value = id;
                    worksheet.Cells[fila, 2].Value = descripcion;
                    worksheet.Cells[fila, 3].Value = kms;
                    worksheet.Cells[fila, 4].Value = precio;
                    worksheet.Cells[fila, 5].Value = moneda;
                    worksheet.Cells[fila, 6].Value = anio;
                    fila += 1;
                }


                decimal promUSD, promARS, promUSDBlueARS, promUSDOficialARS;
                promUSD = cantUSD > 0 ? (totalUSD / cantUSD) : 0;
                promARS = cantARS > 0 ? (totalARS / cantARS) : 0;
                promUSDBlueARS = cantARS > 0 ? (totalUSDBlue / cantARS) : 0;
                promUSDOficialARS = cantARS > 0 ? (totalUSDOficial / cantARS) : 0;
 
                promUSDBlueARS += promUSD;
                promUSDBlueARS = promUSD>0 ? promUSDBlueARS/2:promUSDBlueARS;
                promUSDOficialARS += promUSD;
                promUSDOficialARS = promUSD>0 ? promUSDOficialARS / 2: promUSDOficialARS;


                worksheet.Cells[3, 8].Value = promUSD;
                worksheet.Cells[3, 9].Value = promARS;
                worksheet.Cells[3, 10].Value = promUSDBlueARS;
                worksheet.Cells[3, 11].Value = promUSDOficialARS;
                worksheet.Cells[3, 12].Value = dolarOficial;
                worksheet.Cells[3, 13].Value = dolarBlue;
                

                package.Save();
                return (hoja, promUSDBlueARS);

            }
        }
        catch (Exception ex)
        {
            return (("",0));
        }
    }
    public bool CrearAbrirExcelReportes(string ruta, string hoja, string query)
    {
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<string> headers = new List<string> { "ID", "Descripcion", "Kilometros", "Precio", "Moneda", "Año" };
            List<string> headers2 = new List<string> { "Promedio publicados en USD", "Promedio publicados en ARS", "Promedio todos en USD Blue", "Promedio todos en USD Oficial", "Cotizacion USD Oficial", "Cotizacion USD Blue" };           
            if (File.Exists(ruta))
            {
                // Abrir el archivo existente
                using (var package = new ExcelPackage(new FileInfo(ruta)))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[hoja]; // Selecciona la hoja
                    if (worksheet == null)
                    {
                        worksheet = package.Workbook.Worksheets.Add(hoja); //si no existia la crea
                    }
                    else
                    {
                        Console.WriteLine("El Reporte ya estaba creado para el dia de hoy");
                        return false;
                    }
                    worksheet.Cells[1, 1].Value = query;
                    int i;
                    for (i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cells[2, i+1].Value = headers[i];
                    }
                    int indiceHeader2 = 0;
                    for (i = headers.Count; i < headers2.Count+headers.Count; i++)
                    {
                        worksheet.Cells[2, i + 2].Value = headers2[indiceHeader2];
                        indiceHeader2 += 1;
                    }
                    // Aquí puedes leer o modificar los datos, por ejemplo:
                    //Console.WriteLine("Valor de la celda A1: " + worksheet.Cells[1, 1].Text);
                    // Guardar cambios si es necesario
                    package.Save();
                    return true;
                }
            }
            else
            {
                // Crear un nuevo archivo
                using (var package = new ExcelPackage(new FileInfo(ruta)))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets.Add(hoja); // Crear nueva hoja
                    int i;
                    worksheet.Cells[1, 1].Value = query;
                    for (i = 0; i < headers.Count; i++)
                    {
                        worksheet.Cells[2, i + 1].Value = headers[i];
                    }
                    int indiceHeader2 = 0;
                    for (i = headers.Count; i < headers2.Count + headers.Count; i++)
                    {
                        worksheet.Cells[2, i + 2].Value = headers2[indiceHeader2];
                        indiceHeader2 += 1;
                    }
                    // Guardar el archivo
                    package.Save();
                    return true;
                }
            }
        }
        catch
        {
            return false;
        }
    }
    
    public async Task<decimal> ObtenerPrecioVentaDolarOficial()
    {
        
        try
        {
            HttpClient client = new HttpClient();
            // Realiza la solicitud GET a la API
            HttpResponseMessage response = await client.GetAsync("https://dolarapi.com/v1/dolares/oficial");
            response.EnsureSuccessStatusCode();

            // Lee la respuesta JSON
            string responseBody = await response.Content.ReadAsStringAsync();

            // Deserializa el JSON a un objeto C#
            var dolarData = JsonSerializer.Deserialize<DolarApiResponse>(responseBody);

            // Devuelve el precio de venta
            return dolarData.Venta;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error al obtener el precio del dólar: " + ex.Message);
            return 0;
        }
    }

    public async Task<decimal> ObtenerPrecioVentaDolarBlue()
    {

        try
        {
            HttpClient client = new HttpClient();
            // Realiza la solicitud GET a la API
            HttpResponseMessage response = await client.GetAsync("https://dolarapi.com/v1/dolares/blue");
            response.EnsureSuccessStatusCode();

            // Lee la respuesta JSON
            string responseBody = await response.Content.ReadAsStringAsync();

            // Deserializa el JSON a un objeto C#
            var dolarData = JsonSerializer.Deserialize<DolarApiResponse>(responseBody);

            // Devuelve el precio de venta
            return dolarData.Venta;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error al obtener el precio del dólar: " + ex.Message);
            return 0;
        }
    }

    public Busqueda obtenerBusqueda(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            string jsonContent = File.ReadAllText(filePath);
            if (string.IsNullOrWhiteSpace(jsonContent))
            {
                return null;
            }
            Busqueda reporte = JsonSerializer.Deserialize<Busqueda>(jsonContent);
            return reporte;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al procesar el archivo de busqueda: {ex.Message}");
            return null;
        }
    }

    public CMP obtenerCMP(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            string jsonContent = File.ReadAllText(filePath);
            if (string.IsNullOrWhiteSpace(jsonContent))
            {
                return null;
            }
            CMP reporte = JsonSerializer.Deserialize<CMP>(jsonContent);
            return reporte;
        }
        catch(Exception ex)
        {
            Console.WriteLine($"Error al procesar el archivo comparador: {ex.Message}");
            return null;
        }
    }

    public List<string> GetFilesByPattern(string folderPath)
    {
        List<string> matchingFiles = new List<string>();

        try
        {
            // Obtener todos los archivos de la carpeta
            string[] allFiles = Directory.GetFiles(folderPath);

            // Expresión regular para buscar nombres como busqueda1.txt, busqueda2.txt, etc.
            Regex regex = new Regex(@"busqueda\d+\.txt$", RegexOptions.IgnoreCase);

            // Filtrar los archivos que coincidan con el patrón
            foreach (string file in allFiles)
            {
                if (regex.IsMatch(Path.GetFileName(file)))
                {
                    matchingFiles.Add(file);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al buscar archivos de busqueda: {ex.Message}");
        }

        return matchingFiles;
    }

    public class DolarApiResponse
    {
        [JsonPropertyName("compra")]
        public decimal Compra { get; set; }

        [JsonPropertyName("venta")]
        public decimal Venta { get; set; }
    }
    public class Attribute
    {
        public string Name { get; set; } = string.Empty;
        public string Value_name { get; set; } = string.Empty;
    }
    public class Result
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Condition { get; set; } = string.Empty;
        public string Thumbnail { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public string Currency_id { get; set; } = string.Empty;
        public Attribute[] Attributes { get; set; }
    }
    public class Busqueda
    {
        [JsonPropertyName("nombre_reporte")]
        public string NombreReporte { get; set; }

        [JsonPropertyName("busqueda")]
        public string query { get; set; }

        [JsonPropertyName("anio_inicio")]
        public int AnioInicio { get; set; }

        [JsonPropertyName("anio_fin")]
        public int AnioFin { get; set; }
        [JsonPropertyName("no_buscar")]
        public List<string> NoBuscar { get; set; }

    }
    public class CMP
    {
        [JsonPropertyName("comparador")]
        public bool Comparador { get; set; }
        [JsonPropertyName("fecha_a_comparar")]
        public string FechaAComparar { get; set; }
        [JsonPropertyName("eliminar_anterior")]
        public bool EliminarAnterior { get; set; }
        [JsonPropertyName("notificar_pesos")]
        public int NotificarPesos { get; set; }
        [JsonPropertyName("notificar_usd")]
        public int NotificarUsd { get; set; }//por defecto no notifica pesos
        [JsonPropertyName("id_telegram")]
        public int IdTelegram { get; set; } //por defecto no notifica pesos
        [JsonPropertyName("cookie")]
        public string Cookie { get; set; }
    }
    public class Auto
    {
        public string? Descripcion { get; set; }
        public string? ID { get; set; }
        public string? Kilometros { get; set; }
        public decimal? Precio { get; set; }
        public string? Moneda { get; set; }
    }
}





