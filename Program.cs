// See https://aka.ms/new-console-template for more information

//Client Secret: GWR6eZHrAZu47zusWIX8MinrMnSOjhYh
//App ID: 2398820623563499
//Redirect URI: https://www.matchkraft.com/
//code: https://matchkraft.com/?code=TG-6794214500e5ab0001ebfa0d-63251640
//codigo postman: 
/*
 * 
curl -X POST \
-H 'accept: application/json' \
-H 'content-type: application/x-www-form-urlencoded' \
'https://api.mercadolibre.com/oauth/token' \
-d 'grant_type=authorization_code' \
-d 'client_id=2398820623563499' \
-d 'client_secret=GWR6eZHrAZu47zusWIX8MinrMnSOjhYh' \
-d 'code=TG-6794214500e5ab0001ebfa0d-63251640' \
-d 'redirect_uri=https://www.matchkraft.com/' \
-d 'code_verifier=$CODE_VERIFIER'

Respuesta:

{
    "access_token": "APP_USR-2398820623563499-012614-5826e66cbd8134ea750553cb6c8571bb-63251640",
    "token_type": "Bearer",
    "expires_in": 21600,
    "scope": "offline_access read write",
    "user_id": 63251640,
    "refresh_token": "TG-6794216100e5ab0001ebfb44-63251640"
}

uso de api:
curl -X GET -H 'Authorization: Bearer APP_USR-2398820623563499-012614-5826e66cbd8134ea750553cb6c8571bb-63251640' https://api.mercadolibre.com/sites/MLA/search?q=Motorola%20G6

Refresh token:

curl -X POST \
-H 'accept: application/json' \
-H 'content-type: application/x-www-form-urlencoded' \
'https://api.mercadolibre.com/oauth/token' \
-d 'grant_type=refresh_token' \
-d 'client_id=2398820623563499' \
-d 'client_secret=GWR6eZHrAZu47zusWIX8MinrMnSOjhYh' \
-d 'refresh_token=TG-6794216100e5ab0001ebfb44-63251640'

 */

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


public class Globals
{
    public string ClientSecret { get; set; } = "GWR6eZHrAZu47zusWIX8MinrMnSOjhYh";
    public string AppID { get; set; } = "2398820623563499";
    public string RefreshToken { get; set; } = "TG-6794216100e5ab0001ebfb44-63251640";
    public string Token { get; set; } = "APP_USR-2398820623563499-012614-5826e66cbd8134ea750553cb6c8571bb-63251640";

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
            string token = mainInstance.LeeerTokenDesdeArchivo(rutaDelDirectorio);
            globales.Token = token;
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
                        return;
                    }
                    if(fechaMasReciente == DateTime.Today)
                    {
                        Console.WriteLine("Error, no se puede generar el reporte porque ya fue generado para el dia de hoy");
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
                        nombreReporte = Console.ReadLine() ?? String.Empty;
                        if (nombreReporte.Replace(" ", "") != "")
                        {
                            ingresoCorrecto = true;
                        }
                    }
                    ingresoCorrecto = false;
                    while (!ingresoCorrecto)
                    {
                        Console.WriteLine("Por favor, ingresa tu busqueda de ML:");
                        busquedaUser = Console.ReadLine() ?? String.Empty;
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
                        anioInicioStr = Console.ReadLine() ?? String.Empty;
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
                        anioFinStr = Console.ReadLine() ?? String.Empty;
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
                    await mainInstance.Procesamiento(rutaDelDirectorio, nombreReporte, anioInicio, anioFin, busquedaUser);
                }
                else
                {
                    var tareas = new List<Task>();
                    Console.WriteLine("Ejecutando Reporte del dia--------------------------------------------");
                    foreach (var archivo in rutas)
                    {
                        Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(archivo);
                        Task tarea = mainInstance.Procesamiento(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query);
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
                    return;
                }
                if (mainInstance.EsFechaValida(fechaAComparar)==false){
                    Console.WriteLine("No se puede realizar el reporte porque la fecha a comparar no es valida");
                    return;
                }
                var tareas = new List<Task>();
                Console.WriteLine("Ejecutando Reporte Comparador-----------------------------------------");
                foreach (var archivo in rutas)
                {
                    Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(archivo);
                    Task tarea =  mainInstance.ProcesamientoComparador(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query, fechaAComparar);
                    //Task tarea2 = mainInstance.Procesamiento(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query);
                    tareas.Add(tarea);
                    //tareas.Add(tarea2);
                }
                await Task.WhenAll(tareas);
                Console.WriteLine("Ejecutando Reporte del dia--------------------------------------------");
                tareas = new List<Task>();
                foreach (var archivo in rutas)
                {
                    Busqueda busquedaPorTxt = mainInstance.obtenerBusqueda(archivo);
                    Task tarea = mainInstance.Procesamiento(rutaDelDirectorio, busquedaPorTxt.NombreReporte, busquedaPorTxt.AnioInicio, busquedaPorTxt.AnioFin, busquedaPorTxt.query);
                    tareas.Add(tarea);
                }
                await Task.WhenAll(tareas);
            }
            
           
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

    public bool EsFechaValida(string fecha)
    {
        if (DateTime.TryParse(fecha, out DateTime fechaConvertida))
        {
            return fechaConvertida.Date < DateTime.Now.Date;
        }
        return false;
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

    public async Task ProcesamientoComparador(string rutaDelDirectorio, string nombreReporte, int anioInicio, int anioFin, string busquedaUser, string fechaAComparar)
    {
        DateTime currentDateTime = DateTime.Now;
        DateTime currentDate = currentDateTime.Date;
        string fechaActual = currentDate.ToString("dd-MM-yyyy");
        DateTime fechaAnterior = DateTime.Parse(fechaAComparar);
        string fechaAnteriorText = fechaAnterior.ToString("dd-MM-yyyy");
        string rutaDelDirectorioAnterior = rutaDelDirectorio +  "Reportes\\" + fechaAnteriorText + " " + nombreReporte + ".xlsx";
        rutaDelDirectorio += "Reportes\\" + fechaActual + " " + nombreReporte + " COMPARADOR vs " + fechaAnteriorText + ".xlsx";

        if (!(File.Exists(rutaDelDirectorioAnterior)))
        {
            Console.WriteLine("No es posible hacer un reporte comparador porque no existe el reporte de la fecha: "+ fechaAnteriorText);
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
        for(int i = 0; i<cantidadHojas; i++)
        {
            Console.WriteLine("Realizando consulta...");
            string busqueda = busquedaUser + " " + hoja.ToString();
            busqueda = busqueda.Replace(" ", "%20");
            List<Result> results = await Query(busqueda);
            results = BorrarCajaAutomatica(results);
            Console.WriteLine("Volcando resultados de " + busquedaUser + " " + hoja);

            List<List<string>> tablaHoja = LeerHojaReporteAnterior(rutaDelDirectorioAnterior, i);
            bool creado = CrearAbrirExcelReportes(rutaDelDirectorio, hoja.ToString(), busquedaUser.Replace("%20", " "));
            if (!creado)
            {
                Console.WriteLine("Error al crear reporte");
                return;
            }
            decimal oficialUSD = await ObtenerPrecioVentaDolarOficial();
            decimal blueUSD = await ObtenerPrecioVentaDolarBlue();
            results = FiltrarResultadosRepetidos(results, tablaHoja);
            List<Result> cambiaronPrecio = ObtenerCambiaronPrecio(results, tablaHoja);
            CompletarReporte(rutaDelDirectorio, results, hoja.ToString(), oficialUSD, blueUSD, hoja);
            ModificarReporte(rutaDelDirectorioAnterior,cambiaronPrecio, hoja.ToString());
            ModificarReporte(rutaDelDirectorio, cambiaronPrecio, hoja.ToString());
            hoja++;
        }
    }

    public void ModificarReporte(string ruta, List<Result> cambiaronPrecio, string hoja)
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
                    if (cell.Text.Replace("-", "") == fila.Id.Replace("-", ""))
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

    public List<Result> FiltrarResultadosRepetidos(List<Result> results, List<List<string>> tabla)
    {
        var idsYPrecios = new HashSet<(string, decimal)>();
        foreach(var fila in tabla)
        {
            decimal precio = decimal.Parse(fila[3].ToString());
            idsYPrecios.Add((fila[0].Replace("-",""), precio));
        }
        results.RemoveAll(r => idsYPrecios.Contains((r.Id, r.Price)));
        return results;
    }
    public List<Result> ObtenerCambiaronPrecio(List<Result> results, List<List<string>> tabla)
    {
        var idsYPrecios = new HashSet<(string, decimal)>();
        var ids = new HashSet<string>();
        foreach(var fila in tabla)
        {
            decimal precio = decimal.Parse(fila[3].ToString());
            idsYPrecios.Add((fila[0].Replace("-", ""), precio));
            ids.Add(fila[0].Replace("-", ""));
        }
        return results.Where(r => ids.Contains(r.Id) && !idsYPrecios.Contains((r.Id, r.Price))).ToList();
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
        return tabla;
    }

    public async Task Procesamiento(string rutaDelDirectorio, string nombreReporte, int anioInicio, int anioFin, string busquedaUser)
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
            busqueda = busqueda.Replace(" ", "%20");
            List<Result> results = await Query(busqueda);
            results = BorrarCajaAutomatica(results);


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

    public async Task<List<Result>> Query(string busquedaUser)
    {
        string query = "search?q=<BUSQUEDA>&status=active&limit=50";
        int offset = 0;
        string offsetQuery = "&offset=" + offset.ToString();
        query = query.Replace("<BUSQUEDA>", busquedaUser);
        string queryFinal = query + offsetQuery;
        string url = "https://api.mercadolibre.com/sites/MLA/" + queryFinal;
        string token = _globals.Token;

        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {token}");
            List<Result> results = new List<Result>();
            bool finBucle = false;
            while (!finBucle)
            {
                finBucle = esFinBucle(offset);
                try
                {
                    HttpResponseMessage response = new HttpResponseMessage();
                    string errorDetails = "";
                    if (esPrimerIteracion)
                    {
                        await semaphore.WaitAsync();
                        if (!ReemplazoTokenTerminado)
                        {
                            response = await client.GetAsync(url);
                            errorDetails = await response.Content.ReadAsStringAsync();
                            if (!response.IsSuccessStatusCode)
                            {
                                var jsonObject = JsonSerializer.Deserialize<JsonElement>(errorDetails);
                                string message = jsonObject.GetProperty("message").GetString() ?? new String("");

                                if (message == "invalid access token")
                                {
                                    string newAccessToken = await Refresh();
                                    token = newAccessToken;
                                    _globals.Token = newAccessToken;
                                    string rutaDelDirectorio = AppContext.BaseDirectory;
                                    rutaDelDirectorio += "token.txt";
                                    ReemplazarContenidoArchivo(rutaDelDirectorio, newAccessToken);
                                }
                                if (client.DefaultRequestHeaders.Contains("Authorization"))
                                {
                                    client.DefaultRequestHeaders.Remove("Authorization");
                                }
                                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {_globals.Token}");
                                response = await client.GetAsync(url);
                            }
                            await Task.Delay(3000);
                            ReemplazoTokenTerminado = true;
                            semaphore.Release();
                        }
                        else
                        {
                            if (client.DefaultRequestHeaders.Contains("Authorization"))
                            {
                                client.DefaultRequestHeaders.Remove("Authorization");
                            }
                            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {_globals.Token}");
                            response = await client.GetAsync(url);
                        }
                        esPrimerIteracion = false;
                    }
                    else
                    {
                        response = await client.GetAsync(url);
                        errorDetails = await response.Content.ReadAsStringAsync();
                    }
                    

                    if (response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync();
                        // Deserializa solo la propiedad "results"
                        var options = new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true // Ignora diferencias entre mayúsculas y minúsculas
                        };
                        using var jsonDocument = JsonDocument.Parse(json);
                        var resultsJson = jsonDocument.RootElement.GetProperty("results").GetRawText();
                        List<Result> partialResults = JsonSerializer.Deserialize<List<Result>>(resultsJson, options) ?? new List<Result>();
                        results.AddRange(partialResults);
                        offset += 50;

                    }
                    else
                    {
                        Console.WriteLine($"Error en la consulta: {response.StatusCode}");
                        Console.WriteLine($"Detalles del error: {errorDetails}");
                        finBucle = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ocurrió un error: {ex.Message}");
                    finBucle = true;
                }
                offsetQuery = "&offset=" + offset.ToString();
                queryFinal = query + offsetQuery;
                url = "https://api.mercadolibre.com/sites/MLA/" + queryFinal;
            }
            return results;  
        }
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
    public bool esFinBucle (int offset)
    {
        int limit = 50;
        if((limit + offset)>3999)
        {
            return true;
        }
        return false;
    }
    public async Task<string> Refresh()
    {
        string url = "https://api.mercadolibre.com/oauth/token";
        var formContent = new FormUrlEncodedContent(new[]
        {
         new KeyValuePair<string, string>("grant_type", "refresh_token"),
         new KeyValuePair<string, string>("client_id", "2398820623563499"),
         new KeyValuePair<string, string>("client_secret", "GWR6eZHrAZu47zusWIX8MinrMnSOjhYh"),
         new KeyValuePair<string, string>("refresh_token", _globals.RefreshToken)
        });

        using (HttpClient client = new HttpClient())
        {
            client.DefaultRequestHeaders.Add("accept", "application/json");
            try
            {
                HttpResponseMessage response = await client.PostAsync(url, formContent);

                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    var jsonResponse = JsonDocument.Parse(responseBody);
                    string accessToken = jsonResponse.RootElement.GetProperty("access_token").GetString() ?? string.Empty;
                    return accessToken;
                }
                else
                {
                    Console.WriteLine("Error de api refresh");
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error: {ex.Message}");
                return string.Empty;

            }
        }
    }

    public string LeeerTokenDesdeArchivo(string ruta)
    {
        try
        {
            if (File.Exists(ruta))
            {
                string token = File.ReadAllText(ruta).Trim(); // Lee todo el contenido y elimina espacios en blanco
                return token; // Devuelve el token
            }
            else
            {
                Console.WriteLine($"El archivo no existe en la ruta especificada: {ruta}");
                return string.Empty; // Devuelve un string vacío si el archivo no existe
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ocurrió un error al leer el archivo: {ex.Message}");
            return string.Empty; // Devuelve un string vacío en caso de error
        }

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
        resultsFiltrado.RemoveAll(item => GetAttributeValue(item, "Transmisión") != "Manual");
        //resultsFiltrado.RemoveAll(item => GetAttributeValue(item, "Año") != anioInicio.ToString());
        return resultsFiltrado;
    }

    public void CompletarHistorico(string ruta, HashSet<(string anio, decimal promedio)> tuplaHistorico)
    {
        if (!File.Exists(ruta))
        {
            string rutaTemplate = Path.GetDirectoryName(ruta) + "\\Template HISTORICO.xlsx";
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
                DateTime currentDateTime = DateTime.Now;
                DateTime currentDate = currentDateTime.Date;
                string fechaActual = currentDate.ToString("dd/MM/yyyy");
                if (tabla == null)
                    return;
                
                int ultimaFilaTabla = tabla.Address.End.Row;
                while (ultimaFilaTabla > tabla.Address.Start.Row && string.IsNullOrEmpty(worksheet.Cells[ultimaFilaTabla, 1].Text))
                {
                    ultimaFilaTabla--;
                }

                worksheet.Cells[ultimaFilaTabla + 1, 1].Value = fechaActual;
                worksheet.Cells[ultimaFilaTabla + 1, 2].Value = tupla.promedio;
            }
            package.Save();
        }
    }
    public (string anio, decimal promedio) CompletarReporte(string ruta, List<Result> results, string hoja, decimal dolarOficial, decimal dolarBlue, int anioInicio)
    {
        try
        {
            var itemsOrdenados = results.OrderBy(item => item.Price).ToList();

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
                    string anio = GetAttributeValue(result, "Año");
                    string kms = GetAttributeValue(result, "Kilómetros");
                    string transmision = GetAttributeValue(result, "Transmisión");
                    decimal precio = result.Price;
                    string id = result.Id;
                    id = id.Substring(0, 3) + "-" + id.Substring(3, id.Length-3);
                    string moneda = result.Currency_id;
                    if (moneda == "USD")
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
                    string descripcion = result.Title;
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
                promUSDBlueARS /= 2;
                promUSDOficialARS += promUSD;
                promUSDOficialARS /= 2;


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
        
    }
    public class CMP
    {
        [JsonPropertyName("comparador")]
        public bool Comparador { get; set; }
        [JsonPropertyName("fecha_a_comparar")]
        public string FechaAComparar { get; set; }
    }
}





