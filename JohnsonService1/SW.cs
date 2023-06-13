using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace JohnsonService
{
    
    public partial class SW : ServiceBase
    {
        private System.Timers.Timer timer;
        string path_aplicacion = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string path = "";
          

        //string path = "C:\\Users\\rody\\Documents\\Visual Studio 2013\\Projects\\Excel\\Excel\\bin\\Debug\\";



        int codigo_distribuidor = 0;
        int codigo_proveedor = 0;
        DateTime fecha_inicio = DateTime.Now;
        string usuario = "";
        string contraseña = "";
        string hora_envio = "";
        string url_ventas = "";
        string url_inventario = "";
        string usuarioSQL = "";
        string baseSQL = "";
        string contraseñaSQL = "";
        string serverSQL = "";
        DateTime ult_envio = DateTime.Now;
        //string path = Application.StartupPath;


        public SW()
        {
             
            InitializeComponent();
            this.ServiceName = "Servicio Johnson";
            path = System.IO.Path.GetDirectoryName(path_aplicacion) + "\\";
         
        
       

            try
            {
                if (!System.Diagnostics.EventLog.SourceExists("Servicio Johnson","."))
                {

                    System.Diagnostics.EventLog.CreateEventSource("Servicio Johnson", "Application", ".");
                }

                eventLog1.Source = "Servicio Johnson";
                eventLog1.Log = "Application";

            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Un error ha ocurrido cargando el servicio.", ErrorType.Critical, ex);
            }






        }

        protected override void OnStart(string[] args)
        {



            //ejecuto cada 10 minutos.
                this.timer = new System.Timers.Timer(600000D);  // 30000 milliseconds = 30 seconds
                this.timer.AutoReset = true;
                this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Elapsed);
                this.timer.Start();
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", "Iniciando Rutina " + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Information);

        
	


        //ejecutar()

        

        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {





                CrearConfiguracionXML();
                LeerConfiguracionXML();
                LeerDatosConfiguracion();

                if (ult_envio != DateTime.Now.Date)
                {

                    DateTime horaenvio = Convert.ToDateTime(hora_envio);
                    DateTime horaenviomas10 = horaenvio.AddMinutes(10);
                    DateTime horaenviomenos10 = horaenvio.AddMinutes(-10);

                    if (DateTime.Now > horaenviomenos10 && DateTime.Now < horaenviomas10)
                    {


                        //genero los archivos de 3 dias antes si no existen
                        GenerarExcelFacturacion(-3);
                        GenerarExcelInventario(-3);

                        //genero los archivos de 2 dias antes si no existen
                        GenerarExcelFacturacion(-2);
                        GenerarExcelInventario(-2);
                        
                        
                                                
                        //Primero envio lo antiguo si es que no se pudo enviar
                        EnviarExcelFacturacionSinEnviar();
                        EnviarExcelInventarioSinEnviar();



                        //genero los archivos de hoy
                        GenerarExcelFacturacion();
                        GenerarExcelInventario();




                        //Enviando lo que corresponde al dia de la fecha 
                        EnviarExcelFacturacion();
                        EnviarExcelInventario();

                        
                        ActualizarFechaEnvio();
                        AvanzarNumeracion();

                        System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", "Envio Ejecutado Correctamente", EventLogEntryType.Information);

                    }


                    
                }



                


            }
            catch (Exception ex)
            {
                ErrorHandler.ManejarError("Ha Ocurrido un error en la ejecucion total del servicio", ErrorType.Critical, ex);
                //System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", ex.Message + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Error);
            }

            

            

        }

        private void AvanzarNumeracion()
        {

            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                string SelectConfiguracion = "declare @num integer exec numero 1,'JJX', @num OUTPUT ,1 SELECT  @NUM";
                SqlCommand commandConfiguracion = new SqlCommand(SelectConfiguracion, Conection);






                // Abrir la conexión.
                Conection.Open();
                commandConfiguracion.ExecuteNonQuery();
                Conection.Close();


            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ha Ocurrido un error avanzando la numeracion", ErrorType.Critical, ex);
            }


        


        }


        private void ActualizarFechaEnvio()
        {

            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                string SelectConfiguracion = "update johnson_configuracion set ult_envio =  Convert(varchar,getdate(),103)";
                SqlCommand commandConfiguracion = new SqlCommand(SelectConfiguracion, Conection);






                // Abrir la conexión.
                Conection.Open();
                commandConfiguracion.ExecuteNonQuery();
                Conection.Close();


            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ha Ocurrido un error actualizando la fecha de envio", ErrorType.Critical, ex);
            }

           


        }



        void LeerConfiguracionXML()
        {
            try
            {

                using (XmlReader reader = XmlReader.Create(path + "configuracion.xml"))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            //retorna solo un valor al entrar.

                            switch (reader.Name.ToString())
                            {
                                case "Server":
                                    serverSQL = reader.ReadString();
                                    break;
                                case "Base":
                                    baseSQL = reader.ReadString();
                                    break;

                                case "Usuario":
                                    usuarioSQL = reader.ReadString();
                                    break;

                                case "Contraseña":
                                    contraseña = reader.ReadString();
                                    break;
                            }
                        }

                    }
                }



            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ocurrio un error leyendo el archivo de configuracion XML", ErrorType.Critical, ex);
            }

           


        }

        //envio los archivos que no fueron enviados por falta de internet, o porque la maquina estuvo apagada en el envio.

        void EnviarExcelInventarioSinEnviar()
        {

            try
            {

                // barro todos los aarchivos que son .xlsx
                string[] fileEntries = Directory.GetFiles(path, "*.xlsx");
                //miro solo los archivos inv
                fileEntries = fileEntries.Where(q => q.Contains("_inv_")).ToArray();
                foreach (string fileName in fileEntries)
                {

                    //guardo la fecha de creacion del archivo en cuestion
                    DateTime fechaCrea = Directory.GetCreationTime(fileName);
                    DateTime fechaHoy = DateTime.Now;
                    //resto los dias que tiene  que diferencia para enviarlos
                    TimeSpan ts = fechaHoy.Date - fechaCrea.Date;
                    int diferenciaDias = ts.Days;

                   

                    //envio los archivos que tienen como maximo una diferencia de 3 dias
                    if (diferenciaDias <= 3)
                    {
                        string url = url_inventario;
                        using (var client = new WebClient())
                        {
                            client.Headers.Add(HttpRequestHeader.Authorization, "Basic " + Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(usuario + ":" + contraseña)));
                            byte[] result = client.UploadFile(url, fileName);
                            string responseAsString = Encoding.Default.GetString((result));


                        }

                        //Si el archivo existe lo borro antes.
                        if (File.Exists(fileName + ".ant") == true)
                        {

                            System.IO.File.Delete(fileName + ".ant");

                        }

                        //renombro el archivo
                        System.IO.File.Move(fileName, fileName + ".ant");
                        //Espero 70 segundos
                        Thread.Sleep(70000);
                    }


                }


                
                
            }
            catch (WebException ex)
            {
                string mensaje = "";
                HttpWebResponse httpWebResponse = (HttpWebResponse)ex.Response;
                var responseStream = ex.Response.GetResponseStream();
                if (responseStream != null) { using (var reader = new StreamReader(responseStream)) { mensaje = reader.ReadToEnd(); } }
                mensaje = mensaje.Replace(";", "\r\n");

                ErrorHandler.ManejarError(mensaje, ErrorType.Critical, ex);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", mensaje, EventLogEntryType.Error);

                //poner que guarde un mensaje en el registro
                //MessageBox.Show(mensaje);

            }



        }

        void EnviarExcelFacturacionSinEnviar()
        {

            try
            {

                // barro todos los aarchivos que son .xlsx
                string[] fileEntries = Directory.GetFiles(path, "*.xlsx");
                //miro solo los archivos fac
                fileEntries = fileEntries.Where(q => q.Contains("_fac_")).ToArray();
                foreach (string fileName in fileEntries)
                {

                    //guardo la fecha de creacion del archivo en cuestion
                    DateTime fechaCrea = Directory.GetCreationTime(fileName);
                    DateTime fechaHoy = DateTime.Now;
                    //resto los dias que tiene  que diferencia para enviarlos
                    TimeSpan ts = fechaHoy.Date - fechaCrea.Date;
                    int diferenciaDias = ts.Days;



                    //envio los archivos que tienen como maximo una diferencia de 3 dias
                    if (diferenciaDias <= 3)
                    {
                        string url = url_ventas;
                        using (var client = new WebClient())
                        {
                            client.Headers.Add(HttpRequestHeader.Authorization, "Basic " + Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(usuario + ":" + contraseña)));
                            byte[] result = client.UploadFile(url, fileName);
                            string responseAsString = Encoding.Default.GetString((result));


                        }

                        
                        //Si el archivo existe lo borro antes.
                        if (File.Exists(fileName + ".ant") == true)
                        {

                            System.IO.File.Delete(fileName + ".ant");

                        }

                        //renombro el archivo
                        System.IO.File.Move(fileName, fileName + ".ant");
                        //Espero 70 segundos
                        Thread.Sleep(70000);
                    }


                }




            }
            catch (WebException ex)
            {
                string mensaje = "";
                HttpWebResponse httpWebResponse = (HttpWebResponse)ex.Response;
                var responseStream = ex.Response.GetResponseStream();
                if (responseStream != null) { using (var reader = new StreamReader(responseStream)) { mensaje = reader.ReadToEnd(); } }
                mensaje = mensaje.Replace(";", "\r\n");

                ErrorHandler.ManejarError(mensaje, ErrorType.Critical, ex);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", mensaje, EventLogEntryType.Error);

                //poner que guarde un mensaje en el registro
                //MessageBox.Show(mensaje);

            }





        }


        // genero los archivos para enviar en el dia de la fecha, los datos se generan con la fecha de ayer.
        void GenerarExcelInventario() 
        {
            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                Conection.Open();
                
                // Utilizar una variable para almacenar la instrucción SQL.
                string SelectString = "select (select codigo_distribuidor from Johnson_Configuracion where codigo = 1) IdDistribuidor, (select num.numero from numeracion as num  where num.sucursal=1 and num.tiponum='JJ' and num.letra='X' and num.empresa_id=1) IdPaquete,  stockhistorico.fecha fecha , scj_sinonimos.CodigoSinonimo IdProducto, case when sum(stocktot) > 0 then round(sum(cast(cast(stocktot as numeric(18,2))/cast(uxbcompra as numeric(18,2)) as numeric(18,2))),2) else 0 end cantidad  from stockhistorico, articulo, scj_sinonimos where stockhistorico.Articulo_Codigo= articulo.codigo and scj_sinonimos.CodigoArt=articulo.codigo and   stockhistorico.fecha = Convert(varchar,DATEADD(d,-1,getdate()),103) and deposito_id = 0 and articulo.prove =" + codigo_proveedor + " and scj_sinonimos.CodigoSinonimo <> '' and articulo.baja = 0 group by scj_sinonimos.CodigoSinonimo, stockhistorico.fecha ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", SelectString, EventLogEntryType.Information);


                


                sqlcommand.CommandTimeout = 3600;
                SqlDataAdapter Adaptador = new SqlDataAdapter(SelectString, Conection);

                DataSet DS = new DataSet();

                // Abrir la conexión.
                
                Adaptador.SelectCommand.CommandTimeout = 3600;
                Adaptador.Fill(DS);
                Conection.Close();


                // Creamos un objeto Excel.
                Excel.Application Mi_Excel = default(Excel.Application);
                // Creamos un objeto WorkBook. Para crear el documento Excel.           
                Excel.Workbook LibroExcel = default(Excel.Workbook);
                // Creamos un objeto WorkSheet. Para crear la hoja del documento.
                Excel.Worksheet HojaExcel = default(Excel.Worksheet);
                // Creamos un objeto WorkSheet 2. Para crear la hoja del documento
                Excel.Worksheet HojaExcel2 = default(Excel.Worksheet);




                // Iniciamos una instancia a Excel, y Hacemos visibles para ver como se va creando el reporte, 
                // podemos hacerlo visible al final si se desea.

                Mi_Excel = new Excel.Application();
                Mi_Excel.DisplayAlerts = false;

                //Mi_Excel.Visible = true;

                /* Ahora creamos un nuevo documento y seleccionamos la primera hoja del 
                 * documento en la cual crearemos nuestro informe. 
                 */
                // Creamos una instancia del Workbooks de excel.            
                LibroExcel = Mi_Excel.Workbooks.Add();
                // Creamos una instancia de la primera hoja de trabajo de excel            
                HojaExcel = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[1];
                //pruebo de cambiar el nombre de la hora
                HojaExcel.Name = "datos";
                HojaExcel.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                // Hacemos esta hoja la visible en pantalla 
                // (como seleccionamos la primera esto no es necesario
                // si seleccionamos una diferente a la primera si lo
                // necesitariamos).
                HojaExcel.Activate();

                // Crear el encabezado de nuestro informe.
                // La primera línea une las celdas y las convierte un en una sola.            
                //HojaExcel.Range["A1:E1"].Merge();
                // La segunda línea Asigna el nombre del encabezado.
                HojaExcel.Range["A1:A1"].Value = "IdDistribuidor";
                HojaExcel.Range["B1:B1"].Value = "IdPaquete";
                HojaExcel.Range["C1:C1"].Value = "Fecha";
                HojaExcel.Range["D1:D1"].Value = "IdProducto";
                HojaExcel.Range["E1:E1"].Value = "Cantidad";


                // La tercera línea asigna negrita al titulo.
                // HojaExcel.Range["A1:E1"].Font.Bold = true;
                // La cuarta línea signa un Size a titulo de 15.
                //  HojaExcel.Range["A1:E1"].Font.Size = 15;

                // Crear el subencabezado de nuestro informe
                //HojaExcel.Range["A2:E2"].Merge();

                //HojaExcel.Range["A2:E2"].Value = "ENCUESTA DE SATISFACCIÓN AL CLIENTE EXTERNO";
                //HojaExcel.Range["A2:E2"].Font.Italic = true;
                //HojaExcel.Range["A2:E2"].Font.Size = 13;

                //Excel.Range objCelda = HojaExcel.Range["A3", Type.Missing];
                //objCelda.Value = "ID";

                //objCelda = HojaExcel.Range["B3", Type.Missing];
                //objCelda.Value = "Preguntas";

                //objCelda = HojaExcel.Range["C3", Type.Missing];
                //objCelda.Value = "Opciones";

                //objCelda = HojaExcel.Range["D3", Type.Missing];
                //objCelda.Value = "Valor de la Respuesta";

                //objCelda = HojaExcel.Range["E3", Type.Missing];
                //objCelda.Value = "Numero Votos";

                //objCelda.EntireColumn.NumberFormat = "###,###,###.00";

                int i = 2;
                foreach (DataRow Row in DS.Tables[0].Rows)
                {
                    // IdDistribuidor
                    HojaExcel.Cells[i, "A"] = Row.ItemArray[0];
                    // IdPaquete
                    HojaExcel.Cells[i, "B"] = Row.ItemArray[1];
                    // Fecha
                    HojaExcel.Cells[i, "C"] = Row.ItemArray[2];
                    // IdProducto
                    HojaExcel.Cells[i, "D"] = Row.ItemArray[3];
                    // Cantidad
                    HojaExcel.Cells[i, "E"] = Row.ItemArray[4];

                    // Avanzamos una fila
                    i++;
                }

                // Seleccionar todo el bloque desde A1 hasta D #de filas.
                //Excel.Range Rango = HojaExcel.Range["A3:E" + (i - 1).ToString()];

                //// Selecionado todo el rango especificado

                //Rango.Select();

                //// Ajustamos el ancho de las columnas al ancho máximo del
                //// contenido de sus celdas
                //Rango.Columns.AutoFit();

                //// Asignar filtro por columna
                //Rango.AutoFilter(1);

                //creo otra hoja para el control


                HojaExcel2 = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[2];
                //pruebo de cambiar el nombre de la hora
                HojaExcel2.Name = "verificacion";
                HojaExcel2.Activate();
                HojaExcel2.Range["A1:A1"].Value = "INDICADOR";
                HojaExcel2.Range["B1:B1"].Value = "VALOR";


                HojaExcel2.Range["A2:A2"].Value = "CantRegistros";
                //cuento la cantidad de registros
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!e2:datos!e99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;


                HojaExcel2.Range["A3:A3"].Value = "TotalCajas";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!e2:datos!e99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;
                //guardo el documento


                string pathfull = path + usuario + "_inv_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd")  + ".xlsx";
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", pathfull, EventLogEntryType.Information);
                LibroExcel.SaveAs(pathfull, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                LibroExcel.Close(true);

                // Crear un total general
                //LibroExcel.PrintPreview();

        


            }
            catch (Exception ex)
            {
                
                ErrorHandler.ManejarError("Ocurrio un error Generando el Archivo de Inventario", ErrorType.Critical, ex);
                
            }
        
        
        }

        void GenerarExcelFacturacion() 
        {
            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);








                // Abrir la conexión.
                Conection.Open();

                //comentario para compilar a asdsadsa
                string SelectString = "select (select codigo_distribuidor from Johnson_Configuracion where codigo = 1) idDistribuidor , (select num.numero from numeracion as num  where num.sucursal=1 and num.tiponum='JJ' and num.letra='X' and num.empresa_id=1) IdPaquete, v.fecha fecha , v.comprobante NroComprobante, '' IdPedidoDinesys, case left(referencia,2) when 'NC' THEN referencia END NroComprobanteAsociado,  cod_cli IdCliente, Jonhson_Tipos_Cliente.Codigo IdTipoDeCliente,  cod_ven IdVendedor, scj.codigoSinonimo IdProducto , round(sum(iv.Cantidad*art.uxb/art.uxbcompra),2) Cantidad, case WHEN  sum(iv.Cantidad*art.uxb/art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('FA', 'PE')  THEN 'FC' when sum(iv.Cantidad* art.uxb/ art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('ND') then 'ND' else 'NC' END Tipo,     motivo.descripcion MotivoNC  from numeracion, venta v inner join  item_venta iv on v.comprobante=iv.comprobante  and v.fecha = iv.fecha and v.empresa_id= iv.empresa_id left join motivo on v.motivodev = motivo.codigo  left join articulo art on iv.cod_art=art.codigo inner join scj_sinonimos scj on scj.CodigoArt=art.codigo inner join cliente cli on cli.codigo=v.cod_cli inner join ramo on cli.ramo=ramo.codigo inner join Ramo_Vs_RamoJohnson on Ramo_Vs_RamoJohnson.codigo=ramo.codigo inner join Jonhson_Tipos_Cliente on Jonhson_Tipos_Cliente.Id=Ramo_Vs_RamoJohnson.IDJohnson where TipoNum = 'JJ' and v.fecha = Convert(varchar,DATEADD(d,-1,getdate()),103) and left(v.comprobante, 2) in ('FA','NC','ND','PE')  and art.prove =1 and scj.CodigoSinonimo <> '' and art.baja = 0 and fecha_factura is not null   group by v.fecha, v.comprobante, v.referencia, v.cod_cli, v.cod_ven, scj.CodigoSinonimo, motivo.descripcion, Jonhson_Tipos_Cliente.Codigo having sum(iv.Cantidad*art.uxb/art.uxbcompra) <> 0 ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                sqlcommand.CommandTimeout = 3600;
                SqlDataAdapter Adaptador = new SqlDataAdapter(SelectString, Conection);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", SelectString, EventLogEntryType.Information);


                DataSet DS = new DataSet();
                Adaptador.Fill(DS);
                Conection.Close();


                // Creamos un objeto Excel.
                Excel.Application Mi_Excel = default(Excel.Application);
                // Creamos un objeto WorkBook. Para crear el documento Excel.           
                Excel.Workbook LibroExcel = default(Excel.Workbook);
                // Creamos un objeto WorkSheet. Para crear la hoja del documento.
                Excel.Worksheet HojaExcel = default(Excel.Worksheet);
                // Creamos un objeto WorkSheet 2. Para crear la hoja del documento
                Excel.Worksheet HojaExcel2 = default(Excel.Worksheet);




                // Iniciamos una instancia a Excel, y Hacemos visibles para ver como se va creando el reporte, 
                // podemos hacerlo visible al final si se desea.

                Mi_Excel = new Excel.Application();
                Mi_Excel.DisplayAlerts = false;

                //Mi_Excel.Visible = true;

                /* Ahora creamos un nuevo documento y seleccionamos la primera hoja del 
                 * documento en la cual crearemos nuestro informe. 
                 */
                // Creamos una instancia del Workbooks de excel.            
                LibroExcel = Mi_Excel.Workbooks.Add();
                // Creamos una instancia de la primera hoja de trabajo de excel            
                HojaExcel = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[1];
                //pruebo de cambiar el nombre de la hora
                HojaExcel.Name = "datos";
                HojaExcel.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                // Hacemos esta hoja la visible en pantalla 
                // (como seleccionamos la primera esto no es necesario
                // si seleccionamos una diferente a la primera si lo
                // necesitariamos).
                HojaExcel.Activate();

                // Crear el encabezado de nuestro informe.
                // La primera línea une las celdas y las convierte un en una sola.            
                //HojaExcel.Range["A1:E1"].Merge();
                // La segunda línea Asigna el nombre del encabezado.
                HojaExcel.Range["A1:A1"].Value = "IdDistribuidor";
                HojaExcel.Range["B1:B1"].Value = "IdPaquete";
                HojaExcel.Range["C1:C1"].Value = "Fecha";
                HojaExcel.Range["D1:D1"].Value = "NroComprobante";
                HojaExcel.Range["E1:E1"].Value = "IdPedidoDinesys";
                HojaExcel.Range["F1:F1"].Value = "NroComprobanteAsociado";
                HojaExcel.Range["G1:G1"].Value = "IdCliente";
                HojaExcel.Range["H1:H1"].Value = "IdTipoDeCliente";
                HojaExcel.Range["I1:I1"].Value = "IdVendedor";
                HojaExcel.Range["J1:J1"].Value = "IdProducto";
                HojaExcel.Range["K1:K1"].Value = "Cantidad";
                HojaExcel.Range["L1:L1"].Value = "TipoDeComprobante";
                HojaExcel.Range["M1:M1"].Value = "MotivoNC";

                // La tercera línea asigna negrita al titulo.
                // HojaExcel.Range["A1:E1"].Font.Bold = true;
                // La cuarta línea signa un Size a titulo de 15.
                //  HojaExcel.Range["A1:E1"].Font.Size = 15;

                // Crear el subencabezado de nuestro informe
                //HojaExcel.Range["A2:E2"].Merge();

                //HojaExcel.Range["A2:E2"].Value = "ENCUESTA DE SATISFACCIÓN AL CLIENTE EXTERNO";
                //HojaExcel.Range["A2:E2"].Font.Italic = true;
                //HojaExcel.Range["A2:E2"].Font.Size = 13;

                //Excel.Range objCelda = HojaExcel.Range["A3", Type.Missing];
                //objCelda.Value = "ID";

                //objCelda = HojaExcel.Range["B3", Type.Missing];
                //objCelda.Value = "Preguntas";

                //objCelda = HojaExcel.Range["C3", Type.Missing];
                //objCelda.Value = "Opciones";

                //objCelda = HojaExcel.Range["D3", Type.Missing];
                //objCelda.Value = "Valor de la Respuesta";

                //objCelda = HojaExcel.Range["E3", Type.Missing];
                //objCelda.Value = "Numero Votos";

                //objCelda.EntireColumn.NumberFormat = "###,###,###.00";

                int i = 2;
                foreach (DataRow Row in DS.Tables[0].Rows)
                {
                    // Asignar los valores de los registros a las celdas
                    HojaExcel.Cells[i, "A"] = Row.ItemArray[0];
                    // IdDistribuidor
                    HojaExcel.Cells[i, "B"] = Row.ItemArray[1];
                    // IdPaquete
                    HojaExcel.Cells[i, "C"] = Row.ItemArray[2];
                    // Fecha
                    HojaExcel.Cells[i, "D"] = Row.ItemArray[3];
                    // NroComprobante
                    HojaExcel.Cells[i, "E"] = Row.ItemArray[4];
                    // IdPedidoDinesys
                    HojaExcel.Cells[i, "F"] = Row.ItemArray[5];

                    HojaExcel.Cells[i, "G"] = Row.ItemArray[6];

                    HojaExcel.Cells[i, "H"] = Row.ItemArray[7];

                    HojaExcel.Cells[i, "I"] = Row.ItemArray[8];

                    HojaExcel.Cells[i, "J"] = Row.ItemArray[9];

                    HojaExcel.Cells[i, "K"] = Row.ItemArray[10];

                    HojaExcel.Cells[i, "L"] = Row.ItemArray[11];

                    HojaExcel.Cells[i, "M"] = Row.ItemArray[12];

                    // Avanzamos una fila
                    i++;
                }

                // Seleccionar todo el bloque desde A1 hasta D #de filas.
                //Excel.Range Rango = HojaExcel.Range["A3:E" + (i - 1).ToString()];

                //// Selecionado todo el rango especificado

                //Rango.Select();

                //// Ajustamos el ancho de las columnas al ancho máximo del
                //// contenido de sus celdas
                //Rango.Columns.AutoFit();

                //// Asignar filtro por columna
                //Rango.AutoFilter(1);

                //creo otra hoja para el controlDateTime.Today.ToString("yyyyMMdd")


                HojaExcel2 = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[2];
                //pruebo de cambiar el nombre de la hora
                HojaExcel2.Name = "verificacion";
                HojaExcel2.Activate();
                HojaExcel2.Range["A1:A1"].Value = "INDICADOR";
                HojaExcel2.Range["B1:B1"].Value = "VALOR";


                HojaExcel2.Range["A2:A2"].Value = "CantRegistros";
                //cuento la cantidad de registros
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!K2:datos!K99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;

                HojaExcel2.Range["A3:A3"].Value = "TotalCajas";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!K2:datos!K99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;
                //guardo el documento


                string pathfull = path + usuario + "_fac_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd")  + ".xlsx";
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", pathfull, EventLogEntryType.Information);
                LibroExcel.SaveAs(pathfull, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                //LibroExcel.SaveAs("E:\\dibert_fac_" + DateTime.Today.ToString("yyyyMMdd") + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                LibroExcel.Close(true);

                // Crear un total general
                //LibroExcel.PrintPreview();


            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ocurrio un error Generando el Archivo de Ventas", ErrorType.Critical, ex);
            }
           


        }


        // Se envian los archivos generados en el día de la fecha

        void EnviarExcelInventario()
        {

            try
            {

                string fileToUpload = path + usuario + "_inv_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd") + ".xlsx";
                string url = url_inventario;
                using (var client = new WebClient())
                {
                    client.Headers.Add(HttpRequestHeader.Authorization, "Basic " + Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(usuario + ":" + contraseña)));
                    byte[] result = client.UploadFile(url, fileToUpload);
                    string responseAsString = Encoding.Default.GetString((result));


                }


                //Si el archivo existe lo borro antes.
                if (File.Exists(fileToUpload + ".ant") == true)
                {

                    System.IO.File.Delete(fileToUpload + ".ant");

                }

                //renombro el archivo para que no se vuelva a enviar.
                System.IO.File.Move(fileToUpload, fileToUpload + ".ant");

                //espero 70 segundos para el proximo envio.
                Thread.Sleep(70000);

            }
            catch (WebException ex)
            {
                string mensaje = "";
                HttpWebResponse httpWebResponse = (HttpWebResponse)ex.Response;
                var responseStream = ex.Response.GetResponseStream();
                if (responseStream != null) { using (var reader = new StreamReader(responseStream)) { mensaje = reader.ReadToEnd(); } }
                mensaje = mensaje.Replace(";", "\r\n");

                ErrorHandler.ManejarError(mensaje, ErrorType.Critical, ex);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", mensaje, EventLogEntryType.Error);

                //poner que guarde un mensaje en el registro
                //MessageBox.Show(mensaje);

            }



        }

        void EnviarExcelFacturacion()
        {

            try
            {
                string fileToUpload = path + usuario + "_fac_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd") + ".xlsx";
                string url = url_ventas;

                using (var client = new WebClient())
                {
                    client.Headers.Add(HttpRequestHeader.Authorization, "Basic " + Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(usuario + ":" + contraseña)));
                    byte[] result = client.UploadFile(url, fileToUpload);
                    string responseAsString = Encoding.Default.GetString(result);

                }


                //Si el archivo existe lo borro antes.
                if (File.Exists(fileToUpload + ".ant") == true)
                {

                    System.IO.File.Delete(fileToUpload + ".ant");

                }

                //renombro el archivo para que no se vuelva a enviar.
                System.IO.File.Move(fileToUpload, fileToUpload + ".ant");
                //espero 70 segundos para el proximo envio.
                Thread.Sleep(70000);

            }
            catch (WebException ex)
            {
                string mensaje = "";
                HttpWebResponse httpWebResponse = (HttpWebResponse)ex.Response;
                var responseStream = ex.Response.GetResponseStream();
                if (responseStream != null) { using (var reader = new StreamReader(responseStream)) { mensaje = reader.ReadToEnd(); } }
                mensaje = mensaje.Replace(";", "\r\n");
                ErrorHandler.ManejarError(mensaje, ErrorType.Critical, ex);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", mensaje, EventLogEntryType.Error);

                //poner que gaurde un mensaje en el
                //MessageBox.Show(mensaje);

            }


        }


        // genero archivos que no hayan sido generados por algun motivo

        void GenerarExcelInventario(int vDias)
        {
            try
            {

                if (File.Exists(path + usuario + "_inv_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx.ant") == true) return;

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                Conection.Open();

                // Utilizar una variable para almacenar la instrucción SQL.
                string SelectString = "select (select codigo_distribuidor from Johnson_Configuracion where codigo = 1) IdDistribuidor, (select num.numero from numeracion as num  where num.sucursal=1 and num.tiponum='JJ' and num.letra='X' and num.empresa_id=1) IdPaquete,  stockhistorico.fecha fecha , scj_sinonimos.CodigoSinonimo IdProducto, case when sum(stocktot) > 0 then round(sum(cast(cast(stocktot as numeric(18,2))/cast(uxbcompra as numeric(18,2)) as numeric(18,2))),2) else 0 end cantidad  from stockhistorico, articulo, scj_sinonimos where stockhistorico.Articulo_Codigo= articulo.codigo and scj_sinonimos.CodigoArt=articulo.codigo and   stockhistorico.fecha = Convert(varchar,DATEADD(d," + vDias + ",getdate()),103) and deposito_id = 0 and articulo.prove =" + codigo_proveedor + " and scj_sinonimos.CodigoSinonimo <> '' and articulo.baja = 0 group by scj_sinonimos.CodigoSinonimo, stockhistorico.fecha ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", SelectString, EventLogEntryType.Information);





                sqlcommand.CommandTimeout = 3600;
                SqlDataAdapter Adaptador = new SqlDataAdapter(SelectString, Conection);

                DataSet DS = new DataSet();

                // Abrir la conexión.

                Adaptador.SelectCommand.CommandTimeout = 3600;
                Adaptador.Fill(DS);
                Conection.Close();


                // Creamos un objeto Excel.
                Excel.Application Mi_Excel = default(Excel.Application);
                // Creamos un objeto WorkBook. Para crear el documento Excel.           
                Excel.Workbook LibroExcel = default(Excel.Workbook);
                // Creamos un objeto WorkSheet. Para crear la hoja del documento.
                Excel.Worksheet HojaExcel = default(Excel.Worksheet);
                // Creamos un objeto WorkSheet 2. Para crear la hoja del documento
                Excel.Worksheet HojaExcel2 = default(Excel.Worksheet);




                // Iniciamos una instancia a Excel, y Hacemos visibles para ver como se va creando el reporte, 
                // podemos hacerlo visible al final si se desea.

                Mi_Excel = new Excel.Application();
                Mi_Excel.DisplayAlerts = false;

                //Mi_Excel.Visible = true;

                /* Ahora creamos un nuevo documento y seleccionamos la primera hoja del 
                 * documento en la cual crearemos nuestro informe. 
                 */
                // Creamos una instancia del Workbooks de excel.            
                LibroExcel = Mi_Excel.Workbooks.Add();
                // Creamos una instancia de la primera hoja de trabajo de excel            
                HojaExcel = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[1];
                //pruebo de cambiar el nombre de la hora
                HojaExcel.Name = "datos";
                HojaExcel.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                // Hacemos esta hoja la visible en pantalla 
                // (como seleccionamos la primera esto no es necesario
                // si seleccionamos una diferente a la primera si lo
                // necesitariamos).
                HojaExcel.Activate();

                // Crear el encabezado de nuestro informe.
                // La primera línea une las celdas y las convierte un en una sola.            
                //HojaExcel.Range["A1:E1"].Merge();
                // La segunda línea Asigna el nombre del encabezado.
                HojaExcel.Range["A1:A1"].Value = "IdDistribuidor";
                HojaExcel.Range["B1:B1"].Value = "IdPaquete";
                HojaExcel.Range["C1:C1"].Value = "Fecha";
                HojaExcel.Range["D1:D1"].Value = "IdProducto";
                HojaExcel.Range["E1:E1"].Value = "Cantidad";


                // La tercera línea asigna negrita al titulo.
                // HojaExcel.Range["A1:E1"].Font.Bold = true;
                // La cuarta línea signa un Size a titulo de 15.
                //  HojaExcel.Range["A1:E1"].Font.Size = 15;

                // Crear el subencabezado de nuestro informe
                //HojaExcel.Range["A2:E2"].Merge();

                //HojaExcel.Range["A2:E2"].Value = "ENCUESTA DE SATISFACCIÓN AL CLIENTE EXTERNO";
                //HojaExcel.Range["A2:E2"].Font.Italic = true;
                //HojaExcel.Range["A2:E2"].Font.Size = 13;

                //Excel.Range objCelda = HojaExcel.Range["A3", Type.Missing];
                //objCelda.Value = "ID";

                //objCelda = HojaExcel.Range["B3", Type.Missing];
                //objCelda.Value = "Preguntas";

                //objCelda = HojaExcel.Range["C3", Type.Missing];
                //objCelda.Value = "Opciones";

                //objCelda = HojaExcel.Range["D3", Type.Missing];
                //objCelda.Value = "Valor de la Respuesta";

                //objCelda = HojaExcel.Range["E3", Type.Missing];
                //objCelda.Value = "Numero Votos";

                //objCelda.EntireColumn.NumberFormat = "###,###,###.00";

                int i = 2;
                foreach (DataRow Row in DS.Tables[0].Rows)
                {
                    // IdDistribuidor
                    HojaExcel.Cells[i, "A"] = Row.ItemArray[0];
                    // IdPaquete
                    HojaExcel.Cells[i, "B"] = Row.ItemArray[1];
                    // Fecha
                    HojaExcel.Cells[i, "C"] = Row.ItemArray[2];
                    // IdProducto
                    HojaExcel.Cells[i, "D"] = Row.ItemArray[3];
                    // Cantidad
                    HojaExcel.Cells[i, "E"] = Row.ItemArray[4];

                    // Avanzamos una fila
                    i++;
                }

                // Seleccionar todo el bloque desde A1 hasta D #de filas.
                //Excel.Range Rango = HojaExcel.Range["A3:E" + (i - 1).ToString()];

                //// Selecionado todo el rango especificado

                //Rango.Select();

                //// Ajustamos el ancho de las columnas al ancho máximo del
                //// contenido de sus celdas
                //Rango.Columns.AutoFit();

                //// Asignar filtro por columna
                //Rango.AutoFilter(1);

                //creo otra hoja para el control


                HojaExcel2 = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[2];
                //pruebo de cambiar el nombre de la hora
                HojaExcel2.Name = "verificacion";
                HojaExcel2.Activate();
                HojaExcel2.Range["A1:A1"].Value = "INDICADOR";
                HojaExcel2.Range["B1:B1"].Value = "VALOR";


                HojaExcel2.Range["A2:A2"].Value = "CantRegistros";
                //cuento la cantidad de registros
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!e2:datos!e99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;


                HojaExcel2.Range["A3:A3"].Value = "TotalCajas";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!e2:datos!e99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;


                string pathfull = path + usuario + "_inv_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx";
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", pathfull, EventLogEntryType.Information);
                LibroExcel.SaveAs(pathfull, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                LibroExcel.Close(true);

                // Crear un total general
                //LibroExcel.PrintPreview();




            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ocurrio un error Generando el Archivo de Inventario", ErrorType.Critical, ex);

            }


        }

        void GenerarExcelFacturacion( int vDias )
        {
            try
            {

                 if (File.Exists(path + usuario + "_fac_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx.ant") == true) return;

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);








                // Abrir la conexión.
                Conection.Open();

                //comentario para compilar a asdsadsa
                string SelectString = "select (select codigo_distribuidor from Johnson_Configuracion where codigo = 1) idDistribuidor , (select num.numero from numeracion as num  where num.sucursal=1 and num.tiponum='JJ' and num.letra='X' and num.empresa_id=1) IdPaquete, v.fecha fecha , v.comprobante NroComprobante, '' IdPedidoDinesys, case left(referencia,2) when 'NC' THEN referencia END NroComprobanteAsociado,  cod_cli IdCliente, Jonhson_Tipos_Cliente.Codigo IdTipoDeCliente,  cod_ven IdVendedor, scj.codigoSinonimo IdProducto , round(sum(iv.Cantidad*art.uxb/art.uxbcompra),2) Cantidad, case WHEN  sum(iv.Cantidad*art.uxb/art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('FA', 'PE')  THEN 'FC' when sum(iv.Cantidad* art.uxb/ art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('ND') then 'ND' else 'NC' END Tipo,     motivo.descripcion MotivoNC  from numeracion, venta v inner join  item_venta iv on v.comprobante=iv.comprobante  and v.fecha = iv.fecha and v.empresa_id= iv.empresa_id left join motivo on v.motivodev = motivo.codigo  left join articulo art on iv.cod_art=art.codigo inner join scj_sinonimos scj on scj.CodigoArt=art.codigo inner join cliente cli on cli.codigo=v.cod_cli inner join ramo on cli.ramo=ramo.codigo inner join Ramo_Vs_RamoJohnson on Ramo_Vs_RamoJohnson.codigo=ramo.codigo inner join Jonhson_Tipos_Cliente on Jonhson_Tipos_Cliente.Id=Ramo_Vs_RamoJohnson.IDJohnson where TipoNum = 'JJ' and v.fecha = Convert(varchar,DATEADD(d," + vDias + ",getdate()),103) and left(v.comprobante, 2) in ('FA','NC','ND','PE')  and art.prove =1 and scj.CodigoSinonimo <> '' and art.baja = 0 and fecha_factura is not null   group by v.fecha, v.comprobante, v.referencia, v.cod_cli, v.cod_ven, scj.CodigoSinonimo, motivo.descripcion, Jonhson_Tipos_Cliente.Codigo having sum(iv.Cantidad*art.uxb/art.uxbcompra) <> 0 ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                sqlcommand.CommandTimeout = 3600;
                SqlDataAdapter Adaptador = new SqlDataAdapter(SelectString, Conection);
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", SelectString, EventLogEntryType.Information);


                DataSet DS = new DataSet();
                Adaptador.Fill(DS);
                Conection.Close();


                // Creamos un objeto Excel.
                Excel.Application Mi_Excel = default(Excel.Application);
                // Creamos un objeto WorkBook. Para crear el documento Excel.           
                Excel.Workbook LibroExcel = default(Excel.Workbook);
                // Creamos un objeto WorkSheet. Para crear la hoja del documento.
                Excel.Worksheet HojaExcel = default(Excel.Worksheet);
                // Creamos un objeto WorkSheet 2. Para crear la hoja del documento
                Excel.Worksheet HojaExcel2 = default(Excel.Worksheet);




                // Iniciamos una instancia a Excel, y Hacemos visibles para ver como se va creando el reporte, 
                // podemos hacerlo visible al final si se desea.

                Mi_Excel = new Excel.Application();
                Mi_Excel.DisplayAlerts = false;

                //Mi_Excel.Visible = true;

                /* Ahora creamos un nuevo documento y seleccionamos la primera hoja del 
                 * documento en la cual crearemos nuestro informe. 
                 */
                // Creamos una instancia del Workbooks de excel.            
                LibroExcel = Mi_Excel.Workbooks.Add();
                // Creamos una instancia de la primera hoja de trabajo de excel            
                HojaExcel = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[1];
                //pruebo de cambiar el nombre de la hora
                HojaExcel.Name = "datos";
                HojaExcel.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                // Hacemos esta hoja la visible en pantalla 
                // (como seleccionamos la primera esto no es necesario
                // si seleccionamos una diferente a la primera si lo
                // necesitariamos).
                HojaExcel.Activate();

                // Crear el encabezado de nuestro informe.
                // La primera línea une las celdas y las convierte un en una sola.            
                //HojaExcel.Range["A1:E1"].Merge();
                // La segunda línea Asigna el nombre del encabezado.
                HojaExcel.Range["A1:A1"].Value = "IdDistribuidor";
                HojaExcel.Range["B1:B1"].Value = "IdPaquete";
                HojaExcel.Range["C1:C1"].Value = "Fecha";
                HojaExcel.Range["D1:D1"].Value = "NroComprobante";
                HojaExcel.Range["E1:E1"].Value = "IdPedidoDinesys";
                HojaExcel.Range["F1:F1"].Value = "NroComprobanteAsociado";
                HojaExcel.Range["G1:G1"].Value = "IdCliente";
                HojaExcel.Range["H1:H1"].Value = "IdTipoDeCliente";
                HojaExcel.Range["I1:I1"].Value = "IdVendedor";
                HojaExcel.Range["J1:J1"].Value = "IdProducto";
                HojaExcel.Range["K1:K1"].Value = "Cantidad";
                HojaExcel.Range["L1:L1"].Value = "TipoDeComprobante";
                HojaExcel.Range["M1:M1"].Value = "MotivoNC";

                // La tercera línea asigna negrita al titulo.
                // HojaExcel.Range["A1:E1"].Font.Bold = true;
                // La cuarta línea signa un Size a titulo de 15.
                //  HojaExcel.Range["A1:E1"].Font.Size = 15;

                // Crear el subencabezado de nuestro informe
                //HojaExcel.Range["A2:E2"].Merge();

                //HojaExcel.Range["A2:E2"].Value = "ENCUESTA DE SATISFACCIÓN AL CLIENTE EXTERNO";
                //HojaExcel.Range["A2:E2"].Font.Italic = true;
                //HojaExcel.Range["A2:E2"].Font.Size = 13;

                //Excel.Range objCelda = HojaExcel.Range["A3", Type.Missing];
                //objCelda.Value = "ID";

                //objCelda = HojaExcel.Range["B3", Type.Missing];
                //objCelda.Value = "Preguntas";

                //objCelda = HojaExcel.Range["C3", Type.Missing];
                //objCelda.Value = "Opciones";

                //objCelda = HojaExcel.Range["D3", Type.Missing];
                //objCelda.Value = "Valor de la Respuesta";

                //objCelda = HojaExcel.Range["E3", Type.Missing];
                //objCelda.Value = "Numero Votos";

                //objCelda.EntireColumn.NumberFormat = "###,###,###.00";

                int i = 2;
                foreach (DataRow Row in DS.Tables[0].Rows)
                {
                    // Asignar los valores de los registros a las celdas
                    HojaExcel.Cells[i, "A"] = Row.ItemArray[0];
                    // IdDistribuidor
                    HojaExcel.Cells[i, "B"] = Row.ItemArray[1];
                    // IdPaquete
                    HojaExcel.Cells[i, "C"] = Row.ItemArray[2];
                    // Fecha
                    HojaExcel.Cells[i, "D"] = Row.ItemArray[3];
                    // NroComprobante
                    HojaExcel.Cells[i, "E"] = Row.ItemArray[4];
                    // IdPedidoDinesys
                    HojaExcel.Cells[i, "F"] = Row.ItemArray[5];

                    HojaExcel.Cells[i, "G"] = Row.ItemArray[6];

                    HojaExcel.Cells[i, "H"] = Row.ItemArray[7];

                    HojaExcel.Cells[i, "I"] = Row.ItemArray[8];

                    HojaExcel.Cells[i, "J"] = Row.ItemArray[9];

                    HojaExcel.Cells[i, "K"] = Row.ItemArray[10];

                    HojaExcel.Cells[i, "L"] = Row.ItemArray[11];

                    HojaExcel.Cells[i, "M"] = Row.ItemArray[12];

                    // Avanzamos una fila
                    i++;
                }

                // Seleccionar todo el bloque desde A1 hasta D #de filas.
                //Excel.Range Rango = HojaExcel.Range["A3:E" + (i - 1).ToString()];

                //// Selecionado todo el rango especificado

                //Rango.Select();

                //// Ajustamos el ancho de las columnas al ancho máximo del
                //// contenido de sus celdas
                //Rango.Columns.AutoFit();

                //// Asignar filtro por columna
                //Rango.AutoFilter(1);

                //creo otra hoja para el controlDateTime.Today.ToString("yyyyMMdd")


                HojaExcel2 = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[2];
                //pruebo de cambiar el nombre de la hora
                HojaExcel2.Name = "verificacion";
                HojaExcel2.Activate();
                HojaExcel2.Range["A1:A1"].Value = "INDICADOR";
                HojaExcel2.Range["B1:B1"].Value = "VALOR";


                HojaExcel2.Range["A2:A2"].Value = "CantRegistros";
                //cuento la cantidad de registros
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!K2:datos!K99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;

                HojaExcel2.Range["A3:A3"].Value = "TotalCajas";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!K2:datos!K99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;
                //guardo el documento



                string pathfull = path + usuario + "_fac_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx";
                System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", pathfull, EventLogEntryType.Information);
                LibroExcel.SaveAs(pathfull, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                //LibroExcel.SaveAs("E:\\dibert_fac_" + DateTime.Today.ToString("yyyyMMdd") + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                LibroExcel.Close(true);

                // Crear un total general
                //LibroExcel.PrintPreview();


            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ocurrio un error Generando el Archivo de Ventas", ErrorType.Critical, ex);
            }



        }


        void LeerDatosConfiguracion()
        {
            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
             
                //comento porque ya no necesito esto.
                //System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", "cadena de conexion:" + cadena + " Fecha:" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Information);
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                string SelectConfiguracion = "select * from Johnson_Configuracion where codigo = 1";
                SqlCommand commandConfiguracion = new SqlCommand(SelectConfiguracion, Conection);






                // Abrir la conexión.
                
                Conection.Open();
                SqlDataReader reader = commandConfiguracion.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        // Console.WriteLine("{0}\t{1}", reader.GetInt32(0),
                        //     reader.GetString(1));

                        codigo_distribuidor = reader.GetInt32(1);
                        codigo_proveedor = reader.GetInt32(2);
                        fecha_inicio = reader.GetDateTime(3);
                        usuario = reader.GetString(4);
                        contraseña = reader.GetString(5);
                        hora_envio = reader.GetString(6);
                        url_ventas = reader.GetString(7);
                        url_inventario = reader.GetString(8);
                        ult_envio = reader.GetDateTime(9); 


                    }
                }
                else
                {
                    System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", "El distribuidor no tiene seteado los datos de configuracion" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Error);
                }
                reader.Close();
                Conection.Close();


            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ocurrio un error leyendo los datos de configuracion del distribuidor", ErrorType.Critical, ex);
            }

           
        }

        
        void CrearConfiguracionXML()
        {

            try
            {


                if (File.Exists(path + "configuracion.xml") == false)
                {

                    using (XmlWriter writer = XmlWriter.Create(path + "configuracion.xml"))
                    {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("Configuraciones");


                        writer.WriteStartElement("Servidor");

                        writer.WriteElementString("Server", "localhost");
                        writer.WriteElementString("Base", "idea");
                        writer.WriteElementString("Usuario", "usuario");
                        writer.WriteElementString("Contraseña", "jeje");




                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }



                }




            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Error creando la Configuracion XML", ErrorType.Critical, ex);
            }


           
        }

        protected override void OnStop()
        {

            this.timer.Stop();
            this.timer = null;
            System.Diagnostics.EventLog.WriteEntry("Servicio Johnson", "Rutina de Ejecutar Terminada " + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Information);


        }



    }
}
