using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
namespace BDFService
{
    static class Program
    {
        static string path_aplicacion = System.Reflection.Assembly.GetExecutingAssembly().Location;
        static string path = "";
        static int codigo_distribuidor = 0;
        static int codigo_proveedor = 0;
        static DateTime fecha_inicio = DateTime.Now;
        static string usuario = "";
        static string contraseña = "";
        static string hora_envio = "";
        static string url_ventas = "";
        static string url_inventario = "";
        static string usuarioSQL = "";
        static string baseSQL = "";
        static string contraseñaSQL = "";
        static string serverSQL = "";
        static DateTime ult_envio = DateTime.Now;

        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {
            //ServiceBase[] ServicesToRun;
            //ServicesToRun = new ServiceBase[] 
            //{ 
            //    new SW() 
            //};
            //ServiceBase.Run(ServicesToRun);
            path = System.IO.Path.GetDirectoryName(path_aplicacion) + "\\";


            CrearConfiguracionXML();
            LeerConfiguracionXML();
            LeerDatosConfiguracion();
            Console.WriteLine("Hola");
            if (ult_envio != DateTime.Now.Date)
            {

                DateTime horaenvio = Convert.ToDateTime(hora_envio);
                DateTime horaenviomas10 = horaenvio.AddMinutes(10);
                DateTime horaenviomenos10 = horaenvio.AddMinutes(-10);

              //  if (DateTime.Now > horaenviomenos10 && DateTime.Now < horaenviomas10)
             //   {


                    //genero los archivos de 3 dias antes si no existen
                    GenerarExcelFacturacion(-1);
                    GenerarExcelInventario(-1);
                    GenerarExcelClientes(-1);

                    //genero los archivos de 2 dias antes si no existen
        //             GenerarExcelFacturacion(-2);
        //         GenerarExcelInventario(-2);



                    //Primero envio lo antiguo si es que no se pudo enviar
                 //   EnviarExcelFacturacionSinEnviar();
                 //   EnviarExcelInventarioSinEnviar();



                    //genero los archivos de hoy
                //    GenerarExcelFacturacion();
                 //   GenerarExcelInventario();




                    //Enviando lo que corresponde al dia de la fecha 
                    EnviarExcelFacturacion();
                    EnviarExcelInventario();


                    ActualizarFechaEnvio();
                    AvanzarNumeracion();

                    System.Diagnostics.EventLog.WriteEntry("Servicio BDF", "Envio Ejecutado Correctamente", EventLogEntryType.Information);

            //    }

            }

        }
        private static void AvanzarNumeracion()
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


        private static  void ActualizarFechaEnvio()
        {

            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                string SelectConfiguracion = "update BDF_configuracion set ult_envio =  Convert(varchar,getdate(),103)";
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



        static void LeerConfiguracionXML()
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

        static void EnviarExcelInventarioSinEnviar()
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
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", mensaje, EventLogEntryType.Error);

                //poner que guarde un mensaje en el registro
                //MessageBox.Show(mensaje);

            }



        }

        static void EnviarExcelFacturacionSinEnviar()
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
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", mensaje, EventLogEntryType.Error);

                //poner que guarde un mensaje en el registro
                //MessageBox.Show(mensaje);

            }





        }


        // genero los archivos para enviar en el dia de la fecha, los datos se generan con la fecha de ayer.
        static void GenerarExcelInventario()
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
                string SelectString = "select (select codigo_distribuidor from BDF_Configuracion where codigo = 1) IdDistribuidor, (select num.numero from numeracion as num  where num.sucursal=1 and num.tiponum='BD' and num.letra='X' and num.empresa_id=1) IdPaquete,  stockhistorico.fecha fecha , scj_sinonimos.CodigoSinonimo IdProducto, case when sum(stocktot) > 0 then round(sum(cast(cast(stocktot as numeric(18,2))/cast(uxbcompra as numeric(18,2)) as numeric(18,2))),2) else 0 end cantidad  from stockhistorico, articulo, scj_sinonimos where stockhistorico.Articulo_Codigo= articulo.codigo and scj_sinonimos.CodigoArt=articulo.codigo and   stockhistorico.fecha = Convert(varchar,DATEADD(d,-1,getdate()),103) and deposito_id = 0 and articulo.prove =" + codigo_proveedor + " and scj_sinonimos.CodigoSinonimo <> '' and articulo.baja = 0 group by scj_sinonimos.CodigoSinonimo, stockhistorico.fecha ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", SelectString, EventLogEntryType.Information);





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


                HojaExcel2.Range["A3:A3"].Value = "TotalUnidades";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!e2:datos!e99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;
                //guardo el documento


                string pathfull = path + usuario + "_inv_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd") + ".xlsx";
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", pathfull, EventLogEntryType.Information);
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

        static void GenerarExcelFacturacion()
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
                string SelectString = "select (select codigo_distribuidor from BDF_Configuracion where codigo = 1) idDistribuidor , (select num.numero from numeracion as num  where num.sucursal=1 and num.tiponum='BD' and num.letra='X' and num.empresa_id=1) IdPaquete, v.fecha fecha , v.comprobante NroComprobante, '' IdPedidoDinesys, case left(referencia,2) when 'NC' THEN referencia END NroComprobanteAsociado,  cod_cli IdCliente, Jonhson_Tipos_Cliente.Codigo IdTipoDeCliente,  cod_ven IdVendedor, scj.codigoSinonimo IdProducto , round(sum(iv.Cantidad*art.uxb/art.uxbcompra),2) Cantidad, case WHEN  sum(iv.Cantidad*art.uxb/art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('FA', 'PE')  THEN 'FC' when sum(iv.Cantidad* art.uxb/ art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('ND') then 'ND' else 'NC' END Tipo,     motivo.descripcion MotivoNC  from numeracion, venta v inner join  item_venta iv on v.comprobante=iv.comprobante  and v.fecha = iv.fecha and v.empresa_id= iv.empresa_id left join motivo on v.motivodev = motivo.codigo  left join articulo art on iv.cod_art=art.codigo inner join scj_sinonimos scj on scj.CodigoArt=art.codigo inner join cliente cli on cli.codigo=v.cod_cli inner join ramo on cli.ramo=ramo.codigo inner join Ramo_Vs_RamoBDF on Ramo_Vs_RamoBDF.codigo=ramo.codigo inner join Jonhson_Tipos_Cliente on Jonhson_Tipos_Cliente.Id=Ramo_Vs_RamoBDF.IDBDF where TipoNum = 'BD' and v.fecha = Convert(varchar,DATEADD(d,-1,getdate()),103) and left(v.comprobante, 2) in ('FA','NC','ND','PE')  and art.prove =" + codigo_proveedor + " and scj.CodigoSinonimo <> '' and art.baja = 0 and fecha_factura is not null   group by v.fecha, v.comprobante, v.referencia, v.cod_cli, v.cod_ven, scj.CodigoSinonimo, motivo.descripcion, Jonhson_Tipos_Cliente.Codigo having sum(iv.Cantidad*art.uxb/art.uxbcompra) <> 0 ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                sqlcommand.CommandTimeout = 3600;
                SqlDataAdapter Adaptador = new SqlDataAdapter(SelectString, Conection);
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", SelectString, EventLogEntryType.Information);


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

                HojaExcel2.Range["A3:A3"].Value = "TotalUnidades";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!K2:datos!K99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;
                //guardo el documento


                string pathfull = path + usuario + "_fac_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd") + ".xlsx";
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", pathfull, EventLogEntryType.Information);
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

       static void EnviarExcelInventario()
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
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", mensaje, EventLogEntryType.Error);

                //poner que guarde un mensaje en el registro
                //MessageBox.Show(mensaje);

            }



        }

        static void EnviarExcelFacturacion()
        {

            try
            {
                string fileToUpload = path + usuario + "_vta_" + DateTime.Today.AddDays(-1).ToString("yyyyMMdd") + ".xlsx";
                string url = url_ventas;
                

                using (var client = new WebClient())
                {
                    client.Headers.Add(HttpRequestHeader.Authorization,  "Basic " + Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(usuario + ":" + contraseña)));
                    byte[] result = client.UploadFile(url, "POST",fileToUpload);
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
                System.Diagnostics.EventLog.WriteEntry("Servicio BDF", mensaje, EventLogEntryType.Error);

                //poner que gaurde un mensaje en el
                //MessageBox.Show(mensaje);

            }


        }


        // genero archivos que no hayan sido generados por algun motivo
        static void GenerarExcelClientes(int vDias)
        {
            try
            {

                if (File.Exists(path + usuario + "_mc_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx.ant") == true) return;

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                Conection.Open();

                // Utilizar una variable para almacenar la instrucción SQL.
                string SelectString = "select (select codigo_distribuidor from BDF_Configuracion where codigo = 1) IdDistribuidor, (select num.numero+1 from numeracion as num  where num.sucursal=1 and num.tiponum='BD' and num.letra='X' and 	num.empresa_id=1) IdPaquete, 	cli.codigo,cli.razon_social,'' cadena, 6 provincia,cli.localidad,cli.cp, cli.direccion,'' altura,	cli.cuit,case when geoy=0 then '' else convert(varchar,cli.geoy) end geoy,case when geox=0 then '' else convert(varchar,cli.geox) end geox,cli.ramo 	from cliente cli where 	cli.codigo>0 AND cli.estado='A'";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", SelectString, EventLogEntryType.Information);





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
                LibroExcel.Worksheets.Add();
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
                HojaExcel.Range["C1:C1"].Value = "IdCliente";
                HojaExcel.Range["D1:D1"].Value = "RazonSocial";
                HojaExcel.Range["E1:E1"].Value = "BannerText";
                HojaExcel.Range["F1:F1"].Value = "IdProvincia";
                HojaExcel.Range["G1:G1"].Value = "Localidad";
                HojaExcel.Range["H1:H1"].Value = "CodigoPostal";
                HojaExcel.Range["I1:I1"].Value = "Calle";
                HojaExcel.Range["J1:J1"].Value = "Numero";
                HojaExcel.Range["K1:K1"].Value = "CUIT";
                HojaExcel.Range["L1:L1"].Value = "Latitud";
                HojaExcel.Range["M1:M1"].Value = "Longitud";
                HojaExcel.Range["N1:N1"].Value = "IdTipoCliente";

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
                    HojaExcel.Cells[i, "F"] = Row.ItemArray[5];
                    HojaExcel.Cells[i, "G"] = Row.ItemArray[6];
                    HojaExcel.Cells[i, "H"] = Row.ItemArray[7];
                    HojaExcel.Cells[i, "I"] = Row.ItemArray[8];
                    HojaExcel.Cells[i, "J"] = Row.ItemArray[9];
                    HojaExcel.Cells[i, "K"] = Row.ItemArray[10];
                    HojaExcel.Cells[i, "L"] = Row.ItemArray[11];
                    HojaExcel.Cells[i, "M"] = Row.ItemArray[12];
                    HojaExcel.Cells[i, "N"] = Row.ItemArray[13];

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
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!C2:datos!C99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;


                string pathfull = path + usuario + "_mc_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx";
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", pathfull, EventLogEntryType.Information);
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

        // genero archivos que no hayan sido generados por algun motivo
        static void GenerarExcelInventario(int vDias)
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
                string SelectString = "select (select codigo_distribuidor from BDF_Configuracion where codigo = 1) IdDistribuidor, (select num.numero+1 from numeracion as num  where num.sucursal=1 and num.tiponum='BD' and num.letra='X' and num.empresa_id=1) IdPaquete,   articulo.codigoprov IdProducto, 'PC' UnidadMedida,convert(varchar,stockhistorico.fecha,23) fecha , 1 tipoinventario,'UNI'Deposito, case when sum(stocktot) > 0 then round(sum(cast(cast(stocktot as numeric(18,2))/cast(uxbcompra as numeric(18,2)) as numeric(18,2))),0) else 0 end cantidad from stockhistorico, articulo  where stockhistorico.Articulo_Codigo= articulo.codigo and  stockhistorico.fecha = Convert(varchar,DATEADD(d," + vDias.ToString() + ",getdate()),103) and deposito_id = 0 and articulo.prove =" + codigo_proveedor + " and articulo.baja = 0 group by articulo.codigoprov,stockhistorico.fecha ";
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", SelectString, EventLogEntryType.Information);





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

                LibroExcel.Worksheets.Add();
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
                HojaExcel.Range["C1:D1"].Value = "IdProducto";
                HojaExcel.Range["D1:D1"].Value = "UnidadMedida";
                HojaExcel.Range["E1:E1"].Value = "Fecha";
                HojaExcel.Range["F1:F1"].Value = "IdTipoInventario";
                HojaExcel.Range["G1:G1"].Value = "Deposito";
                HojaExcel.Range["H1:H1"].Value = "Cantidad";
 

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
                    
                    HojaExcel.Cells[i, "C"] = Row.ItemArray[2];
                    
                    HojaExcel.Cells[i, "D"] = Row.ItemArray[3];
                    
                    HojaExcel.Cells[i, "E"] = "'" + Row.ItemArray[4];
                    HojaExcel.Cells[i, "F"] = Row.ItemArray[5];
                    HojaExcel.Cells[i, "G"] = Row.ItemArray[6];
                    HojaExcel.Cells[i, "H"] = Row.ItemArray[7];

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
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!H2:datos!H99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;


                HojaExcel2.Range["A3:A3"].Value = "TotalUnidades";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!H2:datos!H99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;


                string pathfull = path + usuario + "_inv_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx";
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", pathfull, EventLogEntryType.Information);
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

       static  void GenerarExcelFacturacion(int vDias)
        {
            try
            {

                if (File.Exists(path + usuario + "_vta_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx.ant") == true) return;

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);








                // Abrir la conexión.
                Conection.Open();

                //comentario para compilar a asdsadsa
                string SelectString = " select (select codigo_distribuidor from BDF_Configuracion where codigo = 1) idDistribuidor , ";
                SelectString = SelectString + $" (select num.numero+1 from numeracion as num  where num.sucursal=1 and num.tiponum='BD' and num.letra='X' and num.empresa_id=1) IdPaquete, ";
                SelectString = SelectString + $" cod_cli IdCliente, isnull(bdf_Tipos_Cliente.Codigo,8) IdTipoDeCliente,  cod_ven IdVendedor, 'Nombre Vendedor' NombrVen,'Apellido Vendedor' Apellidoven,";
                SelectString = SelectString + $" art.codigoprov IdProducto , 'PC' UnidadMedida,  convert(varchar,v.fecha,23) fecha ,";
                SelectString = SelectString + $" case WHEN  sum(iv.Cantidad*art.uxb/art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('FR','FA', 'PE') ";
                SelectString = SelectString + $" THEN 'OR' when sum(iv.Cantidad* art.uxb/ art.uxbcompra) >= 0 and left(v.comprobante, 2) in ('ND') then 'DE' else 'CR' END Tipo,   round(sum(iv.Cantidad_uni),0) Cantidad, ";
                SelectString = SelectString + $" v.comprobante NroComprobante, isnull(case left(v.comprobante,2) when 'NC' THEN referencia  else '' END,'') NroComprobanteAsociado,  motivo.descripcion MotivoNC  ";
                SelectString = SelectString + $" from numeracion, venta v inner join  item_venta iv on v.comprobante=iv.comprobante  and ";
                SelectString = SelectString + $" v.fecha = iv.fecha and v.empresa_id= iv.empresa_id left join motivo on v.motivodev = motivo.codigo  left join articulo art on ";
                SelectString = SelectString + $" iv.cod_art=art.codigo inner join cliente cli on cli.codigo=v.cod_cli inner join ramo ";
                SelectString = SelectString + $" on cli.ramo=ramo.codigo LEFT join Ramo_Vs_RamoBDF on Ramo_Vs_RamoBDF.codigo=ramo.codigo left join bdf_Tipos_Cliente on bdf_Tipos_Cliente.Id=Ramo_Vs_RamoBDF.IDBDF";
                SelectString = SelectString + $" where TipoNum = 'BD' and v.fecha >= Convert(varchar,DATEADD(d," + vDias.ToString() + ",getdate()),103) and left(v.comprobante, 2) in ('FA','NC','ND','PE') and art.prove =" + codigo_proveedor + "  and art.baja = 0 and fecha_factura is not null   ";
                SelectString = SelectString + $" group by v.fecha, v.comprobante, v.referencia, v.cod_cli, v.cod_ven, art.codigoprov, motivo.descripcion, bdf_Tipos_Cliente.Codigo having sum(iv.Cantidad_uni) <> 0 "; Debug.Write(SelectString);
                SqlCommand sqlcommand = new SqlCommand(SelectString, Conection);
                sqlcommand.CommandTimeout = 3600;
                SqlDataAdapter Adaptador = new SqlDataAdapter(SelectString, Conection);
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", SelectString, EventLogEntryType.Information);


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
                LibroExcel.Worksheets.Add();
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
                HojaExcel.Range["C1:C1"].Value = "IdCliente";
                HojaExcel.Range["D1:D1"].Value = "IdTipoDeCliente";
                HojaExcel.Range["E1:E1"].Value = "IdVendedor";
                HojaExcel.Range["F1:F1"].Value = "NombreVendedor";
                HojaExcel.Range["G1:G1"].Value = "ApellidoVendedor";
                HojaExcel.Range["H1:H1"].Value = "IdProducto";
                HojaExcel.Range["I1:I1"].Value = "UnidadMedida";
                HojaExcel.Range["J1:J1"].Value = "Fecha";
                HojaExcel.Range["K1:K1"].Value = "TipoDocumento";
                HojaExcel.Range["L1:L1"].Value = "Cantidad";
                HojaExcel.Range["M1:M1"].Value = "NroComprobante";
                HojaExcel.Range["N1:N1"].Value = "NroComprobanteAsociado";
                HojaExcel.Range["O1:O1"].Value = "MotivoCR";

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
                    HojaExcel.Cells[i, "J"] = "'"  + Row.ItemArray[9];
                    HojaExcel.Cells[i, "K"] = Row.ItemArray[10];
                    HojaExcel.Cells[i, "L"] = Row.ItemArray[11];
                    HojaExcel.Cells[i, "M"] = Row.ItemArray[12];
                    HojaExcel.Cells[i, "N"] = Row.ItemArray[13];
                    HojaExcel.Cells[i, "O"] = Row.ItemArray[14];

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

                //LibroExcel.Worksheets.Add();
                HojaExcel2 = (Microsoft.Office.Interop.Excel.Worksheet)LibroExcel.Worksheets[2];
                //pruebo de cambiar el nombre de la hora
                HojaExcel2.Name = "verificacion";
                HojaExcel2.Activate();
                HojaExcel2.Range["A1:A1"].Value = "INDICADOR";
                HojaExcel2.Range["B1:B1"].Value = "VALOR";


                HojaExcel2.Range["A2:A2"].Value = "CantRegistros";
                //cuento la cantidad de registros
                HojaExcel2.Range["B2"].Formula = "=CONTAR(datos!L2:datos!L99999)";
                int contar = Convert.ToInt32(HojaExcel2.Range["B2"].Value);
                HojaExcel2.Range["B2"].Value = contar;

                HojaExcel2.Range["A3:A3"].Value = "TotalUnidades";
                //sumo la columna cantidad
                HojaExcel2.Range["B3"].Formula = "=SUMA(datos!L2:datos!L99999)";
                decimal sumar = Convert.ToDecimal(HojaExcel2.Range["B3"].Value);
                sumar = decimal.Round(sumar, 2);
                HojaExcel2.Range["B3"].Value = sumar;
                //guardo el documento



                string pathfull = path + usuario + "_vta_" + DateTime.Today.AddDays(vDias).ToString("yyyyMMdd") + ".xlsx";
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", pathfull, EventLogEntryType.Information);
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


      static  void LeerDatosConfiguracion()
        {
            try
            {

                //guardo el connection string
                string cadena;
                cadena = "Server=" + serverSQL + " ;Database=" + baseSQL + ";User Id=" + usuarioSQL + " ;Password=" + contraseñaSQL + ";";

                //comento porque ya no necesito esto.
                //System.Diagnostics.EventLog.WriteEntry("Servicio BDF", "cadena de conexion:" + cadena + " Fecha:" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Information);
                // Crear un objeto SqlConnection, y luego pasar la ConnectionString al constructor.            
                SqlConnection Conection = new SqlConnection(cadena);


                string SelectConfiguracion = "select * from BDF_Configuracion where codigo = 1";
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
                    System.Diagnostics.EventLog.WriteEntry("Servicio BDF", "El distribuidor no tiene seteado los datos de configuracion" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"), EventLogEntryType.Error);
                }
                reader.Close();
                Conection.Close();


            }
            catch (Exception ex)
            {

                ErrorHandler.ManejarError("Ocurrio un error leyendo los datos de configuracion del distribuidor", ErrorType.Critical, ex);
            }


        }


        static void CrearConfiguracionXML()
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

    }
}
