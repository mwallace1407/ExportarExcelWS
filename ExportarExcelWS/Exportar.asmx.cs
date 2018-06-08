using System;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Web.Services;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Data;

namespace ExportarExcelWS
{
    /// <summary>
    /// Summary description for Service1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
    // [System.Web.Script.Services.ScriptService]
    public class Exportar : System.Web.Services.WebService
    {
        #region Abanks

        #region Polizas_Abanks

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void ReportePolizas(int TipoDocumento, string UserName, DateTime FechaIni, DateTime FechaFin, int NumeroMovimiento, string Cuenta, string DescripcionCuenta, string DescripcionEncabezado, int Moneda, bool BusquedaEstricta, string RutaArchivos, string Archivo, int RegistrosPorHoja)
        {
            try
            {
                int? NumeroMovimientoN = null;
                int? MonedaN = null;

                if (NumeroMovimiento > 0)
                    NumeroMovimientoN = NumeroMovimiento;

                if (Moneda > 0)
                    MonedaN = Moneda;

                if (string.IsNullOrWhiteSpace(Cuenta))
                    Cuenta = null;

                if (string.IsNullOrWhiteSpace(DescripcionCuenta))
                    DescripcionCuenta = null;

                if (string.IsNullOrWhiteSpace(DescripcionEncabezado))
                    DescripcionEncabezado = null;

                //FechaFin = FechaFin.AddDays(1).AddMilliseconds(-1);

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLReportePolizas(TipoDocumento, UserName, FechaIni.Year, FechaIni, FechaFin, NumeroMovimientoN, Cuenta, DescripcionCuenta, DescripcionEncabezado, MonedaN, BusquedaEstricta, RutaArchivos, Archivo, RegistrosPorHoja);
            }
            catch (Exception ex)
            {
                RegistraLog("DLReportePolizas", ex.Message, ex.StackTrace);
            }
        }

        private void DLReportePolizas(int TipoDocumento, string UserName, int Anno, DateTime FechaIni, DateTime FechaFin, int? NumeroMovimiento, string Cuenta, string DescripcionCuenta, string DescripcionEncabezado, int? Moneda, bool BusquedaEstricta, string RutaArchivos, string Archivo, int RegistrosPorHoja)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Historico");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_AbanksPolizas");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Anno", DbType.Int32, Anno);
                db.AddInParameter(selectCommand, "@BusquedaEstricta", DbType.Boolean, BusquedaEstricta);
                db.AddInParameter(selectCommand, "@Numero_Mov", DbType.Int32, NumeroMovimiento);
                db.AddInParameter(selectCommand, "@Fecha_MovIni", DbType.DateTime, FechaIni);
                db.AddInParameter(selectCommand, "@Fecha_MovFin", DbType.DateTime, FechaFin);
                db.AddInParameter(selectCommand, "@Descripcion_Encabezado", DbType.String, DescripcionEncabezado);
                db.AddInParameter(selectCommand, "@Cuenta", DbType.String, Cuenta);
                db.AddInParameter(selectCommand, "@Desc_Cuenta", DbType.String, DescripcionCuenta);
                db.AddInParameter(selectCommand, "@Cod_Moneda_Original", DbType.Int32, Moneda);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLReportePolizas", ex.Message, ex.StackTrace);
            }
        }

        #endregion Polizas_Abanks

        #endregion Abanks

        #region Aplicaciones

        #region Relacion_ServidorApplicaciones

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void RelSrvApp(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Srv_Id)
        {
            try
            {
                if (Srv_Id.Trim() == "")
                    Srv_Id = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLRelSrvApp(TipoDocumento, UserName, RutaArchivos, Archivo, RegistrosPorHoja, Srv_Id);
            }
            catch (Exception ex)
            {
                RegistraLog("RelSrvApp", ex.Message, ex.StackTrace);
            }
        }

        private void DLRelSrvApp(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Srv_Id)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_RelSrv");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Srv_Id", DbType.String, Srv_Id);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLRelSrvApp", ex.Message, ex.StackTrace);
            }
        }

        #endregion Relacion_ServidorApplicaciones

        #region Relacion_AplicacionesBD

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void RelBDApp(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string AppBD_Id)
        {
            try
            {
                if (AppBD_Id.Trim() == "")
                    AppBD_Id = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLRelBDApp(TipoDocumento, UserName, RutaArchivos, Archivo, RegistrosPorHoja, AppBD_Id);
            }
            catch (Exception ex)
            {
                RegistraLog("RelBDApp", ex.Message, ex.StackTrace);
            }
        }

        private void DLRelBDApp(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string AppBD_Id)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_RelBD");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@AppBD_Id", DbType.String, AppBD_Id);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLRelBDApp", ex.Message, ex.StackTrace);
            }
        }

        #endregion Relacion_AplicacionesBD

        #region Relacion_BDServidor

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void RelSrvBD(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Srv_Id)
        {
            try
            {
                if (Srv_Id.Trim() == "")
                    Srv_Id = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLRelSrvBD(TipoDocumento, UserName, RutaArchivos, Archivo, RegistrosPorHoja, Srv_Id);
            }
            catch (Exception ex)
            {
                RegistraLog("RelSrvBD", ex.Message, ex.StackTrace);
            }
        }

        private void DLRelSrvBD(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Srv_Id)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_RelSrvBD");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Srv_Id", DbType.String, Srv_Id);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLRelSrvBD", ex.Message, ex.StackTrace);
            }
        }

        #endregion Relacion_BDServidor

        #region Relacion_DiscosServidor

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void DiscosSrv(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Srv_Id)
        {
            try
            {
                if (Srv_Id.Trim() == "")
                    Srv_Id = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLDiscosSrv(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Srv_Id);
            }
            catch (Exception ex)
            {
                RegistraLog("DiscosSrv", ex.Message, ex.StackTrace);
            }
        }

        private void DLDiscosSrv(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Srv_Id)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_DiscosSrv");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Srv_Id", DbType.String, Srv_Id);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLDiscosSrv", ex.Message, ex.StackTrace);
            }
        }

        #endregion Relacion_DiscosServidor

        #region GeneralAplicaciones

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void GeneralAplicaciones(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string AppSt_Id, string AppT_Id, string App_EnTFS, string App_Productiva)
        {
            try
            {
                bool? App_EnTFSb = null;
                bool? App_Productivab = null;

                App_EnTFS = App_EnTFS.Trim();
                App_Productiva = App_Productiva.Trim();

                if (AppSt_Id.Trim() == "")
                    AppSt_Id = null;

                if (AppT_Id.Trim() == "")
                    AppT_Id = null;

                if (App_EnTFS == "NO")
                    App_EnTFSb = false;
                else if (App_EnTFS == "SI")
                    App_EnTFSb = true;

                if (App_Productiva == "NO")
                    App_Productivab = false;
                else if (App_Productiva == "SI")
                    App_Productivab = true;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLGeneralAplicaciones(UserName, RutaArchivos, Archivo, RegistrosPorHoja, AppSt_Id, AppT_Id, App_EnTFSb, App_Productivab);
            }
            catch (Exception ex)
            {
                RegistraLog("GeneralAplicaciones", ex.Message, ex.StackTrace);
            }
        }

        private void DLGeneralAplicaciones(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string AppSt_Id, string AppT_Id, bool? App_EnTFS, bool? App_Productiva)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_GeneralAplicaciones");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@AppSt_Id", DbType.String, AppSt_Id);
                db.AddInParameter(selectCommand, "@AppT_Id", DbType.String, AppT_Id);
                db.AddInParameter(selectCommand, "@App_EnTFS", DbType.Boolean, App_EnTFS);
                db.AddInParameter(selectCommand, "@App_Productiva", DbType.Boolean, App_Productiva);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLGeneralAplicaciones", ex.Message, ex.StackTrace);
            }
        }

        #endregion GeneralAplicaciones

        #region GeneralServidores

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void GeneralServidores(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string idSO, string idEquipo, string Srv_Tipo, string Srv_EsVirtual, string Srv_Estado)
        {
            try
            {
                bool? Srv_EsVirtualb = null;
                bool? Srv_Estadob = null;

                Srv_EsVirtual = Srv_EsVirtual.Trim();
                Srv_Estado = Srv_Estado.Trim();

                if (idSO.Trim() == "")
                    idSO = null;

                if (idEquipo.Trim() == "")
                    idEquipo = null;

                if (Srv_Tipo.Trim() == "")
                    Srv_Tipo = null;

                if (Srv_EsVirtual == "NO")
                    Srv_EsVirtualb = false;
                else if (Srv_EsVirtual == "SI")
                    Srv_EsVirtualb = true;

                if (Srv_Estado == "NO")
                    Srv_Estadob = false;
                else if (Srv_Estado == "SI")
                    Srv_Estadob = true;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLGeneralServidores(UserName, RutaArchivos, Archivo, RegistrosPorHoja, idSO, idEquipo, Srv_Tipo, Srv_EsVirtualb, Srv_Estadob);
            }
            catch (Exception ex)
            {
                RegistraLog("GeneralServidores", ex.Message, ex.StackTrace);
            }
        }

        private void DLGeneralServidores(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string idSO, string idEquipo, string Srv_Tipo, bool? Srv_EsVirtual, bool? Srv_Estado)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_GeneralServidores");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@idSistemaOperativo", DbType.String, idSO);
                db.AddInParameter(selectCommand, "@idEquipo", DbType.String, idEquipo);
                db.AddInParameter(selectCommand, "@Srv_Tipo", DbType.String, Srv_Tipo);
                db.AddInParameter(selectCommand, "@Srv_EsVirtual", DbType.Boolean, Srv_EsVirtual);
                db.AddInParameter(selectCommand, "@Srv_Estado", DbType.Boolean, Srv_Estado);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLGeneralServidores", ex.Message, ex.StackTrace);
            }
        }

        #endregion GeneralServidores

        #endregion Aplicaciones

        #region Hardware

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void InventarioEquipos(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string idTipoEquipo, string idMarca, string idUbicacion, string idUsuario, string Responsiva, string Modelo, string NoSerie, string FechaMovimientoIni, string FechaMovimientoFin, string idEstado)
        {
            try
            {
                DateTime? FechaIni;
                DateTime? FechaFin;

                if (idTipoEquipo.Trim() == "")
                    idTipoEquipo = null;

                if (idMarca.Trim() == "")
                    idMarca = null;

                if (idUbicacion.Trim() == "")
                    idUbicacion = null;

                if (idUsuario.Trim() == "")
                    idUsuario = null;

                if (Responsiva.Trim() == "")
                    Responsiva = null;

                if (Modelo.Trim() == "")
                    Modelo = null;

                if (NoSerie.Trim() == "")
                    NoSerie = null;

                if (idEstado.Trim() == "")
                    idEstado = null;

                if (FechaMovimientoIni.Trim() == "" || FechaMovimientoIni.Replace(" ", "") == "//")
                    FechaIni = null;
                else
                    FechaIni = DateTime.ParseExact(FechaMovimientoIni, "dd/MM/yyyy", System.Threading.Thread.CurrentThread.CurrentCulture);

                if (FechaMovimientoFin.Trim() == "" || FechaMovimientoFin.Replace(" ", "") == "//")
                    FechaFin = null;
                else
                    FechaFin = DateTime.ParseExact(FechaMovimientoFin + " 23:59:59", "dd/MM/yyyy HH:mm:ss", System.Threading.Thread.CurrentThread.CurrentCulture);

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLInventarioEquipos(UserName, RutaArchivos, Archivo, RegistrosPorHoja, idTipoEquipo, idMarca, idUbicacion, idUsuario, Responsiva, Modelo, NoSerie, FechaIni, FechaFin, idEstado);
            }
            catch (Exception ex)
            {
                RegistraLog("InventarioEquipos", ex.Message, ex.StackTrace);
            }
        }

        private void DLInventarioEquipos(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string idTipoEquipo, string idMarca, string idUbicacion, string idUsuario, string Responsiva, string Modelo, string NoSerie, DateTime? FechaMovimientoIni, DateTime? FechaMovimientoFin, string idEstado)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_InventarioEquipos");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@idTipoEquipo", DbType.String, idTipoEquipo);
                db.AddInParameter(selectCommand, "@idMarca", DbType.String, idMarca);
                db.AddInParameter(selectCommand, "@idUbicacion", DbType.String, idUbicacion);
                db.AddInParameter(selectCommand, "@idUsuario", DbType.String, idUsuario);
                db.AddInParameter(selectCommand, "@Responsiva", DbType.String, Responsiva);
                db.AddInParameter(selectCommand, "@Modelo", DbType.String, Modelo);
                db.AddInParameter(selectCommand, "@NoSerie", DbType.String, NoSerie);
                db.AddInParameter(selectCommand, "@FechaMovimientoIni", DbType.DateTime, FechaMovimientoIni);
                db.AddInParameter(selectCommand, "@FechaMovimientoFin", DbType.DateTime, FechaMovimientoFin);
                db.AddInParameter(selectCommand, "@idEstado", DbType.String, idEstado);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLInventarioEquipos", ex.Message, ex.StackTrace);
            }
        }

        #endregion Hardware



        #region ReportesDinamicos

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void EjecutarRD(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, DataTable Parametros, int RD_Id)
        {
            try
            {
                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLReporteDinamico(UserName, RutaArchivos, Archivo, RegistrosPorHoja, ConvertirDataTableAMatrizEscalonada(Parametros), RD_Id);
            }
            catch (Exception ex)
            {
                RegistraLog("EjecutarRD", ex.Message, ex.StackTrace);
            }
        }

        private string[][] ConvertirDataTableAMatrizEscalonada(DataTable Parametros)
        {
            string[][] Params = new string[Parametros.Rows.Count][];

            try
            {
                for (int w = 0; w < Parametros.Rows.Count; w++)
                    Params[w] = new string[] { Parametros.Rows[w]["Nombre"].ToString(), Parametros.Rows[w]["Valor"].ToString(), Parametros.Rows[w]["Tipo"].ToString(), Parametros.Rows[w]["Longitud"].ToString() };
            }
            catch (Exception ex)
            {
                RegistraLog("ConvertirDTaME", ex.Message, ex.StackTrace);
            }

            return Params;
        }

        private SqlDbType Equivalencia_SQL_SQLDBType(string SQLDataType)
        {
            SqlDbType Tipo;

            Tipo = SqlDbType.Variant;

            try
            {
                switch (SQLDataType)
                {
                    case "bigint":
                        Tipo = SqlDbType.BigInt;
                        break;

                    case "binary":
                        Tipo = SqlDbType.Binary;
                        break;

                    case "bit":
                        Tipo = SqlDbType.Bit;
                        break;

                    case "char":
                        Tipo = SqlDbType.Char;
                        break;

                    case "date":
                        Tipo = SqlDbType.Date;
                        break;

                    case "datetime":
                        Tipo = SqlDbType.DateTime;
                        break;

                    case "datetime2":
                        Tipo = SqlDbType.DateTime2;
                        break;

                    case "datetimeoffset":
                        Tipo = SqlDbType.DateTimeOffset;
                        break;

                    case "decimal":
                        Tipo = SqlDbType.Decimal;
                        break;

                    case "float":
                        Tipo = SqlDbType.Float;
                        break;

                    case "image":
                        Tipo = SqlDbType.Image;
                        break;

                    case "int":
                        Tipo = SqlDbType.Int;
                        break;

                    case "money":
                        Tipo = SqlDbType.Money;
                        break;

                    case "nchar":
                        Tipo = SqlDbType.NChar;
                        break;

                    case "ntext":
                        Tipo = SqlDbType.NText;
                        break;

                    case "numeric":
                        Tipo = SqlDbType.Decimal;
                        break;

                    case "nvarchar":
                        Tipo = SqlDbType.NVarChar;
                        break;

                    case "real":
                        Tipo = SqlDbType.Real;
                        break;

                    case "rowversion":
                        Tipo = SqlDbType.Timestamp;
                        break;

                    case "smalldatetime":
                        Tipo = SqlDbType.SmallDateTime;
                        break;

                    case "smallint":
                        Tipo = SqlDbType.SmallInt;
                        break;

                    case "smallmoney":
                        Tipo = SqlDbType.SmallMoney;
                        break;

                    case "sql_variant":
                        Tipo = SqlDbType.Variant;
                        break;

                    case "text":
                        Tipo = SqlDbType.Text;
                        break;

                    case "time":
                        Tipo = SqlDbType.Time;
                        break;

                    case "timestamp":
                        Tipo = SqlDbType.Timestamp;
                        break;

                    case "tinyint":
                        Tipo = SqlDbType.TinyInt;
                        break;

                    case "uniqueidentifier":
                        Tipo = SqlDbType.UniqueIdentifier;
                        break;

                    case "varbinary":
                        Tipo = SqlDbType.VarBinary;
                        break;

                    case "varchar":
                        Tipo = SqlDbType.VarChar;
                        break;

                    case "xml":
                        Tipo = SqlDbType.Xml;
                        break;

                    default:
                        Tipo = SqlDbType.Variant;
                        break;
                }
            }
            catch (Exception ex)
            {
                RegistraLog("Equivalencia_SQL_SQLDBType", ex.Message, ex.StackTrace);
            }

            return Tipo;
        }

        private DateTime ObtieneFecha(string Fecha)
        {
            DateTime f;

            if (DateTime.TryParseExact(Fecha, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out f))
                return f;
            else
                return new DateTime(1900, 1, 1);
        }

        private byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];

            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);

            return bytes;
        }

        private SqlParameter CreaParametroRD(string Nombre_Parametro, object Valor, SqlDbType Tipo_Dato, int Tamanno = 0)
        {
            SqlParameter param = new SqlParameter();

            try
            {
                if (Tamanno > 0 && (Tipo_Dato == SqlDbType.NChar || Tipo_Dato == SqlDbType.NVarChar || Tipo_Dato == SqlDbType.Char || Tipo_Dato == SqlDbType.VarChar || Tipo_Dato == SqlDbType.NText || Tipo_Dato == SqlDbType.Text))
                {
                    string ValorAlt;
                    param = new SqlParameter(Nombre_Parametro, Tipo_Dato, Tamanno);

                    ValorAlt = Valor.ToString();

                    if (ValorAlt.Length > Tamanno)
                        ValorAlt = ValorAlt.Substring(0, Tamanno);

                    param.Value = Valor;
                }
                else if (Tipo_Dato == SqlDbType.Bit)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        if (Valor.ToString().ToLowerInvariant() == "true" || Valor.ToString() == "1")
                            Valor = true;
                        else
                            Valor = false;
                    }

                    param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                    param.Value = Valor;
                }
                else if (Tipo_Dato == SqlDbType.Date || Tipo_Dato == SqlDbType.DateTime || Tipo_Dato == SqlDbType.DateTime2 || Tipo_Dato == SqlDbType.SmallDateTime)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = ObtieneFecha(Valor.ToString());
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.BigInt)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Int64 Dato = 0;

                        Int64.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.Decimal || Tipo_Dato == SqlDbType.Money || Tipo_Dato == SqlDbType.SmallMoney)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Decimal Dato = 0;

                        Decimal.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.Float)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Double Dato = 0;

                        Double.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.Int)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Int32 Dato = 0;

                        Int32.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.Real)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Single Dato = 0;

                        Single.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.SmallInt)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Int16 Dato = 0;

                        Int16.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.TinyInt)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Byte Dato = 0;

                        Byte.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.UniqueIdentifier)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        Guid Dato = new Guid();

                        Guid.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.DateTimeOffset)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        DateTimeOffset Dato = new DateTimeOffset();

                        DateTimeOffset.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.Time)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        TimeSpan Dato = new TimeSpan();

                        TimeSpan.TryParse(Valor.ToString(), out Dato);
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = Dato;
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else if (Tipo_Dato == SqlDbType.Binary || Tipo_Dato == SqlDbType.Image || Tipo_Dato == SqlDbType.Timestamp || Tipo_Dato == SqlDbType.VarBinary)
                {
                    if (Valor != System.DBNull.Value)
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = GetBytes(Valor.ToString());
                    }
                    else
                    {
                        param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                        param.Value = System.DBNull.Value;
                    }
                }
                else// if (Tipo_Dato != SqlDbType.NChar && Tipo_Dato != SqlDbType.NVarChar && Tipo_Dato != SqlDbType.Char && Tipo_Dato != SqlDbType.VarChar)
                {
                    param = new SqlParameter(Nombre_Parametro, Tipo_Dato);
                    param.Value = Valor;
                }
            }
            catch (Exception ex)
            {
                RegistraLog("CreaParametroRD", ex.Message, ex.StackTrace);
            }

            return param;
        }

        private const string RDNulo = "~CampoNulo~";
        private const int RDNombre = 0;
        private const int RDValor = 1;
        private const int RDTipo = 2;
        private const int RDTamanno = 3;
        private const int RDTotalParam = 4;

        private enum TiposScriptRD
        {
            Texto = 1,
            StoredProcedure = 2,
            SSIS = 3
        }

        private struct DetalleReporte
        {
            public string Conexion;
            public string Script;
            public TiposScriptRD Tipo;

            public DetalleReporte(int RD_Id)
            {
                DataTable Tabla = new DataTable();
                Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
                DbCommand selectCommand = null;

                Conexion = "";
                Script = "";
                Tipo = TiposScriptRD.Texto;

                try
                {
                    selectCommand = db.GetSqlStringCommand("stpS_RptDinamicosDetalle");
                    selectCommand.CommandType = CommandType.StoredProcedure;

                    db.AddInParameter(selectCommand, "@RD_Id", DbType.Int32, RD_Id);
                    Tabla.Load(db.ExecuteReader(selectCommand));

                    if (Tabla.Rows.Count > 0)
                    {
                        Conexion = Tabla.Rows[0]["Conexion"].ToString();
                        Script = Tabla.Rows[0]["Script"].ToString();
                        Tipo = (TiposScriptRD)Convert.ToInt32(Tabla.Rows[0]["TipoR"].ToString());
                    }
                    else
                    {
                        Conexion = "ErrorDetalleReporte: No se obtuvo el detalle del reporte.";
                    }
                }
                catch (Exception ex)
                {
                    Conexion = "ErrorDetalleReporte: " + ex.Message;
                    Script = ex.StackTrace;
                }
            }
        }

        private SqlParameter[] CreaParametrosRD(string[][] Parametros)
        {
            SqlParameter[] paramC = new SqlParameter[Parametros.Length];

            try
            {
                int Total = Parametros.Length;// / RDTotalParam;
                int Longitud = 0;

                for (int w = 0; w < Total; w++)
                {
                    int.TryParse(Parametros[w][RDTamanno], out Longitud);

                    if (Parametros[w][RDValor] == RDNulo)
                    {
                        //paramC.Add(CreaParametroRD(Parametros[w][RDNombre], null, Equivalencia_SQL_SQLDBType(Parametros[w][RDTipo]), Longitud));
                        paramC[w] = CreaParametroRD(Parametros[w][RDNombre], System.DBNull.Value, Equivalencia_SQL_SQLDBType(Parametros[w][RDTipo]), Longitud);
                    }
                    else
                    {
                        //paramC.Add(CreaParametroRD(Parametros[w][RDNombre], Parametros[w][RDValor], Equivalencia_SQL_SQLDBType(Parametros[w][RDTipo]), Longitud));
                        paramC[w] = CreaParametroRD(Parametros[w][RDNombre], Parametros[w][RDValor], Equivalencia_SQL_SQLDBType(Parametros[w][RDTipo]), Longitud);
                    }
                }
            }
            catch (Exception ex)
            {
                RegistraLog("CreaParametrosRD", ex.Message, ex.StackTrace);
            }

            return paramC;
        }

        private void DLReporteDinamico(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string[][] Parametros, int RD_Id)
        {
            string MsjBD = "";
            ExportarExcel exportar = new ExportarExcel();
            DetalleReporte rpt = new DetalleReporte(RD_Id);

            if (rpt.Conexion.Contains("ErrorDetalleReporte"))
            {
                RegistraLog("DetalleReporte", rpt.Conexion, rpt.Script);
                return;
            }

            SqlConnection cn = null;
            SqlCommand cmd = null;
            DataTable Tabla = new DataTable();
            bool DebugUsr = false;

            try
            {
                cn = new SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings[rpt.Conexion].ConnectionString);
                cn.Open();
                cmd = new SqlCommand(rpt.Script, cn);

                switch (rpt.Tipo)
                {
                    case TiposScriptRD.Texto:
                        cmd.CommandType = CommandType.Text;
                        break;

                    case TiposScriptRD.StoredProcedure:
                        cmd.CommandType = CommandType.StoredProcedure;
                        break;

                    default:
                        cmd.CommandType = CommandType.Text;
                        break;
                }

                if (UserName.ToLowerInvariant() == System.Configuration.ConfigurationManager.ConnectionStrings["DebugUser"].ConnectionString.ToLowerInvariant())
                    DebugUsr = true;

                cmd.Parameters.AddRange(CreaParametrosRD(Parametros));

                if (DebugUsr)
                {
                    string Query = "";

                    foreach (SqlParameter param in cmd.Parameters)
                    {
                        Query += "Parametro: " + param.ParameterName + "; Valor: " + param.Value.ToString() + "; Tipo: " + param.SqlDbType.ToString() + Environment.NewLine;
                    }

                    Query += Environment.NewLine + cmd.CommandText;

                    RegistraLogExt("DLReporteDinamico", "DebugQuery", Query);
                }

                cmd.CommandTimeout = 1800; //30 minutos
                MsjBD = exportar.GenerarExcel(cmd.ExecuteReader(), RutaArchivos, Archivo, RegistrosPorHoja);
                cmd.Parameters.Clear();

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: true, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLReporteDinamico", "RD_Id: " + RD_Id.ToString() + Environment.NewLine + ex.Message, ex.StackTrace);
                ActualizaArchivo(UserName, Archivo, DocU_Finalizado: true, DocU_Observaciones: ex.Message);
            }
            finally
            {
                if (cn.State != ConnectionState.Closed)
                {
                    cn.Close();
                }
            }
        }

        #endregion ReportesDinamicos

        #region SAP

        #region Acumulados

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void Acumulados_SAP(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Anio, string Empleados, string Concepto_Nomina)
        {
            try
            {
                if (Anio.Trim() == "")
                    Anio = null;

                if (Empleados.Trim() == "")
                    Empleados = null;

                if (Concepto_Nomina.Trim() == "")
                    Concepto_Nomina = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLAcumulados_SAP(TipoDocumento, UserName, RutaArchivos, Archivo, RegistrosPorHoja, Anio, Empleados, Concepto_Nomina);
            }
            catch (Exception ex)
            {
                RegistraLog("Acumulados_SAP", ex.Message, ex.StackTrace);
            }
        }

        private void DLAcumulados_SAP(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Anio, string No_Pers, string Concepto_Nomina)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_AcumuladosRH");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Anio", DbType.String, Anio);
                db.AddInParameter(selectCommand, "@No_Pers", DbType.String, No_Pers);
                db.AddInParameter(selectCommand, "@Concepto_Nomina", DbType.String, Concepto_Nomina);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLAcumulados_SAP", ex.Message, ex.StackTrace);
            }
        }

        #endregion Acumulados

        #region AsignacionOrganizativa

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void AsignacionOrganizativa(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLAsignacionOrganizativa(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("AsignacionOrganizativa", ex.Message, ex.StackTrace);
            }
        }

        private void DLAsignacionOrganizativa(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_AsignacionOrganizativa");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLAsignacionOrganizativa", ex.Message, ex.StackTrace);
            }
        }

        #endregion AsignacionOrganizativa

        #region Auxiliares

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void AuxiliaresContables(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Ejercicio, string Cuenta_Mayor)
        {
            try
            {
                if (Sociedad.Trim() == "")
                    Sociedad = null;

                if (Ejercicio.Trim() == "")
                    Ejercicio = null;

                if (Cuenta_Mayor.Trim() == "")
                    Cuenta_Mayor = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLAuxiliaresContables(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Sociedad, Ejercicio, Cuenta_Mayor);
            }
            catch (Exception ex)
            {
                RegistraLog("AuxiliaresContables", ex.Message, ex.StackTrace);
            }
        }

        private void DLAuxiliaresContables(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Ejercicio, string Cuenta_Mayor)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_Auxiliares");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Sociedad", DbType.String, Sociedad);
                db.AddInParameter(selectCommand, "@Ejercicio", DbType.String, Ejercicio);
                db.AddInParameter(selectCommand, "@Cuenta_Mayor", DbType.String, Cuenta_Mayor);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLAuxiliaresContables", ex.Message, ex.StackTrace);
            }
        }

        #endregion Auxiliares

        #region AuxiliaresDetalle

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void AuxiliaresDetalle(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Ejercicio, string Cuenta_Mayor)
        {
            try
            {
                if (Sociedad.Trim() == "")
                    Sociedad = null;

                if (Ejercicio.Trim() == "")
                    Ejercicio = null;

                if (Cuenta_Mayor.Trim() == "")
                    Cuenta_Mayor = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLAuxiliaresDetalle(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Sociedad, Ejercicio, Cuenta_Mayor);
            }
            catch (Exception ex)
            {
                RegistraLog("AuxiliaresDetalle", ex.Message, ex.StackTrace);
            }
        }

        private void DLAuxiliaresDetalle(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Ejercicio, string Cuenta_Mayor)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_AuxiliaresDetalle");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Sociedad", DbType.String, Sociedad);
                db.AddInParameter(selectCommand, "@Ejercicio", DbType.String, Ejercicio);
                db.AddInParameter(selectCommand, "@Cuenta_Mayor", DbType.String, Cuenta_Mayor);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLAuxiliaresDetalle", ex.Message, ex.StackTrace);
            }
        }

        #endregion AuxiliaresDetalle

        #region Balanzas

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void BalanzasContables(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Ejercicio, string Cuenta_Mayor)
        {
            try
            {
                if (Sociedad.Trim() == "")
                    Sociedad = null;

                if (Ejercicio.Trim() == "")
                    Ejercicio = null;

                if (Cuenta_Mayor.Trim() == "")
                    Cuenta_Mayor = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLBalanzasContables(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Sociedad, Ejercicio, Cuenta_Mayor);
            }
            catch (Exception ex)
            {
                RegistraLog("BalanzasContables", ex.Message, ex.StackTrace);
            }
        }

        private void DLBalanzasContables(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Ejercicio, string Cuenta_Mayor)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_Balanzas");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Sociedad", DbType.String, Sociedad);
                db.AddInParameter(selectCommand, "@Ejercicio", DbType.String, Ejercicio);
                db.AddInParameter(selectCommand, "@Cuenta_Mayor", DbType.String, Cuenta_Mayor);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLBalanzasContables", ex.Message, ex.StackTrace);
            }
        }

        #endregion Balanzas

        #region ContratosLaborales

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void ContratosLaborales(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLContratosLaborales(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("ContratosLaborales", ex.Message, ex.StackTrace);
            }
        }

        private void DLContratosLaborales(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_ContratoLaboral");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLContratosLaborales", ex.Message, ex.StackTrace);
            }
        }

        #endregion ContratosLaborales

        #region Cuentas

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void Cuentas(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Cuenta_Mayor)
        {
            try
            {
                if (Sociedad.Trim() == "")
                    Sociedad = null;

                if (Cuenta_Mayor.Trim() == "")
                    Cuenta_Mayor = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLCuentas(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Sociedad, Cuenta_Mayor);
            }
            catch (Exception ex)
            {
                RegistraLog("Cuentas", ex.Message, ex.StackTrace);
            }
        }

        private void DLCuentas(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Sociedad, string Cuenta_Mayor)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_CatalogoCuentas");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Sociedad", DbType.String, Sociedad);
                db.AddInParameter(selectCommand, "@Cuenta_Mayor", DbType.String, Cuenta_Mayor);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLCuentas", ex.Message, ex.StackTrace);
            }
        }

        #endregion Cuentas

        #region DatosBancarios

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void DatosBancarios(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLDatosBancarios(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("DatosBancarios", ex.Message, ex.StackTrace);
            }
        }

        private void DLDatosBancarios(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_DatosBancarios");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLDatosBancarios", ex.Message, ex.StackTrace);
            }
        }

        #endregion DatosBancarios

        #region DatosPersonales

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void DatosPersonales(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLDatosPersonales(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("DatosPersonales", ex.Message, ex.StackTrace);
            }
        }

        private void DLDatosPersonales(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_DatosPersonales");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLDatosPersonales", ex.Message, ex.StackTrace);
            }
        }

        #endregion DatosPersonales

        #region Direcciones

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void Direcciones(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLDirecciones(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("Direcciones", ex.Message, ex.StackTrace);
            }
        }

        private void DLDirecciones(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_Direcciones");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLDirecciones", ex.Message, ex.StackTrace);
            }
        }

        #endregion Direcciones

        #region HorariosLaborales

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void HorariosLaborales(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLHorariosLaborales(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("HorariosLaborales", ex.Message, ex.StackTrace);
            }
        }

        private void DLHorariosLaborales(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_HorarioLaboral");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLHorariosLaborales", ex.Message, ex.StackTrace);
            }
        }

        #endregion HorariosLaborales

        #region IngresoEmpleados

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void IngresoEmpleados(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLIngresoEmpleados(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("IngresoEmpleados", ex.Message, ex.StackTrace);
            }
        }

        private void DLIngresoEmpleados(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_IngresoEmpleados");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLIngresoEmpleados", ex.Message, ex.StackTrace);
            }
        }

        #endregion IngresoEmpleados

        #region MedidasRH

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void MedidasRH(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLMedidasRH(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("MedidasRH", ex.Message, ex.StackTrace);
            }
        }

        private void DLMedidasRH(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_Medidas");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLMedidasRH", ex.Message, ex.StackTrace);
            }
        }

        #endregion MedidasRH

        #region NumAntEmp

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void NumAntEmp(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLNumAntEmp(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("NumAntEmp", ex.Message, ex.StackTrace);
            }
        }

        private void DLNumAntEmp(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_NumeroAnterior");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLNumAntEmp", ex.Message, ex.StackTrace);
            }
        }

        #endregion NumAntEmp

        #region RemuneracionEconomica

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void RemuneracionEconomica(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            try
            {
                if (Empleados.Trim() == "")
                    Empleados = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLRemuneracionEconomica(UserName, RutaArchivos, Archivo, RegistrosPorHoja, Empleados);
            }
            catch (Exception ex)
            {
                RegistraLog("RemuneracionEconomica", ex.Message, ex.StackTrace);
            }
        }

        private void DLRemuneracionEconomica(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Empleados)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("SAP");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("spR_RemuneracionEconomica");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@No_pers", DbType.String, Empleados);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLRemuneracionEconomica", ex.Message, ex.StackTrace);
            }
        }

        #endregion RemuneracionEconomica

        #endregion SAP

        #region Servidores

        #region CintasRespaldo

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void CintasRespaldo(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, int Tipo, int Obj_Id, string RC_Cinta)
        {
            try
            {
                int? TipoN = null;
                int? Obj_IdN = null;

                if (Tipo > 0)
                    TipoN = Tipo;

                if (Obj_Id > 0)
                    Obj_IdN = Obj_Id;

                if (RC_Cinta.Trim() == "")
                    RC_Cinta = null;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLCintasRespaldo(UserName, RutaArchivos, Archivo, RegistrosPorHoja, TipoN, Obj_IdN, RC_Cinta);
            }
            catch (Exception ex)
            {
                RegistraLog("CintasRespaldo", ex.Message, ex.StackTrace);
            }
        }

        private void DLCintasRespaldo(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, int? Tipo, int? Obj_Id, string RC_Cinta)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_Cintas");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@TR_Id", DbType.Int32, Tipo);
                db.AddInParameter(selectCommand, "@Obj_Id", DbType.Int32, Obj_Id);
                db.AddInParameter(selectCommand, "@RC_Cinta", DbType.String, RC_Cinta);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("CintasRespaldo", ex.Message, ex.StackTrace);
            }
        }

        #endregion CintasRespaldo

        #region MonitoreoSW

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void MonitoreoSW(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string UsuarioRevisor, string Pass, string Dominio, bool RevisarTodos, string EquipoEspecifico = "")
        {
            try
            {
                DataTable Resultados = new DataTable();
                De_CryptDLL.De_Crypt cripto = new De_CryptDLL.De_Crypt();

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                Pass = cripto.Desencriptar(Pass, StandardKey, true);

                if (RevisarTodos)
                {
                    DataTable Equipos = new DataTable();

                    Equipos = DLLeerEquipos(true);

                    if (Equipos.Rows.Count > 0)
                    {
                        for (int w = 0; w < Equipos.Rows.Count; w++)
                        {
                            string Equipo = Equipos.Rows[w][0].ToString().ToLower();

                            if (!string.IsNullOrWhiteSpace(Equipo))
                            {
                                try
                                {
                                    System.Security.Principal.WindowsImpersonationContext newUser = Impersonate.Impersonation.ImpersonateUser(UsuarioRevisor, Dominio, Pass);
                                    Resultados = InstalledSW.InstalledPrograms.GetRemoteInstalledProgramsDT(Equipo);

                                    if (Resultados.Rows.Count > 0 && Resultados.Columns.Count == 2)
                                    {
                                        DLLimpiarRegistros(Equipo);

                                        for (int w2 = 0; w2 < Resultados.Rows.Count; w2++)
                                        {
                                            DLInsertarRegistro(Equipo, Resultados.Rows[w2][0].ToString(), Resultados.Rows[w2][1].ToString(), UserName);
                                        }
                                    }
                                    else
                                    {
                                        RegistraLog("MonitoreoListaEquiposImp", "Equipo: " + Equipo, "No se obtuvo la lista de programas.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    RegistraLog("MonitoreoListaEquipos", "Equipo: " + Equipo + " --> " + ex.Message, ex.StackTrace);

                                    if (ex is System.Security.SecurityException)
                                        DLInsertarRegistro(Equipo, ex.Message, "", UserName);
                                }
                            }
                        }

                        DLMonitoreoSW(UserName, RutaArchivos, Archivo, RegistrosPorHoja, null);
                    }
                    else
                    {
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja, "No se obtuvo la lista de equipos");
                    }
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(EquipoEspecifico))
                    {
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja, "Debe especificarse un nombre de equipo");
                    }
                    else
                    {
                        try
                        {
                            System.Security.Principal.WindowsImpersonationContext newUser = Impersonate.Impersonation.ImpersonateUser(UsuarioRevisor, Dominio, Pass);
                            Resultados = InstalledSW.InstalledPrograms.GetRemoteInstalledProgramsDT(EquipoEspecifico);

                            if (Resultados.Rows.Count > 0 && Resultados.Columns.Count == 2)
                            {
                                DLLimpiarRegistros(EquipoEspecifico);

                                for (int w = 0; w < Resultados.Rows.Count; w++)
                                {
                                    DLInsertarRegistro(EquipoEspecifico, Resultados.Rows[w][0].ToString(), Resultados.Rows[w][1].ToString(), UserName);
                                }

                                DLMonitoreoSW(UserName, RutaArchivos, Archivo, RegistrosPorHoja, EquipoEspecifico);
                            }
                            else
                            {
                                ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja, "No se obtuvo la lista de programas. Equipo: " + EquipoEspecifico);
                            }
                        }
                        catch (Exception ex)
                        {
                            RegistraLog("MonitoreoSWEqEspecifico", ex.Message, ex.StackTrace);

                            if (ex is System.Security.SecurityException)
                                DLInsertarRegistro(EquipoEspecifico, ex.Message, "", UserName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                RegistraLog("MonitoreoSW", ex.Message, ex.StackTrace);
            }
        }

        private DataTable DLLeerEquipos(bool? SrvEMSW_AplicarRevision = null)
        {
            DataTable Tabla = new DataTable();
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;

            try
            {
                selectCommand = db.GetSqlStringCommand("stpS_ServidoresEquiposMonitoreoSW");
                selectCommand.CommandType = CommandType.StoredProcedure;

                int vSrvEMSW_AplicarRevision = 0;

                if (SrvEMSW_AplicarRevision == null)
                {
                    vSrvEMSW_AplicarRevision = 0;
                }
                else
                {
                    if (SrvEMSW_AplicarRevision == true)
                        vSrvEMSW_AplicarRevision = 1;
                    else
                        vSrvEMSW_AplicarRevision = 2;
                }

                db.AddInParameter(selectCommand, "@SrvEMSW_AplicarRevision", DbType.Int32, vSrvEMSW_AplicarRevision);
                Tabla.Load(db.ExecuteReader(selectCommand));
            }
            catch (Exception ex)
            {
                RegistraLog("DLLeerEquipos", ex.Message, ex.StackTrace);
            }

            return Tabla;
        }

        private void DLLimpiarRegistros(string SrvMSW_Equipo)
        {
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpD_ServidoresMonitoreoSW");
                selectCommand.CommandType = CommandType.StoredProcedure;

                DateTime SrvMSW_Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                db.AddInParameter(selectCommand, "@SrvMSW_Equipo", DbType.String, SrvMSW_Equipo);
                db.AddInParameter(selectCommand, "@SrvMSW_Fecha", DbType.DateTime, SrvMSW_Fecha);

                db.ExecuteNonQuery(selectCommand);
            }
            catch (Exception ex)
            {
                RegistraLog("DLLimpiarRegistros", ex.Message, ex.StackTrace);
            }
        }

        private void DLInsertarRegistro(string SrvMSW_Equipo, string SrvMSW_NombreSW, string SrvMSW_Version, string UserName)
        {
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpI_ServidoresMonitoreoSW");
                selectCommand.CommandType = CommandType.StoredProcedure;

                DateTime SrvMSW_Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                db.AddInParameter(selectCommand, "@SrvMSW_Equipo", DbType.String, SrvMSW_Equipo);
                db.AddInParameter(selectCommand, "@SrvMSW_Fecha", DbType.DateTime, SrvMSW_Fecha);
                db.AddInParameter(selectCommand, "@SrvMSW_NombreSW", DbType.String, SrvMSW_NombreSW);
                db.AddInParameter(selectCommand, "@SrvMSW_Version", DbType.String, SrvMSW_Version);
                db.AddInParameter(selectCommand, "@UserName", DbType.String, UserName);

                db.ExecuteNonQuery(selectCommand);
            }
            catch (Exception ex)
            {
                RegistraLog("DLLimpiarRegistros", ex.Message, ex.StackTrace);
            }
        }

        private void DLMonitoreoSW(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string SrvMSW_Equipo)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpS_ServidoresMonitoreoSW");
                selectCommand.CommandType = CommandType.StoredProcedure;

                DateTime SrvMSW_Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);

                db.AddInParameter(selectCommand, "@SrvMSW_Equipo", DbType.String, SrvMSW_Equipo);
                db.AddInParameter(selectCommand, "@SrvMSW_Fecha", DbType.DateTime, SrvMSW_Fecha);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLMonitoreoSW", ex.Message, ex.StackTrace);
            }
        }

        #endregion MonitoreoSW

        #endregion Servidores

        #region Software

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void InventarioSW(int TipoDocumento, string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string SWE_Id, string SWG_Id, string SW_Descripcion, string SW_Version, string SWEx_NoParte, string SWEx_Llave, string SWEx_Ubicacion, string SWEx_Observaciones, string SWEx_EnExistencia)
        {
            try
            {
                bool? EnExistencia = null;

                SWEx_EnExistencia = SWEx_EnExistencia.Trim();

                if (SWE_Id.Trim() == "")
                    SWE_Id = null;

                if (SWG_Id.Trim() == "")
                    SWG_Id = null;

                if (SW_Descripcion.Trim() == "")
                    SW_Descripcion = null;

                if (SW_Version.Trim() == "")
                    SW_Version = null;

                if (SWEx_NoParte.Trim() == "")
                    SWEx_NoParte = null;

                if (SWEx_Llave.Trim() == "")
                    SWEx_Llave = null;

                if (SWEx_Ubicacion.Trim() == "")
                    SWEx_Ubicacion = null;

                if (SWEx_Observaciones.Trim() == "")
                    SWEx_Observaciones = null;

                if (SWEx_EnExistencia == "NO")
                    EnExistencia = false;
                else if (SWEx_EnExistencia == "SI")
                    EnExistencia = true;

                IngresaArchivo(TipoDocumento, UserName, Archivo);
                DLInventarioSW(UserName, RutaArchivos, Archivo, RegistrosPorHoja, SWE_Id, SWG_Id, SW_Descripcion, SW_Version, SWEx_NoParte, SWEx_Llave, SWEx_Ubicacion, SWEx_Observaciones, EnExistencia);
            }
            catch (Exception ex)
            {
                RegistraLog("InventarioSW", ex.Message, ex.StackTrace);
            }
        }

        private void DLInventarioSW(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string SWE_Id, string SWG_Id, string SW_Descripcion, string SW_Version, string SWEx_NoParte, string SWEx_Llave, string SWEx_Ubicacion, string SWEx_Observaciones, bool? SWEx_EnExistencia)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpR_InventarioSW");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@SWE_Id", DbType.String, SWE_Id);
                db.AddInParameter(selectCommand, "@SWG_Id", DbType.String, SWG_Id);
                db.AddInParameter(selectCommand, "@SW_Descripcion", DbType.String, SW_Descripcion);
                db.AddInParameter(selectCommand, "@SW_Version", DbType.String, SW_Version);
                db.AddInParameter(selectCommand, "@SWEx_NoParte", DbType.String, SWEx_NoParte);
                db.AddInParameter(selectCommand, "@SWEx_Llave", DbType.String, SWEx_Llave);
                db.AddInParameter(selectCommand, "@SWEx_Ubicacion", DbType.String, SWEx_Ubicacion);
                db.AddInParameter(selectCommand, "@SWEx_Observaciones", DbType.String, SWEx_Observaciones);
                db.AddInParameter(selectCommand, "@SWEx_EnExistencia", DbType.Boolean, SWEx_EnExistencia);
                db.AddInParameter(selectCommand, "@IncluirEstadisticas", DbType.Boolean, false);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, true);

                    if (!System.IO.File.Exists(System.IO.Path.Combine(RutaArchivos, Archivo)))
                        ReportarArchivoVacio(UserName, RutaArchivos, Archivo, RegistrosPorHoja);
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLDiscosSrv", ex.Message, ex.StackTrace);
            }
        }

        #endregion Software

        #region WS

        #region RegistrosActividad

        private void RegistraLog(string Modulo, string Msj, string Stack)
        {
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                if (Modulo.Length > 50)
                    Modulo = Modulo.Substring(0, 50);

                if (Msj.Length > 250)
                    Msj = Msj.Substring(0, 250);

                if (Stack.Length > 2000)
                    Stack = Stack.Substring(0, 2000);

                selectCommand = db.GetSqlStringCommand("stpI_Logs");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Log_Sistema", DbType.String, "ExportarExcelWS");
                db.AddInParameter(selectCommand, "@Log_Modulo", DbType.String, Modulo);
                db.AddInParameter(selectCommand, "@Log_Mensaje", DbType.String, Msj);
                db.AddInParameter(selectCommand, "@Log_StackTrace", DbType.String, Stack);

                db.ExecuteNonQuery(selectCommand);
            }
            catch { }
        }

        private void RegistraLogExt(string Modulo, string Msj, string Stack)
        {
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                if (Modulo.Length > 50)
                    Modulo = Modulo.Substring(0, 50);

                if (Msj.Length > 250)
                    Msj = Msj.Substring(0, 250);

                selectCommand = db.GetSqlStringCommand("stpI_LogsExt");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Log_Sistema", DbType.String, "ExportarExcelWS");
                db.AddInParameter(selectCommand, "@Log_Modulo", DbType.String, Modulo);
                db.AddInParameter(selectCommand, "@Log_Mensaje", DbType.String, Msj);
                db.AddInParameter(selectCommand, "@Log_StackTrace", DbType.String, Stack);

                db.ExecuteNonQuery(selectCommand);
            }
            catch { }
        }

        #endregion RegistrosActividad

        #region MetodosGenerales

        #region Variables

        private string StandardKey = "HSC941011SU6";

        #endregion Variables

        private void ReportarArchivoVacio(string UserName, string RutaArchivos, string Archivo, int RegistrosPorHoja, string Mensaje = null)
        {
            string MsjBD = "";
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpS_ReportarArchivoVacio");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@Mensaje", DbType.String, Mensaje);

                MsjBD = exportar.GenerarExcel(db.ExecuteReader(selectCommand), RutaArchivos, Archivo, RegistrosPorHoja);

                if (MsjBD.Contains("Error:"))
                {
                    RegistraLog("Clase", "", MsjBD);
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: false, DocU_Observaciones: MsjBD);
                }
                else
                {
                    ActualizaArchivo(UserName, Archivo, DocU_Finalizado: true, DocU_Observaciones: "No se encontraron registros para su búsqueda");
                }
            }
            catch (Exception ex)
            {
                RegistraLog("DLDiscosSrv", ex.Message, ex.StackTrace);
            }
        }

        private void IngresaArchivo(int DocT_Id, string UserName, string DocU_Nombre, bool DocU_Finalizado = false, bool DocU_Eliminado = false)
        {
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                selectCommand = db.GetSqlStringCommand("stpI_DocumentosUsuario");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@UserName", DbType.String, UserName);
                db.AddInParameter(selectCommand, "@DocT_Id", DbType.Int32, DocT_Id);
                db.AddInParameter(selectCommand, "@DocU_Nombre", DbType.String, DocU_Nombre);
                db.AddInParameter(selectCommand, "@DocU_Finalizado", DbType.Boolean, DocU_Finalizado);
                db.AddInParameter(selectCommand, "@DocU_Eliminado", DbType.Boolean, DocU_Eliminado);

                db.ExecuteNonQuery(selectCommand);
            }
            catch (Exception ex)
            {
                RegistraLog("IngresaArchivo", ex.Message, ex.StackTrace);
            }
        }

        private void ActualizaArchivo(string UserName, string DocU_Nombre, bool DocU_Finalizado, bool DocU_Eliminado = false, string DocU_Observaciones = null)
        {
            Database db = EnterpriseLibraryContainer.Current.GetInstance<Database>("Inventario");
            DbCommand selectCommand = null;
            ExportarExcel exportar = new ExportarExcel();

            try
            {
                if (!string.IsNullOrWhiteSpace(DocU_Observaciones) && DocU_Observaciones.Length > 500)
                    DocU_Observaciones = DocU_Observaciones.Substring(0, 500);

                selectCommand = db.GetSqlStringCommand("stpU_DocumentosUsuario");
                selectCommand.CommandType = CommandType.StoredProcedure;

                db.AddInParameter(selectCommand, "@UserName", DbType.String, UserName);
                db.AddInParameter(selectCommand, "@DocU_Nombre", DbType.String, DocU_Nombre);
                db.AddInParameter(selectCommand, "@DocU_Finalizado", DbType.Boolean, DocU_Finalizado);
                db.AddInParameter(selectCommand, "@DocU_Eliminado", DbType.Boolean, DocU_Eliminado);
                db.AddInParameter(selectCommand, "@DocU_Observaciones", DbType.String, DocU_Observaciones);

                db.ExecuteNonQuery(selectCommand);
                GC.Collect();
            }
            catch (Exception ex)
            {
                RegistraLog("ActualizaArchivo", ex.Message, ex.StackTrace);
            }
        }

        [WebMethod]
        [System.Web.Services.Protocols.SoapDocumentMethod(OneWay = true)]
        public void RegistrarArchivoTempGeneral(int TipoDocumento, string UserName, string Archivo, bool Finalizado)
        {
            try
            {
                IngresaArchivo(TipoDocumento, UserName, Archivo);
                ActualizaArchivo(UserName, Archivo, Finalizado);
            }
            catch (Exception ex)
            {
                RegistraLog("RegistrarArchivoTempGeneral", ex.Message, ex.StackTrace);
            }
        }

        #endregion MetodosGenerales

        #endregion WS
    }
}