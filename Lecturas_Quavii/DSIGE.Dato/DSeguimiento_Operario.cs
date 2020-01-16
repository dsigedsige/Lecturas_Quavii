using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;
using System.IO;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;
using System.Drawing;

namespace DSIGE.Dato
{
    public class DSeguimiento_Operario
    {
     //   public DSeguimiento_Operario() {
     //       DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
     //   }

     //   /// <summary>
     //   /// Autor: rcontreras
     //   /// Fecha: 01-10-2015
     //   /// </summary>
     //   /// <param name="oRq"></param>
     //   /// <returns></returns>
     //   public List<UbicacionOperario> DSeguimientoOperarioGPS(int idOperario)
     //   {
     //       try
     //       {
     //           List<UbicacionOperario> lOb = new List<UbicacionOperario>();

     //           using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_seguimiento_ubica_operario_maps", idOperario))
     //           {
     //               if (iDr != null)
     //               {
     //                   while (iDr.Read())
     //                   {
     //                       lOb.Add(
     //                               new UbicacionOperario()
     //                               {
     //                                   id_gps = Convert.ToInt32(iDr["ID_Gps"]),
     //                                   id_ope = Convert.ToInt32(iDr["id_operario"]),
     //                                   latitud = Convert.ToString(iDr["GPS_Latitud"]),
     //                                   longitud = Convert.ToString(iDr["GPS_Longitud"]),
     //                                   paro = Convert.ToString(iDr["flag_paro"])
     //                               }
     //                           );
     //                   }
     //               }
     //           }

     //           return lOb;
     //       }
     //       catch (Exception e)
     //       {
     //           throw e;
     //       }
     //   }

     ////   id_Operario,ope_nombre,latitud_lectura, longitud_lectura, fechaGPS, tiempo_Espera, Fecha_Inicio_parada, fecha_fin_parada

     //   public List<UbicacionOperario_GPS> DSeguimientoOperario_GPS(int pIdOpe, string  pFechaAsiga,string lista)
     //   {
     //       try
     //       {
     //           List<UbicacionOperario_GPS> lOb = new List<UbicacionOperario_GPS>();

     //           using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL", pIdOpe, pFechaAsiga, lista))
     //           {
     //               if (iDr != null)
     //               {
                        
     //                   while (iDr.Read())
     //                   {
     //                       lOb.Add(
     //                               new UbicacionOperario_GPS()
     //                               {
     //                                   id_Operario = Convert.ToInt32(iDr["id_Operario"]),
     //                                   ope_nombre = Convert.ToString(iDr["ope_nombre"]),
     //                                   latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
     //                                   longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
     //                                   FechaGPS = Convert.ToString(iDr["FechaGPS"]),
     //                                   tiempo_Espera = Convert.ToInt32(iDr["tiempo_Espera"]),
     //                                   Fecha_Inicio_parada = Convert.ToString(iDr["Fecha_Inicio_parada"]),
     //                                   fecha_fin_parada = Convert.ToString(iDr["fecha_fin_parada"]),
     //                                   Minutos_Espera_Conv = Convert.ToString(iDr["Minutos_Espera_Conv"])
     //                               }
     //                           );
     //                   }
     //               }
     //           }

     //           return lOb;
     //       }
     //       catch (Exception e)
     //       {
     //           throw e;
     //       }
     //   }

     //   public List<UbicacionOperario_GPS> DSeguimientoOperario_GPS_II(string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
     //   {
     //       try
     //       {
     //           List<UbicacionOperario_GPS> lOb = new List<UbicacionOperario_GPS>();

     //           using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL_II", pFechaAsiga, servicio, operario, suministro, medidor))
     //           {
     //               if (iDr != null)
     //               {

     //                   while (iDr.Read())
     //                   {
     //                       lOb.Add(
     //                               new UbicacionOperario_GPS()
     //                               {
     //                                   id_Operario = Convert.ToInt32(iDr["id_Operario"]),
     //                                   ope_nombre = Convert.ToString(iDr["ope_nombre"]),
     //                                   latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
     //                                   longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
     //                                   FechaGPS = Convert.ToString(iDr["FechaGPS"]),
     //                                   tiempo_Espera = Convert.ToInt32(iDr["tiempo_Espera"]),
     //                                   Fecha_Inicio_parada = Convert.ToString(iDr["Fecha_Inicio_parada"]),
     //                                   fecha_fin_parada = Convert.ToString(iDr["fecha_fin_parada"]),
     //                                   Minutos_Espera_Conv = Convert.ToString(iDr["Minutos_Espera_Conv"])
     //                               }
     //                           );
     //                   }
     //               }
     //           }

     //           return lOb;
     //       }
     //       catch (Exception e)
     //       {
     //           throw e;
     //       }
     //   }


        public DSeguimiento_Operario()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 01-10-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<UbicacionOperario> DSeguimientoOperarioGPS(int idOperario)
        {
            try
            {
                List<UbicacionOperario> lOb = new List<UbicacionOperario>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_seguimiento_ubica_operario_maps", idOperario))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario()
                                    {
                                        id_gps = Convert.ToInt32(iDr["ID_Gps"]),
                                        id_ope = Convert.ToInt32(iDr["id_operario"]),
                                        latitud = Convert.ToString(iDr["GPS_Latitud"]),
                                        longitud = Convert.ToString(iDr["GPS_Longitud"]),
                                        paro = Convert.ToString(iDr["flag_paro"])
                                    }
                                );
                        }
                    }
                }

                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<UbicacionOperario_GPS> DSeguimientoOperario_GPS(int pIdOpe, string pFechaAsiga, string lista)
        {
            try
            {
                List<UbicacionOperario_GPS> lOb = new List<UbicacionOperario_GPS>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL", pIdOpe, pFechaAsiga, lista))
                {
                    if (iDr != null)
                    {

                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario_GPS()
                                    {
                                        id_Operario = Convert.ToInt32(iDr["id_Operario"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
                                        longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
                                        FechaGPS = Convert.ToString(iDr["FechaGPS"]),
                                        tiempo_Espera = Convert.ToInt32(iDr["tiempo_Espera"]),
                                        Fecha_Inicio_parada = Convert.ToString(iDr["Fecha_Inicio_parada"]),
                                        fecha_fin_parada = Convert.ToString(iDr["fecha_fin_parada"]),
                                        Minutos_Espera_Conv = Convert.ToString(iDr["Minutos_Espera_Conv"])
                                    }
                                );
                        }
                    }
                }

                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        public List<UbicacionOperario_GPS> DSeguimientoOperario_GPS_II(int Local, string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            try
            {
                List<UbicacionOperario_GPS> lOb = new List<UbicacionOperario_GPS>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL_II", pFechaAsiga, servicio, operario, suministro, medidor, Local))
                {
                    if (iDr != null)
                    {

                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario_GPS()
                                    {
                                        id_Operario = Convert.ToInt32(iDr["id_Operario"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
                                        longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
                                        FechaGPS = Convert.ToString(iDr["FechaGPS"]),
                                        tiempo_Espera = Convert.ToInt32(iDr["tiempo_Espera"]),
                                        Fecha_Inicio_parada = Convert.ToString(iDr["Fecha_Inicio_parada"]),
                                        fecha_fin_parada = Convert.ToString(iDr["fecha_fin_parada"]),
                                        Minutos_Espera_Conv = Convert.ToString(iDr["Minutos_Espera_Conv"])
                                    }
                                );
                        }
                    }
                }

                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<UbicacionOperario_GPS> DSeguimientoOperario_GPS_Reparto(string pFechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            try
            {
                List<UbicacionOperario_GPS> lOb = new List<UbicacionOperario_GPS>();

                //using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL_REPARTO", pFechaAsiga, servicio, operario, suministro, medidor))
                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL_REPARTO_NEW", pFechaAsiga, servicio, operario, suministro, medidor))
                {
                    if (iDr != null)
                    {

                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario_GPS()
                                    {
                                        id_Operario = Convert.ToInt32(iDr["id_Operario"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
                                        longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
                                        FechaGPS = Convert.ToString(iDr["FechaGPS"]),
                                        tiempo_Espera = Convert.ToInt32(iDr["tiempo_Espera"]),
                                        Fecha_Inicio_parada = Convert.ToString(iDr["Fecha_Inicio_parada"]),
                                        fecha_fin_parada = Convert.ToString(iDr["fecha_fin_parada"]),
                                        Minutos_Espera_Conv = Convert.ToString(iDr["Minutos_Espera_Conv"])
                                    }
                                );
                        }
                    }
                }

                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<UbicacionOperario_GPS> DSeguimientoOperario_GPS_Resumen(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                List<UbicacionOperario_GPS> lOb = new List<UbicacionOperario_GPS>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SEGUIMIENTO_OPERARIO_MOVIL_RESUMEN", FechaAsiga, servicio, operario))
                {
                    if (iDr != null)
                    {

                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new UbicacionOperario_GPS()
                                    {
                                        id_Operario = Convert.ToInt32(iDr["id_Operario"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
                                        longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
                                        FechaGPS = Convert.ToString(iDr["FechaGPS"]),
                                        tiempo_Espera = Convert.ToInt32(iDr["tiempo_Espera"]),
                                        Fecha_Inicio_parada = Convert.ToString(iDr["Fecha_Inicio_parada"]),
                                        fecha_fin_parada = Convert.ToString(iDr["fecha_fin_parada"]),
                                        Minutos_Espera_Conv = Convert.ToString(iDr["Minutos_Espera_Conv"])
                                    }
                                );
                        }
                    }
                }

                return lOb;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
               
        public string Ddescargar_informacion_reparto_excel(string pFechaAsiga, int servicio, int operario, string suministro, string medidor,int id_usuario)
        {
            var resultado = "";
            var nombreFile = "";
            var nombreExcel = "";
            var servidor = "";

            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["Archivos"];

                nombreFile = id_usuario + "_data_reparto_" + servidor + ".xlsx";
                nombreExcel = "dataReparto";

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_SEGUIMIENTO_OPERARIO_EXCEL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = pFechaAsiga;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = operario;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            if (dt_detalle.Rows.Count <= 0)
                            {
                                resultado = "0|No hay informacion disponible";
                            }
                            else
                            {
                                resultado = GenerarArchivoExcel_data(dt_detalle, nombreFile, nombreExcel);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                resultado = "0|" + ex.Message;
            }
            return resultado;
        }

        public string GenerarArchivoExcel_data(DataTable dt_detalles, string nombreFile, string nombreExcel)
        {
            string _ruta = "";
            string Res = "";
            int _fila = 2;
            try
            {
                _ruta = HttpContext.Current.Server.MapPath("~/Temp/" + nombreFile);
                string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];
                ruta_descarga = ruta_descarga + nombreFile;

                FileInfo _file = new FileInfo(_ruta);
                if (_file.Exists)
                {
                    _file.Delete();
                    _file = new FileInfo(_ruta);
                }

                using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
                {
                    Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add(nombreExcel);
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 9));

                    for (int i = 1; i <= 12; i++)
                    {
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    oWs.Cells[1, 1].Value = "Unidad Lectura";
                    oWs.Cells[1, 2].Value = "Cuenta Contrato";
                    oWs.Cells[1, 3].Value = "Cliente";
                    oWs.Cells[1, 4].Value = "Direccion";
                    oWs.Cells[1, 5].Value = "Urbanización";
                    oWs.Cells[1, 6].Value = "Distrito";

                    oWs.Cells[1, 7].Value = "Latitud";
                    oWs.Cells[1, 8].Value = "Longitud";
                    oWs.Cells[1, 9].Value = "Fecha";

                    oWs.Cells[1, 10].Value = "Hora";
                    oWs.Cells[1, 11].Value = "Porcion";
                    oWs.Cells[1, 12].Value = "Mes";

                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        oWs.Cells[_fila, 1].Value = oBj["Cod_UnidadLectura"].ToString();
                        oWs.Cells[_fila, 2].Value = oBj["CuentaContrato_Reparto"].ToString();
                        oWs.Cells[_fila, 3].Value = oBj["nombreCliente_Reparto"].ToString();
                        oWs.Cells[_fila, 4].Value = oBj["Direccion_Reparto"].ToString();

                        oWs.Cells[_fila, 5].Value = oBj["urbanizacion_Reparto"].ToString();
                        oWs.Cells[_fila, 6].Value = oBj["distrito_Reparto"].ToString();
                        oWs.Cells[_fila, 7].Value = oBj["latitud_lectura"].ToString();
                        oWs.Cells[_fila, 8].Value = oBj["longitud_lectura"].ToString();

                        oWs.Cells[_fila, 9].Value = oBj["fechaRecepcion_Reparto"].ToString();
                        oWs.Cells[_fila, 10].Value = oBj["hora"].ToString();
                        oWs.Cells[_fila, 11].Value = oBj["porcion"].ToString();
                        oWs.Cells[_fila, 12].Value = oBj["mes"].ToString();
                        _fila++;
                    }

                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
 

                    for (int k = 1; k <= 12; k++)
                    {
                        oWs.Column(k).AutoFit();
                    }
                    oEx.Save();
                }

                Res = "1|" + ruta_descarga;
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }

        public List<Seguimiento_Operarios_E> Capa_Dato_get_Operario_Gps(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Seguimiento_Operarios_E> List_detalles = new List<Seguimiento_Operarios_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_S_EFECTIVIDAD_MAPA_OPERARIOGPS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = operario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Seguimiento_Operarios_E obj_entidad = new Seguimiento_Operarios_E();

                                obj_entidad.id_operario = Convert.ToInt32(Fila["id_operario"].ToString());
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.latitud = Fila["latitud"].ToString();
                                obj_entidad.longitud = Fila["longitud"].ToString();
                                obj_entidad.fecha_gps = Fila["fecha_gps"].ToString();

                                List_detalles.Add(obj_entidad);
                            }

                        }
                    }
                }
                return List_detalles;
            }
            catch (Exception )
            {
                throw ;
            }

        }

        public List<Seguimiento_Operarios_E> Capa_Dato_get_Suministros_Mapa(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Seguimiento_Operarios_E> List_detalles = new List<Seguimiento_Operarios_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_S_EFECTIVIDAD_MAPA_SUMINISTROS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = operario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Seguimiento_Operarios_E obj_entidad = new Seguimiento_Operarios_E();
                                
                                obj_entidad.id_lectura = Convert.ToInt32(Fila["id_lectura"].ToString());
                                obj_entidad.suministro = Fila["suministro"].ToString();
                                obj_entidad.id_operario = Convert.ToInt32(Fila["id_operario"].ToString());
                                obj_entidad.operario = Fila["operario"].ToString();

                                obj_entidad.cliente = Fila["cliente"].ToString();
                                obj_entidad.direccion = Fila["direccion"].ToString();
                                obj_entidad.distrito = Fila["distrito"].ToString();

                                obj_entidad.latitud = Fila["latitud"].ToString();
                                obj_entidad.longitud = Fila["longitud"].ToString();

                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.tiene_foto = Fila["tiene_foto"].ToString();

                                List_detalles.Add(obj_entidad);
                            }

                        }
                    }
                }
                return List_detalles;
            }
            catch (Exception)
            {
                throw;
            }

        }


        public List<Seguimiento_Operarios_E> Capa_Datos_get_Suministros_Trabajados(string FechaAsiga, int servicio, int operario)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Seguimiento_Operarios_E> List_detalles = new List<Seguimiento_Operarios_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_S_EFECTIVIDAD_MAPA_SUMINISTROS_LEIDOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = operario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Seguimiento_Operarios_E obj_entidad = new Seguimiento_Operarios_E();

                                obj_entidad.id_lectura = Convert.ToInt32(Fila["id_lectura"].ToString());
                                obj_entidad.suministro = Fila["suministro"].ToString();
                                obj_entidad.id_operario = Convert.ToInt32(Fila["id_operario"].ToString());
                                obj_entidad.operario = Fila["operario"].ToString();

                                obj_entidad.cliente = Fila["cliente"].ToString();
                                obj_entidad.direccion = Fila["direccion"].ToString();
                                obj_entidad.distrito = Fila["distrito"].ToString();

                                obj_entidad.latitud = Fila["latitud"].ToString();
                                obj_entidad.longitud = Fila["longitud"].ToString();

                                obj_entidad.lectura_movil = Fila["lectura_movil"].ToString();
                                obj_entidad.fecha_lectura = Fila["fecha_lectura"].ToString();

                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.tiene_foto = Fila["tiene_foto"].ToString();

                                List_detalles.Add(obj_entidad);
                            }

                        }
                    }
                }
                return List_detalles;
            }
            catch (Exception)
            {
                throw;
            }

        }


    }
}
