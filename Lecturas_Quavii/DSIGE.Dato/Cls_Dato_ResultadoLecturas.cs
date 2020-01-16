using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Configuration;
using Ionic.Zip;
using System.IO;
using System.Web;
using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;
using System.Drawing;

namespace DSIGE.Dato
{
    public class Cls_Dato_ResultadoLecturas
    {
       string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
        
       public List<ResultadoLecturas_E> Capa_Dato_Get_ListaServicio()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_II", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                               obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                               obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                               obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<ResultadoLecturas_E> Capa_Dato_Get_ListaOperarios()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_OPERARIO", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                               obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                               obj_entidad.desc_operario = Fila["desc_operario"].ToString();
 
                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<ResultadoLecturas_E> Capa_Dato_Get_ListaSector()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_CARGO_LECTURAS", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                               obj_entidad.Sector = Fila["Sector"].ToString();
 

                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }
                                     
       public string Capa_Dato_Set_VerificacionFotos(string lecturas)
       {
           string Resultado = "";
           try
           {
               //obteniendo Datos de Facturacion
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_I_U_RESUMEN_FOTOS_VERIFICAR_FOTOS", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@lecturas", SqlDbType.VarChar).Value = lecturas;
                       cmd.ExecuteNonQuery();
                   }
               }
               Resultado = "OK";
           }
           catch (Exception ex)
           {
               Resultado = ex.Message;
           }
           return Resultado;
       }


       public List<ResultadoLecturas_E> Capa_Dato_Get_ResumenLecturas(int Local , string FechaAsignacion, int id_tiposervicio, int id_supervisor, int id_operario_supervisor, string ciclo)
        {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("NEW_SP_S_RESULTADO_LECTURAS", cn))
                   {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;


                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = FechaAsignacion;
                        cmd.Parameters.Add("@id_tiposervicio", SqlDbType.Int).Value = id_tiposervicio;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;
                        cmd.Parameters.Add("@ciclo", SqlDbType.VarChar).Value = ciclo;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = Local;

                        DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                                obj_entidad.total_lectura = Fila["total_lectura"].ToString();
                                obj_entidad.porc_lectura = Fila["porc_lectura"].ToString();
                                obj_entidad.total_ejecutado = Fila["total_ejecutado"].ToString();

                                obj_entidad.porc_ejecutado = Fila["porc_ejecutado"].ToString();
                                obj_entidad.total_pendiente = Fila["total_pendiente"].ToString();
                                obj_entidad.porc_pendiente = Fila["porc_pendiente"].ToString();

                                obj_entidad.total_fotos = Fila["total_fotos"].ToString();
                                obj_entidad.porc_fotos = Fila["porc_fotos"].ToString();         
       
                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<ResultadoLecturas_E> Capa_Dato_Get_ResumenLecturas_Observacion(int anio, int mes, string sector)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_RESULTADO_LECTURAS_OBSERVACION_II", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;

                       cmd.Parameters.Add("@anio", SqlDbType.Int).Value = anio;
                       cmd.Parameters.Add("@mes", SqlDbType.Int).Value = mes;
                       cmd.Parameters.Add("@id_sector", SqlDbType.VarChar).Value = sector;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                               obj_entidad.codigo = Fila["codigo"].ToString();
                               obj_entidad.observacion = Fila["observacion"].ToString();
                               obj_entidad.total = Fila["total"].ToString();

                               obj_entidad.conFoto = Fila["conFoto"].ToString();
                               obj_entidad.sinFoto = Fila["sinFoto"].ToString();
                               obj_entidad.porcentaje = Fila["porcentaje"].ToString();

                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<ResultadoLecturas_E> Capa_Dato_Get_ResumenLecturas_Detallado(int Local , string FechaAsignacion, int id_tiposervicio, int id_supervisor, int id_operario_supervisor, string ciclo)
        {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("NEW_SP_S_RESULTADO_LECTURAS_Detalle_Operario", cn))
                   {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = FechaAsignacion;
                        cmd.Parameters.Add("@id_tiposervicio", SqlDbType.Int).Value = id_tiposervicio;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;
                        cmd.Parameters.Add("@ciclo", SqlDbType.VarChar).Value = ciclo;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = Local;

                        DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                                obj_entidad.id_Operario =Convert.ToInt32(Fila["id_Operario"].ToString());
                                obj_entidad.apellidos_operario = Fila["apellidos_operario"].ToString();
                                obj_entidad.Total = Fila["Total"].ToString();

                                obj_entidad.Ejecutada = Fila["Ejecutada"].ToString();
                                obj_entidad.Pendiente = Fila["Pendiente"].ToString();
                                obj_entidad.Foto = Fila["Foto"].ToString();

                                obj_entidad.Porcentaje = Fila["Porcentaje"].ToString();
                                obj_entidad.Minimo = Fila["Minimo"].ToString();
                                obj_entidad.Maximo = Fila["Maximo"].ToString();

                                obj_entidad.totalHora = Fila["totalHora"].ToString();
                                obj_entidad.FueraHora = Fila["FueraHora"].ToString();         

                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<ResultadoLecturas_E> Capa_Dato_Get_FotoSelfie(int anio, int mes, string sector, int operario)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_RESULTADO_LECTURAS_FOTO_SELFIE_II", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;

                       cmd.Parameters.Add("@anio", SqlDbType.Int).Value = anio;
                       cmd.Parameters.Add("@mes", SqlDbType.Int).Value = mes;
                       cmd.Parameters.Add("@id_sector", SqlDbType.VarChar).Value = sector;
                       cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = operario;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();

                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.suministro_lectura =Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura =Fila["medidor_lectura"].ToString();
                                obj_entidad.fotourl = ruta + Fila["fotourl"].ToString();

                                obj_entidad.fecha = Fila["fecha"].ToString();
                                obj_entidad.hora = Fila["hora"].ToString();
                                obj_entidad.latitud = Fila["latitud"].ToString();
                                obj_entidad.longitud = Fila["longitud"].ToString();   

                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<ResultadoLecturas_E> Capa_Dato_Get_NotasOperario(string fechaAsignacion, int idServicio, int operario)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
    
               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_RESULTADO_LECTURAS_NOTAS", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       
                       cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                       cmd.Parameters.Add("@IdServicio", SqlDbType.Int).Value = idServicio;
                       cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = operario;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();
                               
                                obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                                obj_entidad.desc_operario = Fila["desc_operario"].ToString();
                                obj_entidad.fecha =  Fila["fecha"].ToString();
                                obj_entidad.observacion = Fila["observacion"].ToString();

                               ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }


       public List<ResultadoLecturas_E> Capa_Dato_Get_Observaciones(string fechaAsignacion, int idServicio, int operario, int id_supervisor, int id_operario_supervisor)
        {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<ResultadoLecturas_E> ListServicio = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_RESULTADO_LECTURAS_OPERARIO_OBSERVACION", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;

                       cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                       cmd.Parameters.Add("@IdServicio", SqlDbType.Int).Value = idServicio;
                       cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = operario;

                        DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();
                                
                                obj_entidad.codigo = Fila["codigo"].ToString();
                                obj_entidad.observacion = Fila["observacion"].ToString();
                                obj_entidad.total = Fila["total"].ToString();

                                obj_entidad.conFoto = Fila["conFoto"].ToString();
                                obj_entidad.sinFoto = Fila["sinFoto"].ToString();
                                //obj_entidad.porcentaje = Fila["porcentaje"].ToString();
                                obj_entidad.porcentaje = string.IsNullOrEmpty(Fila["porcentaje"].ToString()) ? obj_entidad.porcentaje = null : Convert.ToDecimal(Fila["porcentaje"].ToString()).ToString("#,##0.00");

                                ListServicio.Add(obj_entidad);
                           }
                       }
                   }
               }
               return ListServicio;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }
 

        public string Capa_Dato_Guardar_NotasOperario(string fechaAsignacion, int idServicio, int operario, string observacion, int usuario)
       {
           string Resultado = "";
           try
           {
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_I_RESULTADO_LECTURAS_NOTAS", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;

                       cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                       cmd.Parameters.Add("@IdServicio", SqlDbType.Int).Value = idServicio;
                       cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = operario;
                       cmd.Parameters.Add("@comentario", SqlDbType.VarChar).Value = observacion;
                       cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;
                       cmd.ExecuteNonQuery();
                   }
               }
               Resultado = "OK";
           }
           catch (Exception ex)
           {
               Resultado = ex.Message;
           }
           return Resultado;
       }


       public List<ResultadoLecturas_E> Capa_Dato_Get_ultimoSector()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<ResultadoLecturas_E> ListDato = new List<ResultadoLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_RESULTADO_LECTURA_MAX_SECTOR", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               ResultadoLecturas_E obj_entidad = new ResultadoLecturas_E();
                               obj_entidad.Sector_lectura = Fila["Sector_lectura"].ToString();
 
                               ListDato.Add(obj_entidad);
                           }
                       }
                   }
               }

               return ListDato;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }


        public string Capa_Dato_Descargar_efectividad(int Local , int servicio, string fecha, int id_usuario, int id_supervisor, int id_operario_supervisor)
        {
            var resultado = "";
            var nombreFile = "";
            var nombreExcel = "";

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                nombreFile = id_usuario + "_Efectividad.xlsx";
                nombreExcel = "Efectividad";
  

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_RESULTADO_LECTURAS_EFECTIVIDAD", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@IdServicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = 0;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = Local;

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
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

            try
            {
                _ruta = HttpContext.Current.Server.MapPath("~/Temp/" + nombreFile);

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

                    for (int i = 1; i <= 8; i++)
                    {
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    oWs.Cells[1, 1].Value = "Año";
                    oWs.Cells[1, 2].Value = "Mes";
                    oWs.Cells[1, 3].Value = "Codigo";
                    oWs.Cells[1, 4].Value = "Operario";

                    oWs.Cells[1, 5].Value = "Total";
                    oWs.Cells[1, 6].Value = "Efectivo";
                    oWs.Cells[1, 7].Value = "No-efectivo";
                    oWs.Cells[1, 8].Value = "%";

                    int ac = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        ac += 1;
                        for (int j = 1; j <= 8; j++)
                        {
                            oWs.Cells[_fila, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }


                        oWs.Cells[_fila, 1].Value = oBj["anio"].ToString();
                        oWs.Cells[_fila, 2].Value = oBj["mes"].ToString();
                        oWs.Cells[_fila, 3].Value = oBj["codigo"].ToString();

                        oWs.Cells[_fila, 4].Value = oBj["observacion"].ToString();

                        oWs.Cells[_fila, 5].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 5].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right; // alinear texto
                        oWs.Cells[_fila, 5].Value = oBj["total"].ToString();

                        oWs.Cells[_fila, 6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 6].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right; // alinear texto
                        oWs.Cells[_fila, 6].Value = oBj["conFoto"].ToString();

                        oWs.Cells[_fila, 7].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 7].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right; // alinear texto
                        oWs.Cells[_fila, 7].Value = oBj["sinFoto"].ToString();

                        oWs.Cells[_fila, 8].Style.Numberformat.Format = "#,##0.00";
                        oWs.Cells[_fila, 8].Value = Math.Round(Convert.ToDecimal(oBj["porcentaje"]), 2);
                        oWs.Cells[_fila, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 8].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right; // alinear texto

                        _fila++;
                    }

                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Column(1).Style.Font.Bold = true;

                    for (int k = 1; k <= 8; k++)
                    {
                        oWs.Column(k).AutoFit();
                    }
                    oEx.Save();
                }

                Res = "1|" + ruta_descarga + nombreFile;
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }
        


        public string Capa_Dato_Descargar_efectividad_detallado(int Local, int servicio, string fecha, int id_usuario)
        {
            var resultado = "";
            var nombreFile = "";
            var nombreExcel = "";

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                nombreFile = id_usuario + "_Efectividad_Detallado.xlsx";
                nombreExcel = "Efectividad_Detallado";
                
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_S_EFECTIVIDAD_DETALLADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@IdServicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = Local;

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
                                resultado = GenerarArchivoExcel_data_detallado(dt_detalle, nombreFile, nombreExcel);
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

        public string GenerarArchivoExcel_data_detallado(DataTable dt_detalles, string nombreFile, string nombreExcel)
        {
            string _ruta = "";
            string Res = "";
            int _fila = 2;
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

            try
            {
                _ruta = HttpContext.Current.Server.MapPath("~/Temp/" + nombreFile);

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
                    //for (int i = 1; i <= 8; i++)
                    //{
                    //    oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    //}

                    oWs.Cells[1, 1].Value = "Año";
                    oWs.Cells[1, 2].Value = "Mes";
                    oWs.Cells[1, 3].Value = "Dia";
                    oWs.Cells[1, 4].Value = "Codigo";
                    oWs.Cells[1, 5].Value = "Operario";
                    oWs.Cells[1, 6].Value = "Total";


                    int ac = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        ac += 1;
                        //for (int j = 1; j <= 6; j++)
                        //{
                        //    oWs.Cells[_fila, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        //}
                        oWs.Cells[_fila, 1].Value = oBj["anio"].ToString();
                        oWs.Cells[_fila, 2].Value = oBj["mes"].ToString();
                        oWs.Cells[_fila, 3].Value = oBj["dia"].ToString();

                        oWs.Cells[_fila, 4].Value = oBj["id_Operario_Lectura"].ToString();
                        oWs.Cells[_fila, 5].Value = oBj["operario"].ToString();

                        oWs.Cells[_fila, 6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 6].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right; // alinear texto
                        oWs.Cells[_fila, 6].Style.Numberformat.Format = "#,##0.00";
                        oWs.Cells[_fila, 6].Value = Convert.ToDecimal(oBj["cant_reg"]);

                        _fila++;
                    }

                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Column(1).Style.Font.Bold = true;

                    for (int k = 1; k <= 6; k++)
                    {
                        oWs.Column(k).AutoFit();
                    }
                    oEx.Save();
                }

                Res = "1|" + ruta_descarga +  nombreFile;
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }



    }
}
