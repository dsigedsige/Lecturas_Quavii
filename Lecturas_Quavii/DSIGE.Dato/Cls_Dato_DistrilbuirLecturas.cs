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
using System.ComponentModel;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;
using System.Drawing;
using System.Web;
//using System.Web.Mail;
//using System.Net.Mail;
//using System.Net.Mime;
//using System.Net;


using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Data.OleDb;
using System.Threading;

namespace DSIGE.Dato
{
    public class Cls_Dato_DistrilbuirLecturas
    {
       string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
       OleDbConnection cn;
        
       private OleDbConnection ConectarExcel(string nomExcel)
       {
           cn = new OleDbConnection();
           try
           {
               cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + nomExcel + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
               cn.Open();
               return cn;
           }
           catch (Exception)
           {
               cn.Close();
               throw;
           }
       }

       public List<DistribuirLecturas_E> Capa_Dato_Get_ListaLocales()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> Listlocal = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_LOCALES", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();

                               obj_entidad.id_Local = Convert.ToInt32(Fila["id_Local"].ToString());
                               obj_entidad.nombre_local = Fila["nombre_local"].ToString();
                               Listlocal.Add(obj_entidad);
                           }
                       }
                   }
               }

               return Listlocal;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<DistribuirLecturas_E> Capa_Dato_Get_ListaServicio()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListServicio = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_IV", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();

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

       public List<DistribuirLecturas_E> Capa_Dato_Get_ListarInformacionLecturas(int local, string fechaAsigna, int servicio, string opcion, int id_supervisor, int id_operario_supervisor, string tipoCliente)
        {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                    //using (SqlCommand cmd = new SqlCommand("NEW_SP_S_DISTRIBUIR_LECTURAS", cn))
                    using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUIR_LECTURAS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;
                        cmd.Parameters.Add("@tipoCliente", SqlDbType.VarChar).Value = tipoCliente;

                        DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();

                               if (opcion =="D")
                               {
                                   if (String.IsNullOrEmpty(Fila["Codigo_operario"].ToString()))
                                   {
                                       obj_entidad.checkeado = false;
                                   }
                                   else
                                   {
                                       obj_entidad.checkeado = true;
                                   }
                                   obj_entidad.disabled = true;
                               }
                               if (opcion == "R")
                               {
                                   obj_entidad.checkeado = false;
                                   obj_entidad.disabled = false;
                               }
                        
                               obj_entidad.Unidad_Lecturas = Fila["Unidad_Lecturas"].ToString();
                               obj_entidad.Cantidad_Manzana = Fila["Cantidad_Manzana"].ToString();
                               obj_entidad.Cantidad = Fila["Cantidad"].ToString();
                               obj_entidad.Codigo_operario = Fila["Codigo_operario"].ToString();
                               obj_entidad.id_Local = Convert.ToInt32(Fila["id_Local"].ToString());
                               obj_entidad.fechaAsignacion_Lectura = Fila["fechaAsignacion_Lectura"].ToString();
                               obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.flag_relectura = Fila["flag_relectura"].ToString();
                                

                               ListData.Add(obj_entidad);
                           }
                       }
                   }
               }

               return ListData;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

 
       public List<DistribuirLecturas_E> Capa_Dato_Get_ListarInformacionLecturasDetalleDistribucion(int local, string fechaAsigna, int servicio, int id_operario, int zona, int sector)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUCION_LECTRURAS_LECTURAS_DETALLE", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                       cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                       cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                       cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                       cmd.Parameters.Add("@zona", SqlDbType.Int).Value = zona;
                       cmd.Parameters.Add("@sector", SqlDbType.Int).Value = sector;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();
                               obj_entidad.check = false;
                               obj_entidad.flag_operario = false;
                               obj_entidad.colorItem = "";
                               obj_entidad.id_Lectura = Fila["id_lectura"].ToString();
                               obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                               obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                               obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                               obj_entidad.nombreCliente_lectura = Fila["nombreCliente_lectura"].ToString();
                               //obj_entidad.RT = Fila["RT"].ToString();
                               //obj_entidad.S1 = Fila["S1"].ToString();
                               //obj_entidad.CC = Convert.ToString(Fila["CC"]);
                               //obj_entidad.CL = Fila["CL"].ToString();
                               obj_entidad.SECTOR = Fila["SECTOR"].ToString();
                               obj_entidad.ZONA = Fila["ZONA"].ToString();
                               obj_entidad.CANTIDAD = Convert.ToInt32(Fila["CANTIDAD"].ToString());
                               obj_entidad.FECHA = Fila["FECHA"].ToString();
                               obj_entidad.id_Local = Convert.ToInt32(Fila["id_Local"].ToString());

                               ListData.Add(obj_entidad);
                           }
                       }
                   }
               }

               return ListData;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }


       public List<DistribuirLecturas_E> Capa_Dato_Get_ListarInformacionLecturasDetalle(int id_local, string fechaAsigna, int id_servicio, string unidad_Lectura, int id_operario, string opcion)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
               int indice  =0;

               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUCION_RELECTURA_DETALLE_V2", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                       cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                       cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = id_servicio;
                       cmd.Parameters.Add("@unidad_Lectura", SqlDbType.VarChar).Value = unidad_Lectura;
                       cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                       cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();

                                obj_entidad.check = false;
                                obj_entidad.flag_operario = false;
                                obj_entidad.colorItem = "";
                                obj_entidad.index = indice + 1;       
                                obj_entidad.Unidad_Lecturas = Fila["Unidad_Lecturas"].ToString();                         
                                obj_entidad.Codigo_operario = Fila["Codigo_operario"].ToString();
                                obj_entidad.Manzana = Fila["Manzana"].ToString();  

                               ListData.Add(obj_entidad);
                               indice = indice + 1;

                           }
                       }
                   }
               }

               return ListData;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

        public List<DistribuirLecturas_E> Capa_Dato_Get_ListarInformacionLecturasDetalle_general(int id_local, string fechaAsigna, int id_servicio, string unidad_Lectura, int id_operario, string opcion)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
                int indice = 0;

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("PROC_S_DISTRIBUCION_RELECTURA_GENERAL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@unidad_Lectura", SqlDbType.VarChar).Value = unidad_Lectura;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();

                                obj_entidad.check = false;
                                obj_entidad.flag_operario = false;
                                obj_entidad.colorItem = "";
                                obj_entidad.index = indice + 1;
                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.Unidad_Lecturas = Fila["Unidad_Lecturas"].ToString();
                                obj_entidad.Codigo_operario = Fila["Codigo_operario"].ToString();
                                obj_entidad.Manzana = Fila["Manzana"].ToString();

                                ListData.Add(obj_entidad);
                                indice = indice + 1;

                            }
                        }
                    }
                }

                return ListData;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




        public string Capa_Dato_Set_ReasignacionTrabajos(List<string> ListaTrabajos, int id_usuario, string FechaMovil, string servicio)
       {
           string Resultado = "";
           string RT = "";
           string S1 = "";
           string CL = "";
           string SECTOR = "";
           string ZONA = "";
           string FECHA = "";
           string LOCAL = "";

           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   for (int i = 0; i < ListaTrabajos.Count; i++)
                   {
                       string[] Obj_parametros = ListaTrabajos[i].Split('|');
                       RT = Obj_parametros[0].ToString();
                       S1 = Obj_parametros[1].ToString();
                       CL = Obj_parametros[2].ToString();
                       SECTOR = Obj_parametros[3].ToString();
                       ZONA = Obj_parametros[4].ToString();
                       FECHA = Obj_parametros[5].ToString();
                       LOCAL = Obj_parametros[6].ToString();

                       //using (SqlCommand cmd = new SqlCommand("SP_I_REASIGNAR_LECTURAS", cn))
                       using (SqlCommand cmd = new SqlCommand("SP_I_REASIGNAR_LECTURAS_NEW", cn))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;
                           cmd.Parameters.Add("@RT", SqlDbType.Int).Value = RT;
                           cmd.Parameters.Add("@S1", SqlDbType.Int).Value = S1;
                           cmd.Parameters.Add("@CL", SqlDbType.Int).Value = CL;
                           cmd.Parameters.Add("@SECTOR", SqlDbType.VarChar).Value = SECTOR;
                           cmd.Parameters.Add("@ZONA", SqlDbType.VarChar).Value = ZONA;
                           cmd.Parameters.Add("@FECHA", SqlDbType.VarChar).Value = FECHA;
                           cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                           cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = LOCAL;
                           cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                           cmd.Parameters.Add("@id_servicio", SqlDbType.VarChar).Value = servicio;
                           cmd.ExecuteNonQuery();
                       }
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
      

       public string Capa_Dato_Set_ReasignacionTrabajosDetalle(List<string> ListaTrabajos, int id_usuario, string FechaMovil, string servicio)
       {
           string Resultado = "";
           string RT = "";
           string S1 = "";
           string CL = "";
           string SECTOR = "";
           string ZONA = "";
           string FECHA = "";
           string LOCAL = "";
           string id_lectura = "";

           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   for (int i = 0; i < ListaTrabajos.Count; i++)
                   {
                       string[] Obj_parametros = ListaTrabajos[i].Split('|');
                       RT = Obj_parametros[0].ToString();
                       S1 = Obj_parametros[1].ToString();
                       CL = Obj_parametros[2].ToString();
                       SECTOR = Obj_parametros[3].ToString();
                       ZONA = Obj_parametros[4].ToString();
                       FECHA = Obj_parametros[5].ToString();
                       LOCAL = Obj_parametros[6].ToString();
                       id_lectura= Obj_parametros[7].ToString();
                       
                       //using (SqlCommand cmd = new SqlCommand("SP_I_REASIGNAR_LECTURAS_DETALLE", cn))
                       using (SqlCommand cmd = new SqlCommand("SP_I_REASIGNAR_LECTURAS_DETALLE_NEW", cn))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;

                           cmd.Parameters.Add("@RT", SqlDbType.Int).Value = RT;
                           cmd.Parameters.Add("@S1", SqlDbType.Int).Value = S1;
                           cmd.Parameters.Add("@CL", SqlDbType.Int).Value = CL;
                           cmd.Parameters.Add("@SECTOR", SqlDbType.VarChar).Value = SECTOR;
                           cmd.Parameters.Add("@ZONA", SqlDbType.VarChar).Value = ZONA;
                           cmd.Parameters.Add("@FECHA", SqlDbType.VarChar).Value = FECHA;
                           cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                           cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = LOCAL;
                           cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                           cmd.Parameters.Add("@id_Lectura", SqlDbType.VarChar).Value = id_lectura;
                           cmd.Parameters.Add("@id_servicio", SqlDbType.VarChar).Value = servicio;
                           
                           cmd.ExecuteNonQuery();
                       }
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

       public List<DistribuirLecturas_E> Capa_Dato_Get_ValidarCodigoSector(int RT, string FECHA)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUCION_LECTURAS_VALIDAR_SECTOR", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@RT", SqlDbType.Int).Value = RT;
                       cmd.Parameters.Add("@FECHA", SqlDbType.VarChar).Value = FECHA;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();
                               obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                               ListData.Add(obj_entidad);
                           }
                       }
                   }
               }

               return ListData;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }
        
       public string Capa_Dato_Set_EnvioMovil(string FechaEnvio, int id_usuario, int  TipoServicio, string Fechaimportacion)
       {

           string Resultado = "";
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               //obteniendo Datos de Facturacion
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   //Informacion de Guias Cabeceras
                   using (SqlCommand cmd = new SqlCommand("SP_I_DISTRIBUCION_TRABAJOS_ENVIO_MOVIL", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                       cmd.Parameters.Add("@FechaImportacion ", SqlDbType.VarChar).Value = Fechaimportacion;
                       cmd.Parameters.Add("@fechaMovil", SqlDbType.VarChar).Value = FechaEnvio;
                       cmd.Parameters.Add("@TipoServicio", SqlDbType.Int).Value = TipoServicio;
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

       public string Capa_Dato_Set_ProcesarConsumo(string FechaEnvio, int id_usuario)
       {

           string Resultado = "";
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               //obteniendo Datos de Facturacion
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   //Informacion de Guias Cabeceras
                   using (SqlCommand cmd = new SqlCommand("DSIGE_Proceso_Lecturas_MinMax", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@Usuario", SqlDbType.Int).Value = id_usuario;
                       cmd.Parameters.Add("@FechaAsignacion ", SqlDbType.VarChar).Value = FechaEnvio;
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

        public string Capa_Dato_Proceso_lectura_pendiente(string Fechaasignacion, int id_usuario, string fecha_relectura)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    //Informacion de Guias Cabeceras
                    using (SqlCommand cmd = new SqlCommand("PROC_S_GENERACION_LECTURAS_PENDIENTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_asignacion ", SqlDbType.VarChar).Value = Fechaasignacion;
                        cmd.Parameters.Add("@fecha_relectura ", SqlDbType.VarChar).Value = fecha_relectura;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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


        

       public List<DistribuirLecturas_E> Capa_Dato_Get_ListarInformacionLecturas(int local, string anio, string mes)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUCION_RELACION_PLANTILLAS", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                       cmd.Parameters.Add("@anio", SqlDbType.Int).Value = Convert.ToInt32(anio);
                       cmd.Parameters.Add("@mes", SqlDbType.Int).Value = Convert.ToInt32(mes);
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                          foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();
                               obj_entidad.id_Local =Convert.ToInt32(Fila["id_Local"].ToString());
                               obj_entidad.fecha_Asignacion = Fila["fecha_Asignacion"].ToString();
                               ListData.Add(obj_entidad);
                           }
                       }
                   }
               }

               return ListData;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }

       public List<DistribuirLecturas_E> Capa_Dato_Get_ValidarOperario(int id_operario)
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUIR_LECTURA_VALIDAR_OPERARIO", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = id_operario;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();
                               obj_entidad.CANT_REGISTROS = Convert.ToInt32(Fila["CANT_REGISTROS"].ToString());
                               ListData.Add(obj_entidad);
                           }
                       }
                   }
               }

               return ListData;
           }
           catch (Exception ex)
           {
               throw ex;
           }
       }
        
       public string Capa_Dato_Get_DistribuirOrdenes(int id_local, string fechaAsignacion, int id_servicio, string unidad_lectura, int id_operario)
       {
           string Resultado = "";
           try
           {
               //obteniendo Datos de Facturacion
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUIR_LECTURA_POR_OPERARIO", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                       cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = id_servicio;
                       cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsignacion;
                       cmd.Parameters.Add("@unidad_lectura", SqlDbType.VarChar).Value = unidad_lectura;
                       cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
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
        
       public string Capa_Dato_Set_Distribucion_EnviarMovil(List<Ordenes_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil,  string opcion, int id_usuario)
       {
           string Resultado = "";
           try
           {

               PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(Ordenes_E));
               DataTable table = new DataTable();

               foreach (PropertyDescriptor prop in properties) {
                   table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
               }

               foreach (Ordenes_E item in ListaOrdenes)
               {
                   DataRow row = table.NewRow();
                   foreach (PropertyDescriptor prop in properties)
                       row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                   table.Rows.Add(row);
               }


               using (SqlConnection con = new SqlConnection(cadenaCnx))
               {
                   con.Open();
                   ///eliminando registros del usuario
                   using (SqlCommand cmd = new SqlCommand("SP_D_DISTRIBUCION_ORDENES", con))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                       cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                       cmd.ExecuteNonQuery();
                   }
                   using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                   {
                       bulkCopy.BatchSize = 500;
                       bulkCopy.NotifyAfter = 1000;
                       bulkCopy.DestinationTableName = "T_DistribuirOrdenes";
                       bulkCopy.WriteToServer(table);

                       //Actualizando campos 
                       string Sql = "UPDATE T_DistribuirOrdenes SET id_local='" + id_local + "' ,  FechaAsigna='" + FechaAsigna + "' ,  id_servicio='" + servicio + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE opcion='" + opcion + "' AND  id_usuario IS NULL";
                       using (SqlCommand cmd = new SqlCommand(Sql, con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.Text;
                           cmd.ExecuteNonQuery();
                       }
                       //// insertando en la tabla Lecturas
                       using (SqlCommand cmd = new SqlCommand("SP_I_DISTRIBUCION_ENVIO_MOVIL", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;

                           cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                           cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                           cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                           cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                           cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                           cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

                           cmd.ExecuteNonQuery();
                       }
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

       public string Capa_Dato_Set_Reasignacion_EnviarMovil(List<Ordenes_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int id_usuario)
       {
           string Resultado = "";
           try
           {
               PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(Ordenes_E));
               DataTable table = new DataTable();

               foreach (PropertyDescriptor prop in properties)
               {
                   table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
               }

               foreach (Ordenes_E item in ListaOrdenes)
               {
                   DataRow row = table.NewRow();
                   foreach (PropertyDescriptor prop in properties)
                       row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                   table.Rows.Add(row);
               }


               using (SqlConnection con = new SqlConnection(cadenaCnx))
               {
                   con.Open();
                   ///eliminando registros del usuario
                   using (SqlCommand cmd = new SqlCommand("SP_D_DISTRIBUCION_ORDENES", con))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                       cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                       cmd.ExecuteNonQuery();
                   }
                   using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                   {
                       bulkCopy.BatchSize = 500;
                       bulkCopy.NotifyAfter = 1000;
                       bulkCopy.DestinationTableName = "T_DistribuirOrdenes";
                       bulkCopy.WriteToServer(table);

                       //Actualizando campos 
                       string Sql = "UPDATE T_DistribuirOrdenes SET id_local='" + id_local + "' ,  FechaAsigna='" + FechaAsigna + "' ,  id_servicio='" + servicio + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE opcion='" + opcion + "' AND  id_usuario IS NULL";
                       using (SqlCommand cmd = new SqlCommand(Sql, con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.Text;
                           cmd.ExecuteNonQuery();
                       }
                       // update en la tabla Lecturas
                       using (SqlCommand cmd = new SqlCommand("SP_I_REASIGNACION_ENVIO_MOVIL", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;

                           cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                           cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                           cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                           cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                           cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                           cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

                           cmd.ExecuteNonQuery();
                       }
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


       public string Capa_Dato_Set_Distribucion_EnviarMovil_detallado(List<OrdenesDetalle_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int id_usuario, string Flag_detallado)
       {
           string Resultado = "";
           try
           {

               PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(OrdenesDetalle_E));
               DataTable table = new DataTable();

               foreach (PropertyDescriptor prop in properties)
               {
                   table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
               }

               foreach (OrdenesDetalle_E item in ListaOrdenes)
               {
                   DataRow row = table.NewRow();
                   foreach (PropertyDescriptor prop in properties)
                       row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                   table.Rows.Add(row);
               }


               using (SqlConnection con = new SqlConnection(cadenaCnx))
               {
                   con.Open();
                   ///eliminando registros del usuario
                   using (SqlCommand cmd = new SqlCommand("SP_D_DISTRIBUCION_ORDENES_DETALLE", con))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                       cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                       cmd.ExecuteNonQuery();
                   }
                   using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                   {
                       bulkCopy.BatchSize = 500;
                       bulkCopy.NotifyAfter = 1000;
                       bulkCopy.DestinationTableName = "T_DistribuirOrdenes_detalle";
                       bulkCopy.WriteToServer(table);

                       //Actualizando campos 
                       string Sql = "UPDATE T_DistribuirOrdenes_detalle SET id_local='" + id_local + "' ,  FechaAsigna='" + FechaAsigna + "' ,  id_servicio='" + servicio + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE opcion='" + opcion + "' AND  id_usuario IS NULL";
                       using (SqlCommand cmd = new SqlCommand(Sql, con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.Text;
                           cmd.ExecuteNonQuery();
                       }
                       //// insertando en la tabla Lecturas
                       using (SqlCommand cmd = new SqlCommand("SP_I_DISTRIBUCION_ENVIO_MOVIL_DETALLADO", con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;

                           cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                           cmd.Parameters.Add("@flag_detallado", SqlDbType.VarChar).Value = Flag_detallado;
                           cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                           cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                           cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                           cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                           cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

                           cmd.ExecuteNonQuery();
                       }
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

       public string Capa_Dato_Set_Reasignacion_EnviarMovil_detallado(List<OrdenesDetalle_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int id_usuario, string Flag_detallado)
       {
           string Resultado = "";
           try
           {
               PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(OrdenesDetalle_E));
               DataTable table = new DataTable();

               foreach (PropertyDescriptor prop in properties)
               {
                   table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
               }

               foreach (OrdenesDetalle_E item in ListaOrdenes)
               {
                   DataRow row = table.NewRow();
                   foreach (PropertyDescriptor prop in properties)
                       row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                   table.Rows.Add(row);
               }


               using (SqlConnection con = new SqlConnection(cadenaCnx))
               {
                   con.Open();
                    ///eliminando registros del usuario
                    //using (SqlCommand cmd = new SqlCommand("SP_D_DISTRIBUCION_ORDENES_DETALLE_V2", con))
                    using (SqlCommand cmd = new SqlCommand("SP_D_DISTRIBUCION_ORDENES_DETALLE", con))
                    {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                       cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                       cmd.ExecuteNonQuery();
                   }
                   using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                   {
                       bulkCopy.BatchSize = 500;
                       bulkCopy.NotifyAfter = 1000;
                       bulkCopy.DestinationTableName = "T_DistribuirOrdenes_detalle";
                       bulkCopy.WriteToServer(table);

                       //Actualizando campos 
                       string Sql = "UPDATE T_DistribuirOrdenes_detalle SET id_local='" + id_local + "' ,  FechaAsigna='" + FechaAsigna + "' ,  id_servicio='" + servicio + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE opcion='" + opcion + "' AND  id_usuario IS NULL";
                       using (SqlCommand cmd = new SqlCommand(Sql, con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.Text;
                           cmd.ExecuteNonQuery();
                       }
                        // update en la tabla Lecturas
                        using (SqlCommand cmd = new SqlCommand("SP_I_REASIGNACION_ENVIO_MOVIL_DETALLE", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = opcion;
                            cmd.Parameters.Add("@flag_detallado", SqlDbType.VarChar).Value = Flag_detallado;
                            cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                            cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                            cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                            cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

                            cmd.ExecuteNonQuery();
                        }
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

       public string Capa_Dato_Set_ProcesarConsumo(string FechaEnvio, int id_usuario, int id_servicio)
       {
           string Resultado = "";
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               //obteniendo Datos de Facturacion
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   //Informacion de Guias Cabeceras
                   using (SqlCommand cmd = new SqlCommand("DSIGE_Proceso_Lecturas_MinMax", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@FechaAsignacion ", SqlDbType.VarChar).Value = FechaEnvio;
                       cmd.Parameters.Add("@id_servicio ", SqlDbType.VarChar).Value = id_servicio;
                       cmd.Parameters.Add("@Usuario", SqlDbType.Int).Value = id_usuario;
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


        public string Capa_Dato_Set_ProcesarConsumo_new(string FechaEnvio, int id_usuario, int id_servicio)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                //obteniendo Datos de Facturacion
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    //Informacion de Guias Cabeceras
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Proceso_Lecturas_MinMax_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsignacion ", SqlDbType.VarChar).Value = FechaEnvio;
                        cmd.Parameters.Add("@id_servicio ", SqlDbType.VarChar).Value = id_servicio;
                        cmd.Parameters.Add("@Usuario", SqlDbType.Int).Value = id_usuario;
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

        public string Capa_Dato_DetalleLecturas_Correo(List<string> ListaTrabajos, int id_usuario, int id_local, string FechaAsigna, int servicio)
       {
           string mensaje = "";
           string Res = "";
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
               List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                   List<Archivo_E> listArchivos = new List<Archivo_E>();
   
                   /// generando los archivos excel
                   for (int i = 0; i < ListaTrabajos.Count(); i++)
                   {          
                       using (SqlCommand cmd = new SqlCommand("SP_S_DISTRIBUCION_RELECTURA_CORREO", cn))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.StoredProcedure;
                           cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                           cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                           cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                           cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value =  ListaTrabajos[i];  

                           DataTable dt_detalle = new DataTable();
                           using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                           {
                               da.Fill(dt_detalle);
                               if (dt_detalle.Rows.Count <= 0)
                               {
                                   mensaje = "0|No hay informacion disponible";
                               }
                               else
                               {
                                   Archivo_E obj_entidad = new Archivo_E();
                                   Res = GenerarArchivoExcel(dt_detalle, ListaTrabajos[i], servicio);

                                   string[] ArchivosExcel = Res.Split('|');                                   
                                   obj_entidad.email = ArchivosExcel[0];
                                   obj_entidad.ruta = ArchivosExcel[1];
                                   obj_entidad.nombreFile = ArchivosExcel[2];                                   

                                   if (obj_entidad.email !=null || obj_entidad.email !="")
                                   {
                                       listArchivos.Add(obj_entidad);
                                   }         
                               }
                           }
                       }
                   }

                    //---generando el envio de correo

                    for (int i = 0; i < listArchivos.Count(); i++)
                    {
                        mensaje = SendEmail(listArchivos[i].email, listArchivos[i].ruta, listArchivos[i].nombreFile);
                        Thread.Sleep(1000);
                    }
                }
            }
           catch (Exception ex)
           {
               mensaje = ex.Message;
           }
           return mensaje;
       }


       public string SendEmail(string correo, string rutaFile, string nombreFile)
       {
           string path;
           string pathExist;
           string mensaje = "";
           try
           {
               var body = "Listado de Lecturas Calidda";
               var message = new MailMessage();

               foreach (var curr_address in correo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
               {
                   message.To.Add(new MailAddress(curr_address));
               }
               //if (!string.IsNullOrEmpty(cc))
               //{
               //    message.To.Add(new MailAddress(cc));
               //}
               message.From = new MailAddress("lecturas.bureauveritas@gmail.com");
               message.Subject = "Suministros de Trabajo " + DateTime.Today.ToString("dd-MM-yyyy"); 
               message.Body = body;
               message.IsBodyHtml = true;
               message.Priority = MailPriority.Normal;

               path = rutaFile;
               Attachment attachment = new Attachment(path, MediaTypeNames.Application.Octet);
               attachment.Name = nombreFile;
               ContentDisposition disposition = attachment.ContentDisposition;
               disposition.CreationDate = System.IO.File.GetCreationTime(path);
               disposition.ModificationDate = System.IO.File.GetLastWriteTime(path);
               disposition.ReadDate = System.IO.File.GetLastAccessTime(path);

               message.Attachments.Add(attachment);
               using (var smtp = new SmtpClient())
               {
                   var credential = new NetworkCredential
                   {
                       UserName = "lecturas.bureauveritas@gmail.com",
                       Password = "Lecturas2019"
                   };
                    smtp.Host = "smtp.gmail.com";
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = credential;
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.Send(message);
               }
               mensaje = "OK";
           }
           catch (Exception ex)
           {
               mensaje = ex.Message;
                string name = correo;
           }
           return mensaje;
       }
        
       public string GenerarArchivoExcel_reclamos(DataTable dt_detalles, string id_operario, int idServices)
       {
            string nombreServicio = "";
            string _servidor ="";
            string _ruta = "";
            string Res = "";
            int _fila =2;
            var correo = "";
            try
            {
                if (idServices == 1)
                {
                    nombreServicio = "MAIL_LECTURAS_";
                }
                else
                {
                    nombreServicio = "MAIL_RECLAMOS_";
                }

                _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
                _ruta = HttpContext.Current.Server.MapPath("~/Temp/" + id_operario +"_"+ nombreServicio + _servidor);

                FileInfo _file = new FileInfo(_ruta);
                if (_file.Exists)
                {
                    _file.Delete();
                    _file = new FileInfo(_ruta);
                }


                using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
                {
                    Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("DatosSuministros");
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 9));

                    for (int i = 1; i <= 7; i++)
                    {
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    oWs.Cells[1, 1].Value = "#";
                    oWs.Cells[1, 2].Value = "Medidor Lecturas";
                    oWs.Cells[1, 3].Value = "Lecturas";
                    oWs.Cells[1, 4].Value = "Direccion Lectura";
                    oWs.Cells[1, 5].Value = "Manzana";
                    oWs.Cells[1, 6].Value = "Unidad Lectura";
                    oWs.Cells[1, 7].Value = "Foto";




                    int ac = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        ac += 1;
                        for (int j = 1; j <= 7; j++)
                        {
                            oWs.Cells[_fila, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }

                        oWs.Cells[_fila, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto  
                        oWs.Cells[_fila, 1].Value = ac;

                        correo = oBj["email_operario"].ToString();

                        oWs.Cells[_fila, 2].Value = oBj["medidor_lectura"].ToString();

                        oWs.Cells[_fila, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 3].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto
                        oWs.Cells[_fila, 3].Value = "";

                        oWs.Cells[_fila, 4].Value = oBj["direccion_lectura"].ToString();
                        oWs.Cells[_fila, 5].Value = oBj["Manzana"].ToString();
                        oWs.Cells[_fila, 6].Value = oBj["Unidad_Lecturas"].ToString();
                        oWs.Cells[_fila, 7].Value = oBj["foto"].ToString();

                        _fila++;
                    }
                    oWs.Cells.Style.Font.Size = 8; //letra tamaño  
                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Column(1).Style.Font.Bold = true;

                    for (int k = 1; k <= 7; k++)
                    {
                        oWs.Column(k).AutoFit();
                    }
                    oEx.Save();
                }

                Res = correo + "|" + _ruta + "|" + id_operario +"_"+ nombreServicio + _servidor;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Res;       
       }



        public string GenerarArchivoExcel(DataTable dt_detalles, string id_operario, int idServices)
        {
            string nombreServicio = "";
            string _servidor = "";
            string _ruta = "";
            string Res = "";
            int _fila = 2;
            var correo = "";
            try
            {
                if (idServices == 1)
                {
                    nombreServicio = "MAIL_LECTURAS_";
                }
                else
                {
                    nombreServicio = "MAIL_RECLAMOS_";
                }

                _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
                _ruta = HttpContext.Current.Server.MapPath("~/Temp/" + id_operario + "_" + nombreServicio + _servidor);

                FileInfo _file = new FileInfo(_ruta);
                if (_file.Exists)
                {
                    _file.Delete();
                    _file = new FileInfo(_ruta);
                }


                using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
                {
                    Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("DatosSuministros");
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 9));

                    for (int i = 1; i <= 10; i++)
                    {
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    oWs.Cells[1, 1].Value = "#";
                    oWs.Cells[1, 2].Value = "Medidor Lecturas";
                    oWs.Cells[1, 3].Value = "Lecturas";
                    oWs.Cells[1, 4].Value = "Direccion Lectura";
                    oWs.Cells[1, 5].Value = "Manzana";
                    oWs.Cells[1, 6].Value = "Unidad Lectura";
                    oWs.Cells[1, 7].Value = "Foto";

                    oWs.Cells[1, 8].Value = "Nota de lectura";
                    oWs.Cells[1, 9].Value = "Observacion";
                    oWs.Cells[1, 10].Value = "Suministro";


                    int ac = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        ac += 1;
                        for (int j = 1; j <= 10; j++)
                        {
                            oWs.Cells[_fila, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }

                        oWs.Cells[_fila, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto  
                        oWs.Cells[_fila, 1].Value = ac;

                        correo = oBj["email_operario"].ToString();

                        oWs.Cells[_fila, 2].Value = oBj["medidor_lectura"].ToString();

                        oWs.Cells[_fila, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 3].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto
                        oWs.Cells[_fila, 3].Value = "";

                        oWs.Cells[_fila, 4].Value = oBj["direccion_lectura"].ToString();
                        oWs.Cells[_fila, 5].Value = oBj["Manzana"].ToString();
                        oWs.Cells[_fila, 6].Value = oBj["Unidad_Lecturas"].ToString();
                        oWs.Cells[_fila, 7].Value = oBj["foto"].ToString();

                        oWs.Cells[_fila, 8].Value = "";
                        oWs.Cells[_fila, 9].Value = "";
                        oWs.Cells[_fila, 10].Value = oBj["suministro_lectura"].ToString();

                        _fila++;
                    }
                    oWs.Cells.Style.Font.Size = 8; //letra tamaño  
                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Column(1).Style.Font.Bold = true;

                    for (int k = 1; k <= 10; k++)
                    {
                        oWs.Column(k).AutoFit();
                    }
                    oEx.Save();
                }

                Res = correo + "|" + _ruta + "|" + id_operario + "_" + nombreServicio + _servidor;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Res;
        }


 

        public string Capa_Dato_GuardarArchivo(HttpPostedFileBase file, int idlocal, string fechaAsignacion, int idServicio, int idusuario , string fecha_lectura)
        {
           DataTable dt = new DataTable();
           string mensaje = "";
           try
           {
               string fechaActual = String.Format("{0:ddMMyyyy_hhmmss}", DateTime.Now);
               string nomExcel = idusuario + "_" + fechaActual + "_" + file.FileName;
               string fileLocation = System.Web.Hosting.HostingEnvironment.MapPath("~/Upload") + "\\" + nomExcel;
               file.SaveAs(fileLocation);

               string sql = "SELECT * FROM [DatosSuministros$]";

               OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
               da.SelectCommand.CommandType = CommandType.Text;
               da.Fill(dt);
               cn.Close();

               using (SqlConnection con = new SqlConnection(cadenaCnx))
               {
                   con.Open();
                   //eliminando registros del usuario
                   using (SqlCommand cmd = new SqlCommand("SP_D_TEMPORAL_SUMINISTRO_OPERARIO_V2", con))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = idusuario;
                       cmd.ExecuteNonQuery();
                   }

                   //guardando al informacion de la importacion
                   using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                   {
                       bulkCopy.BatchSize = 500;
                       bulkCopy.NotifyAfter = 1000;
                       bulkCopy.DestinationTableName = "TEMPORAL_SUMINISTRO_OPERARIO_V2";
                       bulkCopy.WriteToServer(dt);

                       //Actualizando campos 
                       string Sql = "UPDATE TEMPORAL_SUMINISTRO_OPERARIO_V2 SET  nombreArchivo='" + file.FileName + "',   loc_id ='" + idlocal + "' , id_usuario_exportador='" + idusuario + "', fechaAsignacion='" + fechaAsignacion + "'  WHERE id_usuario_exportador IS NULL    ";

                       using (SqlCommand cmd = new SqlCommand(Sql, con))
                       {
                           cmd.CommandTimeout = 0;
                           cmd.CommandType = CommandType.Text;
                           cmd.ExecuteNonQuery();
                       }
                   }

                   //validando los codigos de observaciones
                   var validarCod = "";
                   using (SqlCommand cmd = new SqlCommand("SP_S_SUMINISTRO_MANUAL_OBS_V2", con))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = idusuario;

                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter dadapter = new SqlDataAdapter(cmd))
                       {
                           dadapter.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               DistribuirLecturas_E obj_entidad = new DistribuirLecturas_E();
                               if (string.IsNullOrEmpty(Fila["cod_observacion"].ToString()))
                               {
                                   validarCod = "";
                               }
                               else {
                                   validarCod = Fila["cod_observacion"].ToString();
                               } 
                           }
                       }
                   }

                   if (validarCod == "")
                   {
                       int Cnssql = 0;
                        using (SqlCommand cmd = new SqlCommand("SP_U_SUMINISTROS_MANUAL_V2", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@idlocal", SqlDbType.Int).Value = idlocal;
                            cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                            cmd.Parameters.Add("@idServicio", SqlDbType.Int).Value = idServicio;
                            cmd.Parameters.Add("@idusuario", SqlDbType.Int).Value = idusuario;
                            cmd.Parameters.Add("@fecha_lectura", SqlDbType.VarChar).Value = fecha_lectura;
                            Cnssql = cmd.ExecuteNonQuery();
                        }

                        if (Cnssql == 0)
                       {
                           mensaje =  "1|No se encontro información para Actualizar, ó verifique los filtros (Fecha asignacion, Servicio)";
                       }
                       else
                       {
                           mensaje = "2|OK";
                       }
                     
                   }
                   else {
                       mensaje ="0|" + validarCod;
                   }            
               }
           }
           catch (Exception ex)
           {
               cn.Close();
               mensaje = "-1|" + ex.Message;
           }
           finally
           {
               cn.Close();
           }

           return mensaje;
       }
        

    }
}
