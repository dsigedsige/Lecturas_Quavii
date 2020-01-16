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
namespace DSIGE.Dato
{
    public class Cls_Dato_CambioEstado
    {
       string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
 
       public string Capa_Dato_Set_CambioEstado_Masivo(int local, int servicio, int operario , string fecha_asigna,  int usuario)
       {
           string Resultado = "";
           try
           {
               //obteniendo Datos de Facturacion
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();
                    //using (SqlCommand cmd = new SqlCommand("SP_U_CAMBIO_ESTADO_MASIVO", cn))
                    using (SqlCommand cmd = new SqlCommand("SP_U_CAMBIO_ESTADO_MASIVO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = operario;
                        cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fecha_asigna;
                        cmd.Parameters.Add("@Usuario", SqlDbType.Int).Value = usuario;
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
        
        public string Capa_Dato_GenerarAnular_Masivo(int servicio, string fecha_Asigna, int estado, int operario, int usuario)
        {
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_U_CAMBIO_ESTADO_DAR_BAJA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
 
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha_Asigna", SqlDbType.VarChar).Value = fecha_Asigna;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@operario", SqlDbType.Int).Value = operario;
                        cmd.Parameters.Add("@Usuario", SqlDbType.Int).Value = usuario;
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

        


       public List<Cambio_Estado_Masivo_E> Capa_Dato_Get_ListaLocales()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<Cambio_Estado_Masivo_E> Listlocal = new List<Cambio_Estado_Masivo_E>();
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
                               Cambio_Estado_Masivo_E obj_entidad = new Cambio_Estado_Masivo_E();

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


       public List<Cambio_Estado_Masivo_E> Capa_Dato_Get_ListaServicio()
       {
           try
           {
               cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

               List<Cambio_Estado_Masivo_E> ListServicio = new List<Cambio_Estado_Masivo_E>();
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
               {
                   cn.Open();

                   using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_III", cn))
                   {
                       cmd.CommandTimeout = 0;
                       cmd.CommandType = CommandType.StoredProcedure;
                       DataTable dt_detalle = new DataTable();
                       using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                       {
                           da.Fill(dt_detalle);
                           foreach (DataRow Fila in dt_detalle.Rows)
                           {
                               Cambio_Estado_Masivo_E obj_entidad = new Cambio_Estado_Masivo_E();

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


        public List<Cambio_Estado_Masivo_E> Capa_Dato_Get_ListaServicio_usuario(int id_usuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cambio_Estado_Masivo_E> ListServicio = new List<Cambio_Estado_Masivo_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_ANULAR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cambio_Estado_Masivo_E obj_entidad = new Cambio_Estado_Masivo_E();

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




    }
}
