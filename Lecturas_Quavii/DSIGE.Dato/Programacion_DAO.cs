using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Dato
{
   public class Programacion_DAO
    {

        public object capa_dato_get_Suministros(int Local , string FechaAsiga, int servicio, int estado, string distrito, string horaInicial, string horaFinal)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_PROGRAMACION_SUMINISTROS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = Local;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@distrito", SqlDbType.VarChar).Value = distrito;
                        cmd.Parameters.Add("@horaInicial", SqlDbType.VarChar).Value = horaInicial;
                        cmd.Parameters.Add("@horaFinal", SqlDbType.VarChar).Value = horaFinal;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                       }
                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            return dt_detalle;
        }

        public object capa_dato_get_Distrito(string fechaAsignacion, int servicio)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_PROGRAMACION_SUMINISTROS_DISTRITO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            return dt_detalle;
        }



        public object capa_dato_get_Suministros_sinGps(int Local, string FechaAsiga, int servicio, int estado, string distrito, string horaInicial, string horaFinal)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_PROGRAMACION_SUMINISTROS_SIN_GPS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = Local;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@distrito", SqlDbType.VarChar).Value = distrito;
                        cmd.Parameters.Add("@horaInicial", SqlDbType.VarChar).Value = horaInicial;
                        cmd.Parameters.Add("@horaFinal", SqlDbType.VarChar).Value = horaFinal;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            return dt_detalle;
        }



        



        public object capa_dato_get_ListandoAsignados(string FechaAsiga, int servicio)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("DSIGE_RESUMEN_CORTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            return dt_detalle;
        }



        public object capa_dato_get_Suministros_Operario_Gps(string FechaAsiga, int servicio, int estado)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_PROGRAMACION_SUMINISTROS_OPERARIO_GPS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = FechaAsiga;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            return dt_detalle;
        }



        public string capa_dato_set_actualizarOperario(string obj_cortes, string fecha_asignacion, int servicio, int operario, string fecha_movil, int usuario)
        {
            var resultado = "";
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_U_PROGRAMACION_SUMINISTROS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@obj_cortes", SqlDbType.VarChar).Value = obj_cortes;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@operario", SqlDbType.Int).Value = operario;
                        cmd.Parameters.Add("@fecha_movil", SqlDbType.VarChar).Value = fecha_movil;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                        resultado = "OK";
                    }
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            return resultado;
        }

    }
}
