using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data;

namespace DSIGE.Dato
{
    public class Cls_Dato_Evento_Movil_Operario
    {

        //listado

        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;


        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Dato_Locales()
        {
            try
            {

                List<Cls_Entidad_Evento_Movil_Operario> obj_List_Locales = new List<Cls_Entidad_Evento_Movil_Operario>();
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

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Evento_Movil_Operario Entidad = new Cls_Entidad_Evento_Movil_Operario();
 
                                Entidad.id_Local = Convert.ToInt32(row["id_Local"].ToString());
                                Entidad.nombre_local = Convert.ToString(row["nombre_local"].ToString());

                                obj_List_Locales.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Locales;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Dato_Operarios(int id_local)
        {
            try
            {

                List<Cls_Entidad_Evento_Movil_Operario> obj_List_Locales = new List<Cls_Entidad_Evento_Movil_Operario>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_EVENTO_MOV_OPERARIO_OPERARIOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Local", SqlDbType.Int).Value = id_local;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Evento_Movil_Operario Entidad = new Cls_Entidad_Evento_Movil_Operario();

                                Entidad.id_Operario = Convert.ToInt32(row["id_Operario"].ToString());
                                Entidad.Nombre_Operario = Convert.ToString(row["Nombre_Operario"].ToString());

                                obj_List_Locales.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Locales;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

       // 



        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Dato_Listando_Datos_Cabecera(string id_local, string id_operario, string fecha_ini, string fecha_fin, string lista)
        {
            try
            {

                List<Cls_Entidad_Evento_Movil_Operario> obj_List_Operario = new List<Cls_Entidad_Evento_Movil_Operario>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_EVENTO_MOV_OPERARIO_RESUMEN", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.NVarChar).Value = id_operario;
                        cmd.Parameters.Add("@fecha_ini", SqlDbType.NVarChar).Value = fecha_ini;
                        cmd.Parameters.Add("@fecha_fin", SqlDbType.NVarChar).Value = fecha_fin;
                        cmd.Parameters.Add("@Lista", SqlDbType.NVarChar).Value = lista;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Evento_Movil_Operario Entidad = new Cls_Entidad_Evento_Movil_Operario();

    
                                Entidad.FECHA = Convert.ToString(row["FECHA"].ToString());
                                Entidad.HORA = Convert.ToString(row["HORA"].ToString());
                                Entidad.id_Operario = Convert.ToInt32(row["id_Operario"].ToString());
                                Entidad.Nombre_Operario = Convert.ToString(row["OPERARIO"].ToString());
                                Entidad.GPS = Convert.ToString(row["GPS"].ToString());
                                Entidad.BATERIA = Convert.ToString(row["BATERIA"].ToString());
                                Entidad.DATOS = Convert.ToString(row["DATOS"].ToString());
                                Entidad.MODO_AVION = Convert.ToString(row["MODO_AVION"].ToString());
                                obj_List_Operario.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Operario;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Dato_Datos_Operario_Detallado(int id_operario, string fecha)
        {
            try
            {

                List<Cls_Entidad_Evento_Movil_Operario> obj_List_Operario_Detallado = new List<Cls_Entidad_Evento_Movil_Operario>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_EVENTO_MOV_OPERARIO_DETALLADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@fecha", SqlDbType.NVarChar).Value = fecha; 

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Evento_Movil_Operario Entidad = new Cls_Entidad_Evento_Movil_Operario();

                                Entidad.FECHA = Convert.ToString(row["FECHA"].ToString());
                                Entidad.id_Operario = Convert.ToInt32(row["id_Operario"].ToString());
                                Entidad.Nombre_Operario = Convert.ToString(row["OPERARIO"].ToString());
                                Entidad.GPS = Convert.ToString(row["GPS"].ToString());
                                Entidad.BATERIA = Convert.ToString(row["BATERIA"].ToString());
                                Entidad.DATOS = Convert.ToString(row["DATOS"].ToString());
                                Entidad.MODO_AVION = Convert.ToString(row["MODO_AVION"].ToString());
                                Entidad.latitud_lectura = Convert.ToString(row["latitud_lectura"].ToString());
                                Entidad.longitud_lectura = Convert.ToString(row["longitud_lectura"].ToString());
                                obj_List_Operario_Detallado.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Operario_Detallado;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        
        //   Eventos movil Operario seguimiento

        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Dato_Listando_Datos_Cabecera_Seguimiento_Operario(string id_local, int id_operario, string fecha_ini, string fecha_fin, string lista)
        {
            try
            {

                List<Cls_Entidad_Evento_Movil_Operario> obj_List_Operario = new List<Cls_Entidad_Evento_Movil_Operario>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_EVENTO_MOV_OPERARIO_RECORRIDO_RESUMEN", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@fecha_ini", SqlDbType.NVarChar).Value = fecha_ini;
                        cmd.Parameters.Add("@fecha_fin", SqlDbType.NVarChar).Value = fecha_fin;
                        cmd.Parameters.Add("@Lista", SqlDbType.NVarChar).Value = lista;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Evento_Movil_Operario Entidad = new Cls_Entidad_Evento_Movil_Operario();                              
                                Entidad.FECHA = Convert.ToString(row["FECHA"].ToString());
                                Entidad.id_Operario = Convert.ToInt32(row["id_Operario"].ToString());
                                Entidad.Nombre_Operario = Convert.ToString(row["OPERARIO"].ToString());
                                Entidad.fecha_inicio_recorrido = Convert.ToString(row["fecha_inicio"].ToString());
                                Entidad.fecha_fin_recorrido = Convert.ToString(row["fecha_fin"].ToString());
                                Entidad.cantidad_paradas = Convert.ToInt32(row["cantidad_paradas"].ToString());
 
                                obj_List_Operario.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Operario;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        //  fin  Eventos movil Operario seguimiento



    }
}
