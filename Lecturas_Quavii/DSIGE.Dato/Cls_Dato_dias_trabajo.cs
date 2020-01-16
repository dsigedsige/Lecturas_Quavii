using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using DSIGE.Modelo;



namespace DSIGE.Dato
{
    public class Cls_Dato_dias_trabajo
    {
        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
       
        //string cadenaCnx = "Data Source = .; Initial Catalog = Cobra_Toma_Datos; Integrated Security = True";

        //listado

        public List<Cls_Entidad_dias_trabajo> Capa_Dato_Get_Listado_dias_trabajo()
        {
            try
            {
                List<Cls_Entidad_dias_trabajo> obj_List_Dias_trabajo = new List<Cls_Entidad_dias_trabajo>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_DIAS_TRABAJO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
             
                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_dias_trabajo Entidad = new Cls_Entidad_dias_trabajo();

                                Entidad.id_DiaTrabajo = Convert.ToInt32(row["id_DiaTrabajo"].ToString());
                                Entidad.id_Local =Convert.ToInt32(row["id_Local"].ToString());
                                Entidad.nombre_local = row["nombre_local"].ToString();
                                Entidad.id_Operario = Convert.ToInt32(row["id_Operario"].ToString());
                                Entidad.desc_operario = row["desc_operario"].ToString();
                                Entidad.NombreDia = row["NombreDia"].ToString();
                                Entidad.HoraEntrada = row["HoraEntrada"].ToString();
                                Entidad.HoraSalida = row["HoraSalida"].ToString();
                                Entidad.estado = Convert.ToInt32(row["estado"].ToString());

                                obj_List_Dias_trabajo.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Dias_trabajo;
            }
            catch (Exception e)
            {

                throw e;
            }
        }




        public List<Cls_Entidad_dias_trabajo> Capa_Dato_Get_Listado_operario()
        {
            try
            {
                List<Cls_Entidad_dias_trabajo> obj_List_Dias_trabajo = new List<Cls_Entidad_dias_trabajo>();

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

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_dias_trabajo Entidad = new Cls_Entidad_dias_trabajo();

                                Entidad.id_Operario = Convert.ToInt32(row["id_Operario"].ToString());
                                Entidad.desc_operario = row["desc_operario"].ToString();
                                obj_List_Dias_trabajo.Add(Entidad);
                            }
                        }
                    }
                }

                return obj_List_Dias_trabajo;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        //guardando 


        public bool Capa_Dato_Guardar_Informacion(Cls_Entidad_dias_trabajo Entidad)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_DIA_TRABAJO", cn))
                    {

                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = Entidad.id_Operario;
                        cmd.Parameters.Add("@NombreDia", SqlDbType.NVarChar).Value = Entidad.NombreDia;
                        cmd.Parameters.Add("@HoraEntrada", SqlDbType.NVarChar).Value = Entidad.HoraEntrada;
                        cmd.Parameters.Add("@HoraSalida", SqlDbType.NVarChar).Value = Entidad.HoraSalida;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = Entidad.estado;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = Entidad.usuario_creacion;
                        cmd.Parameters.Add("@id_Local", SqlDbType.Int).Value = Entidad.id_Local;

                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        ///Modificando

        public bool Capa_Dato_Modificar_Informacion(Cls_Entidad_dias_trabajo Entidad)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_DIA_TRABAJO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_DiaTrabajo", SqlDbType.Int).Value = Entidad.id_DiaTrabajo;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = Entidad.id_Operario;
                        cmd.Parameters.Add("@NombreDia", SqlDbType.NVarChar).Value = Entidad.NombreDia;
                        cmd.Parameters.Add("@HoraEntrada", SqlDbType.NVarChar).Value = Entidad.HoraEntrada;
                        cmd.Parameters.Add("@HoraSalida", SqlDbType.NVarChar).Value = Entidad.HoraSalida;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = Entidad.estado;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = Entidad.usuario_edicion;
                        cmd.Parameters.Add("@id_Local", SqlDbType.Int).Value = Entidad.id_Local;
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        ///ANULANDO

        public bool Capa_Dato_Anular_Informacion(Cls_Entidad_dias_trabajo Entidad)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_A_DIA_TRABAJO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        
                        cmd.Parameters.Add("@id_DiaTrabajo", SqlDbType.Int).Value = Entidad.id_DiaTrabajo;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = Entidad.estado;
                        cmd.Parameters.Add("@usuario_edicion", SqlDbType.Int).Value = Entidad.usuario_edicion;
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        ///auditoria

        public List<Cls_Entidad_dias_trabajo> Capa_Dato_Get_Listado_Auditoria(int id_dia)
        {
            try
            {

                List<Cls_Entidad_dias_trabajo> obj_List_Dias_trabajo = new List<Cls_Entidad_dias_trabajo>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_DIA_TRABAJO_AUDITORIA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_DiaTrabajo", SqlDbType.Int).Value = id_dia;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_dias_trabajo Entidad = new Cls_Entidad_dias_trabajo();
                                Entidad.usuario_creacion = Convert.ToInt32(row["usuario_creacion"].ToString());
                                Entidad.nombre_usu_crea = row["nombre_usu_crea"].ToString();
                                Entidad.fecha_creacion = row["fecha_creacion"].ToString();
                                Entidad.usuario_edicion = Convert.ToInt16(row["usuario_edicion"].ToString());
                                Entidad.nombre_usu_modi = row["nombre_usu_modi"].ToString();
                                Entidad.fecha_edicion = row["fecha_edicion"].ToString();
                                obj_List_Dias_trabajo.Add(Entidad);
                            }
                        }


                    }
                }

                return obj_List_Dias_trabajo;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

 


    }
}
