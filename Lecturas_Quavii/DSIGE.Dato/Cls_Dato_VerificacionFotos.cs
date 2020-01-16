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
    public class Cls_Dato_VerificacionFotos
    {
        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

        public List<VerificacionFoto_E> Capa_Dato_Get_ListaServicio()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO_VERIFICACION_FOTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();

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


        public List<HistoricoLectura> Capa_Dato_Get_ListaHistoricoLectura(string suministro)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<HistoricoLectura> ListHistorico = new List<HistoricoLectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Lectura_Historico_new", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Suministro", SqlDbType.VarChar).Value = suministro;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                HistoricoLectura obj_entidad = new HistoricoLectura();

                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.TipoServicio = Fila["TipoServicio"].ToString();
                                obj_entidad.Lectura = Fila["Lectura"].ToString();
                                obj_entidad.fechaAsignacion_lectura = Fila["fechaAsignacion_lectura"].ToString();
                                obj_entidad.consumo = Fila["consumo"].ToString();
                                obj_entidad.Observacion = Fila["Observacion"].ToString();
                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.foto = Fila["foto"].ToString();
                                ListHistorico.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListHistorico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<VerificacionFoto_E> Capa_Dato_Get_ListaOperarios()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
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
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();

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

        public List<VerificacionFoto_E> Capa_Dato_Get_ListaSector()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
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
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();

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

        public List<VerificacionFoto_E> Capa_Dato_Get_ListaObservacion()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACION_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();


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

        public List<VerificacionFoto_E> Capa_Dato_Get_ListaFotoLecturas(int local, string fecha, int servicio, int operario, int observacion, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_VERIFICACION_FOTO_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_local", SqlDbType.VarChar).Value = local;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.VarChar).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.VarChar).Value = operario; 
                        cmd.Parameters.Add("@id_observacion", SqlDbType.VarChar).Value = observacion;

                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();

                                obj_entidad.checkeado = false;
                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.id_Operario_Lectura = Fila["id_Operario_Lectura"].ToString();
                                obj_entidad.token = Fila["token"].ToString();
                                
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.fotourl = ruta + Fila["fotourl"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.lectura_movil = Fila["lectura_movil"].ToString();
                                //obj_entidad.usuario_modificacion = Fila["usuario_modificacion"].ToString();
                                //obj_entidad.fecha_modificacion = Fila["fecha_modificacion"].ToString();

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


        public List<VerificacionFoto_E> Capa_Dato_Get_ListaFotoLecturas_SinObs(string fecha, int servicio, int operario, int observacion, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_VERIFICACION_FOTO_LECTURA_SIN_OBS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.VarChar).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.VarChar).Value = operario;
                        cmd.Parameters.Add("@id_observacion", SqlDbType.VarChar).Value = observacion;

                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();

                                obj_entidad.checkeado = false;
                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.id_Operario_Lectura = Fila["id_Operario_Lectura"].ToString();
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.fotourl = ruta + Fila["fotourl"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.lectura_movil = Fila["lectura_movil"].ToString();

                                //obj_entidad.usuario_modificacion = Fila["usuario_modificacion"].ToString();
                                //obj_entidad.fecha_modificacion = Fila["fecha_modificacion"].ToString();

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
                    using (SqlCommand cmd = new SqlCommand("SP_I_U_VERIFICAR_FOTOS", cn))
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

        public string Capa_Dato_EnviarFotosSinObservacion(string lecturas)
        {
            string Resultado = "";
            try
            {
                //obteniendo Datos de Facturacion
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_U_VERIFICAR_FOTOS_FOTOS_SIN_OBSERVACION", cn))
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


        public string Capa_Dato_CambiarLectura(int id_lectura, string confirmacion_lectura)
        {
            string Resultado = "";
            try
            {
                //obteniendo Datos de Facturacion
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_ACTUALIZAR_CONFIR_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = id_lectura;
                        cmd.Parameters.Add("@confirmacion_lectura", SqlDbType.VarChar).Value = confirmacion_lectura;
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

        public string Capa_Dato_Set_VerificacionFotos_Inconsistencias(string lecturas)
        {
            string Resultado = "";
            try
            {
                //obteniendo Datos de Facturacion
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_U_VERIFICAR_FOTOS_INCONSISTENCIAS", cn))
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

        public List<VerificacionFoto_E> Capa_Dato_Get_ListaFotoLecturasValidacion(string fecha, int servicio, int operario, string sector, int observacion)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<VerificacionFoto_E> ListServicio = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_VALIDACION_FOTO_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.VarChar).Value = servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.VarChar).Value = operario;
                        cmd.Parameters.Add("@id_sector", SqlDbType.VarChar).Value = sector;


                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();
                                var rutaFisica = System.Web.Hosting.HostingEnvironment.MapPath("~/Foto/" + Convert.ToString(Fila["fotourl"]));
                                if (File.Exists(rutaFisica))
                                {
                                    obj_entidad.existeFoto = true;
                                }
                                else
                                {
                                    obj_entidad.existeFoto = false;
                                }
                                obj_entidad.checkeado = false;
                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.id_Operario_Lectura = Fila["id_Operario_Lectura"].ToString();
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.fotourl = ruta + Fila["fotourl"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.lectura_movil = Fila["lectura_movil"].ToString();


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


        public string Capa_Dato_Get_Update_LecturaActual(int id_lectura, int id_usuario, string lectura_actual)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_VERIFICACION_LECTURAS_LISTAR_UPDATE_LECTURA_ACTUAL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = id_lectura;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.Parameters.Add("@lectura_actual", SqlDbType.VarChar).Value = lectura_actual;
                        cmd.ExecuteNonQuery();
                        Resultado = "OK";
                    }
                }

            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }


        public List<VerificacionFoto_E> Capa_Dato_Get_MostrarFotos(int idLectura)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<VerificacionFoto_E> ListData = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURAS_FOTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Lectura", SqlDbType.Int).Value = idLectura;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();
                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.foto = ruta + Fila["foto"].ToString();
                                obj_entidad.url = ruta + Fila["foto"].ToString();
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


        public List<VerificacionFoto_E> Capa_Dato_Get_MostrarFotos_Reparto(int idReparto)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<VerificacionFoto_E> ListData = new List<VerificacionFoto_E>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_REPARTO_FOTO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_reparto", SqlDbType.Int).Value = idReparto;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                VerificacionFoto_E obj_entidad = new VerificacionFoto_E();
                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Reparto"].ToString());
                                obj_entidad.foto = ruta + Fila["foto"].ToString();
                                obj_entidad.url = ruta + Fila["foto"].ToString();
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

    }
}
