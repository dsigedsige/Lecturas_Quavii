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
using System.Threading;

namespace DSIGE.Dato
{
    public class Cls_Dato_AsignaOrdenTrabajo
    {
        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
        public List<Cls_Entidad_AsignaOrdenTrabajo.Locales> Capa_Dato_Get_ListaLocales()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Locales> Listlocal = new List<Cls_Entidad_AsignaOrdenTrabajo.Locales>();
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
                                Cls_Entidad_AsignaOrdenTrabajo.Locales obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Locales();

                                obj_entidad.loc_id = Convert.ToInt32(Fila["id_Local"].ToString());
                                obj_entidad.loc_nombre = Fila["nombre_local"].ToString();
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

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaServicio()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVICIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

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

        /// <summary>
        /// ListaServicios por Usuario segun permiso
        /// </summary>
        /// <returns></returns>
        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaServicioXusuario(int idusuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVIUSUARIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@IDUSUARIO", SqlDbType.Int).Value = idusuario;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {


                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.cantidad = Convert.ToInt32(Fila["cantidad"].ToString());

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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaServicioXusuario_II(int idusuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SERVIUSUARIO_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@IDUSUARIO", SqlDbType.Int).Value = idusuario;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {


                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.cantidad = Convert.ToInt32(Fila["cantidad"].ToString());

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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaServicio_new()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ENVIAR_TRABAJO_CLIENTE_SERVICIOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Estados> Capa_Dato_Get_ListaEstados()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Estados> ListEstados = new List<Cls_Entidad_AsignaOrdenTrabajo.Estados>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ESTADOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Estados obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Estados();

                                obj_entidad.id_Estado = Convert.ToInt32(Fila["id_Estado"].ToString());
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                ListEstados.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListEstados;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// lISTADO DE ESTADO RECEPCIONADO Y ENVIADO CLIENTE
        /// </summary>
        /// <returns></returns>
        /// 


        public string Capa_Dato_Set_ActualizandoLecturasRelecturas(int cod_lectura, string DatoModificar, string campoModificar)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_LECTURAS_RELECTURAS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = cod_lectura;
                        cmd.Parameters.Add("@DatoModificar", SqlDbType.VarChar).Value = DatoModificar;
                        cmd.Parameters.Add("@CampoModificar", SqlDbType.VarChar).Value = campoModificar;
                        cmd.ExecuteNonQuery();
                        Resultado = "Ok";
                    }
                }
            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }

        public string Capa_Dato_Set_ActualizandoCortesReconexiones(int cod_Cortelectura, string DatoModificar, string campoModificar)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_CORTES_RECONEXIONES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_Cortelectura", SqlDbType.Int).Value = cod_Cortelectura;
                        cmd.Parameters.Add("@DatoModificar", SqlDbType.VarChar).Value = DatoModificar;
                        cmd.Parameters.Add("@CampoModificar", SqlDbType.VarChar).Value = campoModificar;
                        cmd.ExecuteNonQuery();
                        Resultado = "Ok";
                    }
                }
            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }


                       
        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_ObservacionesLecturas()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACIONES_LECT_RELECT", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                ListObservaciones.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_ObservacionesCortes()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACIONES_CORTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                ListObservaciones.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_ObservacionesCortesResultado()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACIONES_CORTES_RESULTADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                ListObservaciones.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_ObservacionesReconexion()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACIONES_RECONEXIONES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                ListObservaciones.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_ObservacionesReconexionResultado()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACIONES_RECONEXIONES_RESULTADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                ListObservaciones.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Estados> Capa_Dato_Get_ListaEstadosLecturaEnviado()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Estados> ListEstados = new List<Cls_Entidad_AsignaOrdenTrabajo.Estados>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ESTADOSMOVIL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Estados obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Estados();

                                obj_entidad.id_Estado = Convert.ToInt32(Fila["id_Estado"].ToString());
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                ListEstados.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListEstados;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Estados> Capa_Dato_Get_estadosAll()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Estados> ListEstados = new List<Cls_Entidad_AsignaOrdenTrabajo.Estados>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ESTADOS_ALL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Estados obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Estados();

                                obj_entidad.id_Estado = Convert.ToInt32(Fila["id_Estado"].ToString());
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                ListEstados.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListEstados;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public object Capa_Dato_Get_ListaPreEnvioExcelTexto(int id_local, int id_tipo_servicio, int estado, string fechaAsignacion)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Trabajos_Cliente> listTrabajos = new List<Cls_Entidad_AsignaOrdenTrabajo.Trabajos_Cliente>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Enviar_Trabajos_al_Cliente", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@Servicios", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fechaAsignacion;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Trabajos_Cliente obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Trabajos_Cliente();

                                obj_entidad.ITEM = Convert.ToInt32(Fila["ITEM"].ToString());
                                obj_entidad.INSTALACION = Fila["INSTALACION"].ToString();
                                obj_entidad.EQUIPO = Fila["EQUIPO"].ToString();
                                obj_entidad.MEDIDOR = Fila["MEDIDOR"].ToString();
                                obj_entidad.CTACTO = Fila["CTACTO"].ToString();
                                obj_entidad.FecPlanLectura = Fila["FecPlanLectura"].ToString();
                                obj_entidad.FecRealPlanLectura = Fila["FecRealPlanLectura"].ToString();
                                obj_entidad.Hora_Plantilla = Fila["Hora_Plantilla"].ToString();
                                obj_entidad.Lectura = Fila["Lectura"].ToString();
                                obj_entidad.Nota = Fila["Nota"].ToString();
                                obj_entidad.Cod_Comentario = Fila["Cod_Comentario"].ToString();
                                obj_entidad.Comentario = Fila["Comentario"].ToString();
                                obj_entidad.Codigo_Lector = Fila["Codigo_Lector"].ToString();
                                listTrabajos.Add(obj_entidad);
                            }
                        }
                    }
                }

                return listTrabajos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.Tecnico> Capa_Dato_Get_ListaTecnico(int id_local, int id_tipo_servicio, int id_empresa)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Tecnico> ListTecnico = new List<Cls_Entidad_AsignaOrdenTrabajo.Tecnico>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("dsige_tecnico_asignado", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@idEmpresa", SqlDbType.Int).Value = id_empresa;


                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Tecnico obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Tecnico();

                                obj_entidad.ope_id = Convert.ToInt32(Fila["ope_id"].ToString());
                                obj_entidad.ope_nombre = Fila["ope_nombre"].ToString();
                                ListTecnico.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListTecnico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.OrdenesTrabajo> Capa_Dato_Get_ListandoOrdenesTrabajoGeneral(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion , string horainicial, string horafinal)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.OrdenesTrabajo> ListTecnico = new List<Cls_Entidad_AsignaOrdenTrabajo.OrdenesTrabajo>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ASIGNACION_ORDEN_TRABAJO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@tecnicoAsignado", SqlDbType.Int).Value = tecnico;
                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;                        
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@horaInicial", SqlDbType.VarChar).Value = horainicial;
                        cmd.Parameters.Add("@horaFinal", SqlDbType.VarChar).Value = horafinal;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.OrdenesTrabajo obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.OrdenesTrabajo();


                                if (id_tipo_servicio == 1 || id_tipo_servicio == 2)
                                {
                                    obj_entidad.checkeado = false;
                                    obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"]);
                                    obj_entidad.suministro_lectura = Convert.ToString(Fila["suministro_lectura"]);
                                    obj_entidad.medidor_lectura = Convert.ToString(Fila["medidor_lectura"]);
                                    obj_entidad.direccion_lectura = Convert.ToString(Fila["direccion_lectura"]);
                                    obj_entidad.lectura_Anterior = Convert.ToString(Fila["lectura_Anterior"]);
                                    obj_entidad.limite_Minimo_lectura = Convert.ToString(Fila["limite_Minimo_lectura"]);
                                    obj_entidad.limite_Maximo_lectura = Convert.ToString(Fila["limite_Maximo_lectura"]);

                                    obj_entidad.id_Operario_Lectura = Convert.ToString(Fila["id_Operario_Lectura"]);
                                    obj_entidad.ope_nombre = Convert.ToString(Fila["ope_nombre"]);
                                    //obj_entidad.fechaAsignacion_Lectura = string.IsNullOrEmpty(Fila["fechaAsignacion_Lectura"].ToString()) ? "" : Convert.ToDateTime(Fila["fechaAsignacion_Lectura"]).ToString("dd/MM/yyyy");
                                    obj_entidad.fechaAsignacion_Lectura = Fila["fechaAsignacion_Lectura"].ToString();
                                    obj_entidad.nombreCliente_lectura = Convert.ToString(Fila["nombreCliente_lectura"]);
                                    obj_entidad.estado = Convert.ToInt32(Fila["estado"]);
                                    obj_entidad.abreviatura_estado = Convert.ToString(Fila["abreviatura_estado"]);
                                    obj_entidad.nroInstalacion_lectura = Convert.ToString(Fila["nroInstalacion_lectura"]);

                                }
                                else if (id_tipo_servicio == 3 || id_tipo_servicio == 4)
                                {

                                    obj_entidad.checkeado = false;
                                    obj_entidad.ord_asig = Convert.ToInt32(Fila["ord_asig"]);
                                    //obj_entidad.fecha_importacion = string.IsNullOrEmpty(Fila["fecha_importacion"].ToString()) ? "" : Convert.ToDateTime(Fila["fecha_importacion"]).ToString("dd/MM/yyyy");
                                    obj_entidad.fecha_importacion = Fila["fecha_importacion"].ToString();
                                    obj_entidad.solicitud = Convert.ToString(Fila["solicitud"]);
                                    obj_entidad.cuenta_contrato = Convert.ToString(Fila["cuenta_contrato"]);
                                    obj_entidad.nro_serie_medidor = Convert.ToString(Fila["nro_serie_medidor"]);

                                    obj_entidad.cliente = Convert.ToString(Fila["cliente"]);
                                    obj_entidad.direc_instalacion = Convert.ToString(Fila["direc_instalacion"]);
                                    obj_entidad.distrito = Convert.ToString(Fila["distrito"]);
                                    obj_entidad.UL = Convert.ToString(Fila["UL"]);
                                    obj_entidad.apellido_lector = Convert.ToString(Fila["apellido_lector"]);

                                    obj_entidad.grand_total = Convert.ToString(Fila["grand_total"]);
                                    obj_entidad.count = Convert.ToString(Fila["count"]);
                                    obj_entidad.accion = Convert.ToString(Fila["accion"]);
                                    obj_entidad.resultado_corte = Convert.ToString(Fila["resultado_corte"]);
                                    //obj_entidad.fecha_corte = string.IsNullOrEmpty(Fila["fecha_corte"].ToString()) ? "" : Convert.ToDateTime(Fila["fecha_corte"]).ToString("dd/MM/yyyy");
                                    obj_entidad.fecha_corte = Fila["fecha_corte"].ToString();
                                    obj_entidad.mes_corte = Convert.ToString(Fila["mes_corte"]);
                                    obj_entidad.hora_corte = Convert.ToString(Fila["hora_corte"]);

                                    obj_entidad.lectura_corte = Convert.ToString(Fila["lectura_corte"]);
                                    obj_entidad.tipo_corte = Convert.ToString(Fila["tipo_corte"]);
                                    obj_entidad.cod_OBS = Convert.ToString(Fila["cod_OBS"]);
                                    obj_entidad.observacion_lectura = Convert.ToString(Fila["observacion_lectura"]);
                                    obj_entidad.detalle_observaciones = Convert.ToString(Fila["detalle_observaciones"]);
                                    obj_entidad.abreviatura_estado = Convert.ToString(Fila["estado"]);
                                    obj_entidad.id_Operario_corte = Convert.ToInt32(Fila["id_Operario_corte"]);
                                    obj_entidad.id_Corte = Convert.ToInt32(Fila["id_Corte"]);
                                    obj_entidad.suministro_corte = Convert.ToString(Fila["suministro_corte"]);

                                }

                                ListTecnico.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListTecnico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Dato_Get_ListaLecturaEnviarCliente(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, 
                                                                                            string fechaAsignacion, string tipoCliente, int id_supervisor, int id_operario_supervisor, string horainicial, string horafinal)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListTecnico = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("NEW_DSIGE_S_LECTURA_ENVIAR_CLIENTE_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@tecnicoAsignado", SqlDbType.Int).Value = tecnico;
                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@tipoCliente", SqlDbType.Int).Value = Convert.ToInt32(tipoCliente);
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        cmd.Parameters.Add("@horainicial", SqlDbType.VarChar).Value = horainicial;
                        cmd.Parameters.Add("@horafinal", SqlDbType.VarChar).Value = horafinal;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();
                                if (id_tipo_servicio == 1 || id_tipo_servicio == 2 || id_tipo_servicio == 9)
                                {
                                    obj_entidad.checkeado = false;

                                    obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"]);
                                    obj_entidad.foto = Convert.ToString(Fila["foto"]);
                                    obj_entidad.fechLectura = Convert.ToString(Fila["fechLectura"]);
                                    obj_entidad.suministro_lectura = Convert.ToString(Fila["suministro_lectura"]);
                                    obj_entidad.medidor_lectura = Convert.ToString(Fila["medidor_lectura"]);

                                    obj_entidad.Consumo1 = Convert.ToString(Fila["Consumo1"]);
                                    obj_entidad.Consumo2 = Convert.ToString(Fila["Consumo2"]);
                                    obj_entidad.Consumo3 = Convert.ToString(Fila["Consumo3"]);
                                    obj_entidad.ConsumoProm = Convert.ToString(Fila["ConsumoProm"]);
                                    obj_entidad.lectura_Anterior = Convert.ToString(Fila["lectura_Anterior"]);

                                    obj_entidad.lecMovil = Convert.ToString(Fila["lecMovil"]);
                                    obj_entidad.ConsumoActual = Convert.ToString(Fila["ConsumoActual"]);
                                    obj_entidad.Porcentaje = Convert.ToString(Fila["Porcentaje"]);
                                    obj_entidad.PorcentajeMayo = Convert.ToString(Fila["PorcentajeMayo"]);

                                    obj_entidad.LecturaMenor = Convert.ToString(Fila["LecturaMenor"]);
                                    obj_entidad.lecManual = Convert.ToString(Fila["lecManual"]);

                                    obj_entidad.idLecObs = Convert.ToString(Fila["idLecObs"]);
                                    obj_entidad.ope_nombre = Convert.ToString(Fila["ope_nombre"]);
                                    obj_entidad.Zona_lectura = Convert.ToString(Fila["Zona_lectura"]);
                                    obj_entidad.lecObs = Convert.ToString(Fila["lecObs"]);
                                    obj_entidad.direccion_lectura = Convert.ToString(Fila["direccion_lectura"]);

                                    obj_entidad.fechaAsignacion_Lectura = Convert.ToString(Fila["fechaAsignacion_Lectura"]);
                                    obj_entidad.nombreCliente_lectura = Convert.ToString(Fila["nombreCliente_lectura"]);
                                    obj_entidad.abreviatura_estado = Convert.ToString(Fila["abreviatura_estado"]);

                                    obj_entidad.fotourl = ruta + Convert.ToString(Fila["fotourl"]);
                                    obj_entidad.descripcion_observacion = Convert.ToString(Fila["descripcion_observacion"]);
                                    obj_entidad.desplazamiento = Convert.ToString(Fila["desplazamiento"]);

                                }
                                else if (id_tipo_servicio == 3 || id_tipo_servicio == 4) //---Cortes y Reconexiones
                                {

                                    obj_entidad.checkeado = false;
                                    obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"]);
                                    obj_entidad.orden = Fila["orden"].ToString();
                                    obj_entidad.foto = Fila["foto"].ToString();
                                    obj_entidad.fecha_corte = Fila["fecha_corte"].ToString();
                                    obj_entidad.suministro_corte = Fila["suministro_corte"].ToString();
                                    obj_entidad.cuenta_contrato = Fila["cuenta_contrato"].ToString();
                                    obj_entidad.medidor_corte = Fila["medidor_corte"].ToString();
                                    obj_entidad.cliente = Fila["cliente"].ToString();
                                    obj_entidad.lectura = Fila["lectura"].ToString();
                                    obj_entidad.observacion = Fila["observacion"].ToString();
                                    obj_entidad.ope_nombre = Fila["ope_nombre"].ToString();
                                    obj_entidad.direccion = Fila["direccion"].ToString();
                                    obj_entidad.distrito = Fila["distrito"].ToString();
                                    obj_entidad.fotourl = ruta + Convert.ToString(Fila["fotourl"]);
                                    obj_entidad.id_observacionResultado_corte = Fila["id_observacionResultado_corte"].ToString();
                                    obj_entidad.id_Observacion_corte = Fila["id_Observacion_corte"].ToString();
                                    obj_entidad.id_resultadoObsCorte = Fila["id_resultadoObsCorte"].ToString();
                                    obj_entidad.ubicacion_Medidor = Fila["ubicacion_Medidor"].ToString();
                                }
                                else if (id_tipo_servicio == 6)  // Repartos
                                {
                                    obj_entidad.checkeado = false;
                                    obj_entidad.id_Reparto = Convert.ToInt32(Fila["id_Reparto"]);
                                    obj_entidad.NroRecibo_Reparto = Fila["NroRecibo_Reparto"].ToString();
                                    obj_entidad.CuentaContrato_Reparto = Fila["CuentaContrato_Reparto"].ToString();
                                    obj_entidad.nombreCliente_Reparto = Fila["nombreCliente_Reparto"].ToString();
                                    obj_entidad.direccion_Reparto = Fila["direccion_Reparto"].ToString();
                                    obj_entidad.foto = Fila["foto"].ToString();
                                    obj_entidad.fecha_reparto = Fila["fecha_reparto"].ToString();
                                    obj_entidad.observacion = Fila["observacion"].ToString();
                                    obj_entidad.ope_nombre = Fila["ope_nombre"].ToString();
                                    obj_entidad.aviso_Reparto = Fila["aviso_Reparto"].ToString();
                                    obj_entidad.cargo_Reparto = Fila["cargo_Reparto"].ToString();
                                    obj_entidad.muestra_Reparto = Fila["muestra_Reparto"].ToString();
                                }

                                ListTecnico.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListTecnico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Dato_Get_ListaLectura_EnviarCliente(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor,int tecnico, string fechaAsignacion, string tipoCliente, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListData = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_ENVIAR_TRABAJO_CLIENTE_GENERAL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@tecnicoAsignado", SqlDbType.Int).Value = tecnico;
                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@tipoCliente", SqlDbType.Int).Value = Convert.ToInt32(tipoCliente);

                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();
                                obj_entidad.CL = Fila["CL"].ToString();
                                obj_entidad.OPERARIO = Fila["OPERARIO"].ToString();
                                obj_entidad.CANT_ASIGNADO = Convert.ToInt32(Fila["CANT_ASIGNADO"].ToString());
                                obj_entidad.CANT_REALIZADO = Convert.ToInt32(Fila["CANT_REALIZADO"].ToString());
                                obj_entidad.CANT_PENDIENTE = Convert.ToInt32(Fila["CANT_PENDIENTE"].ToString());
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


        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Dato_DescargaExcel_PreEnvio(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListTecnico = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_LECTURA_ENVIAR_CLIENTE_DESCARGAR_PRE_ENVIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@tecnicoAsignado", SqlDbType.Int).Value = tecnico;
                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;

                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();

                                //------LECTURA  RELECTURA Y RECLAMOS ----
                                if (id_tipo_servicio == 1 || id_tipo_servicio == 2 || id_tipo_servicio == 9)
                                {
                                    obj_entidad.nroInstalacion_lectura = Convert.ToString(Fila["nroInstalacion_lectura"]);
                                    obj_entidad.nroEquipo_lectura = Convert.ToString(Fila["nroEquipo_lectura"]);
                                    obj_entidad.medidor_lectura = Convert.ToString(Fila["medidor_lectura"]);
                                    obj_entidad.CuentaContrato = Convert.ToString(Fila["CuentaContrato"]);

                                    obj_entidad.FechaPlanLectura = Fila["FechaPlanLectura"].ToString();
                                    obj_entidad.FechaRealLectura = Fila["FechaRealLectura"].ToString();

                                    obj_entidad.HoraLecturaReal = Convert.ToString(Fila["HoraLecturaReal"]);
                                    obj_entidad.Lectura = Convert.ToString(Fila["Lectura"]);

                                    obj_entidad.NotaLecturista = Convert.ToString(Fila["NotaLecturista"]);
                                    obj_entidad.CodigoComentario = Convert.ToString(Fila["CodigoComentario"]);
                                    obj_entidad.Comentario = Convert.ToString(Fila["Comentario"]);
                                    obj_entidad.CodigoLector = Convert.ToString(Fila["CodigoLector"]);
                                    obj_entidad.desplazamiento = Convert.ToString(Fila["desplazamiento"]);
                                    obj_entidad.tieneFoto = Convert.ToString(Fila["tieneFoto"]);                                    

                                }
                                else if (id_tipo_servicio == 3)
                                {

                                    obj_entidad.claseAviso = Convert.ToString(Fila["Clase_aviso"]);
                                    obj_entidad.aviso = Convert.ToString(Fila["Aviso"]);
                                    obj_entidad.fechaAviso = Convert.ToString(Fila["Fecha_aviso"]);

                                    obj_entidad.docBloqueo = Convert.ToString(Fila["Documento_bloqueo"]);
                                    obj_entidad.ctaContrato = Convert.ToString(Fila["Cta_Contr"]);
                                    obj_entidad.nombreInterlocutor = Convert.ToString(Fila["Nombre_interlocutor"]);

                                    obj_entidad.claseCuenta = Convert.ToString(Fila["Clase_cuenta"]);
                                    obj_entidad.Deuda_Soles = Convert.ToString(Fila["Deuda_Soles"]);

                                    obj_entidad.cantRecibos = Convert.ToString(Fila["Cantidad_recibos"]);
                                    obj_entidad.instalacion = Convert.ToString(Fila["Instalacion"]);

                                    obj_entidad.nroSerieMedidor = Convert.ToString(Fila["N_Ser_Me"]);
                                    obj_entidad.direcInstalacion = Convert.ToString(Fila["Direccion_Instal"]);
                                    obj_entidad.distritoInstalacion = Convert.ToString(Fila["Distrito_Instal"]);

                                    obj_entidad.unidadLectura = Convert.ToString(Fila["Unidad_Lectura"]);
                                    obj_entidad.Ejecutante = Convert.ToString(Fila["Ejecutante"]);
                                    obj_entidad.orden = Convert.ToString(Fila["Orden"]);
                                    obj_entidad.creadoPor = Convert.ToString(Fila["Creado_por"]);

                                    obj_entidad.estatusSistema = Convert.ToString(Fila["Status_sistema"]);
                                    obj_entidad.codificacion = Convert.ToString(Fila["Codificacion"]);

                                    obj_entidad.codCodificacion = Convert.ToString(Fila["Cod_codificacion"]);
                                    obj_entidad.contador = Convert.ToString(Fila["Contador"]);
                                    obj_entidad.tcos = Convert.ToString(Fila["tcos"]);
                                    obj_entidad.hora = Convert.ToString(Fila["HORA"]);
                                    obj_entidad.lectura = Convert.ToString(Fila["LECTURA"]);

                                    obj_entidad.imposibilidad = Convert.ToString(Fila["IMPOSIBILIDAD"]);

                                }
                                else if (id_tipo_servicio == 4)
                                {
                                    obj_entidad.claseAviso = Convert.ToString(Fila["claseAviso"]);
                                    obj_entidad.aviso = Convert.ToString(Fila["aviso"]);
                                    obj_entidad.fechaAviso = Convert.ToString(Fila["fechaAviso"]);

                                    obj_entidad.docBloqueo = Convert.ToString(Fila["docBloqueo"]);
                                    obj_entidad.ctaContrato = Convert.ToString(Fila["ctaContrato"]);
                                    obj_entidad.nombreInterlocutor = Convert.ToString(Fila["nombreInterlocutor"]);

                                    obj_entidad.claseCuenta = Convert.ToString(Fila["claseCuenta"]);
                                    obj_entidad.aviso = Convert.ToString(Fila["aviso"]);
                                    obj_entidad.cantRecibos = Convert.ToString(Fila["cantRecibos"]);
                                    obj_entidad.instalacion = Convert.ToString(Fila["instalacion"]);

                                    obj_entidad.nroSerieMedidor = Convert.ToString(Fila["nroSerieMedidor"]);
                                    obj_entidad.direcInstalacion = Convert.ToString(Fila["direcInstalacion"]);
                                    obj_entidad.distritoInstalacion = Convert.ToString(Fila["distritoInstalacion"]);

                                    obj_entidad.unidadLectura = Convert.ToString(Fila["unidadLectura"]);
                                    obj_entidad.Ejecutante = Convert.ToString(Fila["Ejecutante"]);
                                    obj_entidad.orden = Convert.ToString(Fila["orden"]);

                                    obj_entidad.creadoPor = Convert.ToString(Fila["creadoPor"]);
                                    obj_entidad.contador = Convert.ToString(Fila["contador"]);
                                    obj_entidad.estadoInstalacion = Convert.ToString(Fila["estadoInstalacion"]);

                                    obj_entidad.estatusUsuario = Convert.ToString(Fila["estatusUsuario"]);
                                    obj_entidad.estatusSistema = Convert.ToString(Fila["estatusSistema"]);
                                    obj_entidad.codificacion = Convert.ToString(Fila["codificacion"]);

                                    obj_entidad.codCodificacion = Convert.ToString(Fila["codCodificacion"]);
                                    obj_entidad.tcos = Convert.ToString(Fila["tcos"]);
                                    obj_entidad.hora = Convert.ToString(Fila["hora"]);
                                    obj_entidad.lectura = Convert.ToString(Fila["lectura"]);

                                    obj_entidad.imposibilidad = Convert.ToString(Fila["imposibilidad"]);
                                    obj_entidad.nombreCliente = Convert.ToString(Fila["nombreCliente"]);

                                }

                                ListTecnico.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListTecnico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        private void CreateZipFile(List<string> items, string destination)
        {
            using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
            {
                foreach (string item in items)
                {
                    if (System.IO.File.Exists(item))
                    {
                        // Add the file in the root folder inside our zip file
                        zip.AddFile(item, "");
                    }
                    // if the item is a folder    
                    else if (System.IO.Directory.Exists(item))
                    {
                        // Add the folder in our zip file with the folder name as its name
                        zip.AddDirectory(item, new System.IO.DirectoryInfo(item).Name);
                    }
                }
                // Finally save the zip file to the destination we want
                zip.Save(destination);
            }
        }
        

        //public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Dato_DescargaFoto(string List_codigos, int servicio, int flag_multiple)
        //{
        //    try
        //    {
        //        cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
        //        string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];
        //        string[] fileEntries = null;
        //        string NombreFinal = "";
        //        string CadenaNombre = "";

        //        int ac = 0;
        //        int opcion = 0;

        //        if (servicio == 1 || servicio == 2)
        //        {
        //            opcion = 1;
        //        }
        //        else if (servicio == 3 || servicio == 4)
        //        {
        //            opcion = 2;
        //        }
        //        else if (servicio == 6)
        //        {
        //            opcion = 3;
        //        }

        //        List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListFotos = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
        //        using (SqlConnection cn = new SqlConnection(cadenaCnx))
        //        {
        //            cn.Open();
        //            using (SqlCommand cmd = new SqlCommand("SP_S_ENVIAR_TRABAJO_CLIENTE_FOTOS_TODOS_II", cn))
        //            {
        //                cmd.CommandTimeout = 0;
        //                cmd.CommandType = CommandType.StoredProcedure;

        //                cmd.Parameters.Add("@codLectura", SqlDbType.VarChar).Value = List_codigos;
        //                cmd.Parameters.Add("@opcion", SqlDbType.Int).Value = opcion;
        //                DataTable dt_detalle = new DataTable();

        //                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
        //                {
        //                    da.Fill(dt_detalle);
        //                    ac = 0;

        //                    if (flag_multiple > 1)
        //                    {
        //                        if (dt_detalle.Rows.Count > 0)
        //                        {
        //                            fileEntries = new string[dt_detalle.Rows.Count];
        //                        }
        //                    }

        //                    foreach (DataRow Fila in dt_detalle.Rows)
        //                    {
        //                        Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();

        //                        obj_entidad.nombreFoto = Convert.ToString(Fila["fotourl"]);
        //                        obj_entidad.fotourl = ruta + Convert.ToString(Fila["fotourl"]);
        //                        obj_entidad.cuenta_contrato = Convert.ToString(Fila["cuenta_contrato"]);
        //                        obj_entidad.fecha = Convert.ToString(Fila["fecha"]);
        //                        ListFotos.Add(obj_entidad);

        //                        if (flag_multiple > 1)
        //                        {
        //                            almacenando un matriz
        //                            fileEntries[ac] = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
        //                            ac = ac + 1;
        //                        }
        //                    }



        //                    if (flag_multiple > 1)
        //                    {
        //                        if (dt_detalle.Rows.Count > 0)
        //                        {

        //                            var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/LecturasEnviarCliente.zip");

        //                            if (System.IO.File.Exists(pathExist))
        //                            {
        //                                System.IO.File.Delete(pathExist);
        //                            }

        //                            Guardando el archivo en Zip                
        //                            using (ZipFile zip = new ZipFile())
        //                            {
        //                                zip.AddFiles(fileEntries, "");
        //                                zip.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/LecturasEnviarCliente.zip"));
        //                            }

        //                            using (ZipFile zip = ZipFile.Read(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/LecturasEnviarCliente.zip")))
        //                            {

        //                                for (int i = 0; i < ListFotos.Count; i++)
        //                                {
        //                                    NombreFinal = ListFotos[i].cuenta_contrato + "_" + ListFotos[i].fecha + "_" + ListFotos[i].nombreFoto + "_" + i + ".jpg";
        //                                    zip[ListFotos[i].nombreFoto].FileName = NombreFinal;
        //                                }
        //                                zip.Save();
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        return ListFotos;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

                                                         
        public string Capa_Dato_DescargaFoto(string List_codigos, int servicio, int flag_multiple, int id_usuario)
        {
            string Res = "";
            string nombreFile = id_usuario + "LecturasEnviar"+ servicio + "Cliente.zip";
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];
                string[] fileEntries = null;
                string nombreFoto = "";

                int ac = 0;
                int opcion = 0;

                if (servicio == 1 || servicio == 2 || servicio == 9)
                {
                    opcion = 1;
                }
                else if (servicio == 3 || servicio == 4)
                {
                    opcion = 2;
                }
                else if (servicio == 6)
                {
                    opcion = 3;
                }

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_ENVIAR_TRABAJO_CLIENTE_FOTOS_MASIVO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@codLectura", SqlDbType.VarChar).Value = List_codigos;
                        cmd.Parameters.Add("@opcion", SqlDbType.Int).Value = opcion;
                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            ac = 0;

                            if (flag_multiple > 1)
                            {
                                if (dt_detalle.Rows.Count > 0)
                                {
                                    fileEntries = new string[dt_detalle.Rows.Count];
                                }
                            }

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Res = "2|" + ruta + Convert.ToString(Fila["fotourl"]);
     
                                if (flag_multiple > 1)
                                {
                                    //almacenando un matriz
                                    var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                    if (System.IO.File.Exists(pathExist))
                                    {
                                        fileEntries[ac] = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                        ac = ac + 1;
                                    }
                                    else {
                                        //-----almacenando la url de la foto para borrarlo luego
                                        nombreFoto = Convert.ToString(Fila["fotourl"]);
                                        using (SqlCommand cmds = new SqlCommand("sp_i_fotos_error", cn))
                                        {
                                            cmds.CommandTimeout = 0;
                                            cmds.CommandType = CommandType.StoredProcedure;
                                            cmds.Parameters.Add("@ruta_foto", SqlDbType.VarChar).Value = nombreFoto;
                                            cmds.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }
                              

                            if (flag_multiple > 1)
                            {
                                List<string> listFotos = new List<string>(fileEntries);
                                for (int index = 0; index < listFotos.Count; index++)
                                {
                                    bool nullOrEmpty = string.IsNullOrEmpty(listFotos[index]);
                                    if (nullOrEmpty)
                                    {
                                        listFotos.RemoveAt(index);
                                        --index;
                                    }
                                }

                                if (dt_detalle.Rows.Count > 0)
                                {

                                    var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/" + nombreFile);

                                    if (System.IO.File.Exists(pathExist))
                                    {
                                        System.IO.File.Delete(pathExist);
                                    }

                                    //Guardando el archivo en Zip                
                                    using (ZipFile zip = new ZipFile())
                                    {
                                        zip.AddFiles(listFotos, "");
                                        zip.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/"+ nombreFile));
                                    }
                                }
                            }
                        }
                    }
                }


                
                if (flag_multiple > 1)
                {
                    Res = "1|" + ruta_descarga + nombreFile;
                }    
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }


        public string Capa_Dato_DescargaFoto_sinObservacion(string List_codigos, int servicio, int flag_multiple, int id_usuario)
        {

            string Res = "";
            string nombreFile = id_usuario + "LecturasEnviar" + servicio + "Cliente.zip";
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];
                string[] fileEntries = null;
                string nombreFoto = "";

                int ac = 0;
                int opcion = 0;

                if (servicio == 1 || servicio == 2 || servicio == 9)
                {
                    opcion = 1;
                }
                else if (servicio == 3 || servicio == 4)
                {
                    opcion = 2;
                }
                else if (servicio == 6)
                {
                    opcion = 3;
                }

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_ENVIAR_TRABAJO_CLIENTE_FOTOS_MASIVO_SINOBSERVACION", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@codLectura", SqlDbType.VarChar).Value = List_codigos;
                        cmd.Parameters.Add("@opcion", SqlDbType.Int).Value = opcion;
                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            ac = 0;

                            if (flag_multiple > 1)
                            {
                                if (dt_detalle.Rows.Count > 0)
                                {
                                    fileEntries = new string[dt_detalle.Rows.Count];
                                }
                            }

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Res = "2|" + ruta + Convert.ToString(Fila["fotourl"]);

                                if (flag_multiple > 1)
                                {
                                    //almacenando un matriz
                                    var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                    if (System.IO.File.Exists(pathExist))
                                    {
                                        fileEntries[ac] = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                        ac = ac + 1;
                                    }
                                    else
                                    {
                                        //-----almacenando la url de la foto para borrarlo luego
                                        nombreFoto = Convert.ToString(Fila["fotourl"]);
                                        using (SqlCommand cmds = new SqlCommand("sp_i_fotos_error", cn))
                                        {
                                            cmds.CommandTimeout = 0;
                                            cmds.CommandType = CommandType.StoredProcedure;
                                            cmds.Parameters.Add("@ruta_foto", SqlDbType.VarChar).Value = nombreFoto;
                                            cmds.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }


                            if (flag_multiple > 1)
                            {
                                List<string> listFotos = new List<string>(fileEntries);
                                for (int index = 0; index < listFotos.Count; index++)
                                {
                                    bool nullOrEmpty = string.IsNullOrEmpty(listFotos[index]);
                                    if (nullOrEmpty)
                                    {
                                        listFotos.RemoveAt(index);
                                        --index;
                                    }
                                }

                                if (dt_detalle.Rows.Count > 0)
                                {

                                    var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/" + nombreFile);

                                    if (System.IO.File.Exists(pathExist))
                                    {
                                        System.IO.File.Delete(pathExist);
                                    }

                                    //Guardando el archivo en Zip                
                                    using (ZipFile zip = new ZipFile())
                                    {
                                        zip.AddFiles(listFotos, "");
                                        zip.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/" + nombreFile));
                                    }
                                }
                            }
                        }
                    }
                }

                if (flag_multiple > 1)
                {
                    Res = "1|" + ruta_descarga + nombreFile;
                }
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }
               
        public string Capa_Dato_DescargaFoto_new(int local, int servicio, string fecha_asignacion, int id_usuario , int id_supervisor, int id_operario_supervisor)
        {
            string Res = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                string nombreFile = id_usuario + "LecturasEnviarCliente_resumen.zip";

                string[] fileEntries = null;
                string nombreFoto = "";
                int ac = 0;

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListFotos = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_ENVIAR_TRABAJO_CLIENTE_FOTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_asignacion;

                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            ac = 0;

             
                            if (dt_detalle.Rows.Count > 0)
                            {
                                fileEntries = new string[dt_detalle.Rows.Count];
                            }
              

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();
                                obj_entidad.fotourl = ruta + Convert.ToString(Fila["fotourl"]);
                                ListFotos.Add(obj_entidad);

                                //almacenando un matriz
                                var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                if (System.IO.File.Exists(pathExist))
                                {
                                    fileEntries[ac] = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                    ac = ac + 1;
                                }
                                else {
                                    //-----almacenando la url de la foto para borrarlo luego
                                    nombreFoto = Convert.ToString(Fila["fotourl"]);
                                    using (SqlCommand cmds = new SqlCommand("sp_i_fotos_error", cn))
                                    {
                                        cmds.CommandTimeout = 0;
                                        cmds.CommandType = CommandType.StoredProcedure;
                                        cmds.Parameters.Add("@ruta_foto", SqlDbType.VarChar).Value = nombreFoto;
                                        cmds.ExecuteNonQuery();
                                    }
                                }
                
                            }

                            List<string> listFotos = new List<string>(fileEntries);

                            for (int index = 0; index < listFotos.Count; index++)
                            {
                                bool nullOrEmpty = string.IsNullOrEmpty(listFotos[index]);
                                if (nullOrEmpty)
                                {
                                    listFotos.RemoveAt(index);
                                    --index;
                                }
                            }


                            if (dt_detalle.Rows.Count > 0)
                            {

                                var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/"+ nombreFile);

                                if (System.IO.File.Exists(pathExist))
                                {
                                    System.IO.File.Delete(pathExist);
                                }

                                //Guardando el archivo en Zip                
                                using (ZipFile zip = new ZipFile())
                                {
                                    zip.AddFiles(listFotos, "");
                                    zip.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/"+ nombreFile));
                                }
                            }
                        }
                    }
                }

                string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

                Res = "1|" + ruta_descarga + nombreFile;
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }
               
        public string Capa_Dato_DescargaFoto_new_sinObservacion(int local, int servicio, string fecha_asignacion, int id_usuario, int id_supervisor, int id_operario_supervisor)
        {
            string Res = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

                string nombreFile = id_usuario + "LecturasEnviarCliente_resumen.zip";

                string[] fileEntries = null;
                string nombreFoto = "";
                int ac = 0;

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListFotos = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_ENVIAR_TRABAJO_CLIENTE_FOTOS_SIN_OBSERVACION", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_asignacion;

                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            ac = 0;


                            if (dt_detalle.Rows.Count > 0)
                            {
                                fileEntries = new string[dt_detalle.Rows.Count];
                            }


                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();
                                obj_entidad.fotourl = ruta + Convert.ToString(Fila["fotourl"]);
                                ListFotos.Add(obj_entidad);

                                //almacenando un matriz
                                var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                if (System.IO.File.Exists(pathExist))
                                {
                                    fileEntries[ac] = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                    ac = ac + 1;
                                }
                                else
                                {
                                    //-----almacenando la url de la foto para borrarlo luego
                                    nombreFoto = Convert.ToString(Fila["fotourl"]);
                                    using (SqlCommand cmds = new SqlCommand("sp_i_fotos_error", cn))
                                    {
                                        cmds.CommandTimeout = 0;
                                        cmds.CommandType = CommandType.StoredProcedure;
                                        cmds.Parameters.Add("@ruta_foto", SqlDbType.VarChar).Value = nombreFoto;
                                        cmds.ExecuteNonQuery();
                                    }
                                }

                            }

                            List<string> listFotos = new List<string>(fileEntries);

                            for (int index = 0; index < listFotos.Count; index++)
                            {
                                bool nullOrEmpty = string.IsNullOrEmpty(listFotos[index]);
                                if (nullOrEmpty)
                                {
                                    listFotos.RemoveAt(index);
                                    --index;
                                }
                            }


                            if (dt_detalle.Rows.Count > 0)
                            {

                                var pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/" + nombreFile);

                                if (System.IO.File.Exists(pathExist))
                                {
                                    System.IO.File.Delete(pathExist);
                                }

                                //Guardando el archivo en Zip                
                                using (ZipFile zip = new ZipFile())
                                {
                                    zip.AddFiles(listFotos, "");
                                    zip.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/" + nombreFile));
                                }
                            }
                        }
                    }
                }
                Res = "1|" + ruta_descarga + nombreFile;
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }





        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> Capa_Dato_DescargaExcel_EnvioCliente(int empresa, int id_local, int id_tipo_servicio, int estado, string suministro, string medidor, int tecnico, string fechaAsignacion)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListTecnico = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_ENVIAR_CLIENTE_DESCARGAR_ENVIOCLIENTE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = id_tipo_servicio;
                        cmd.Parameters.Add("@idLocal", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@tecnicoAsignado", SqlDbType.Int).Value = tecnico;
                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();


                                if (id_tipo_servicio == 1 || id_tipo_servicio == 2)
                                {
                                    obj_entidad.nroInstalacion_lectura = Convert.ToString(Fila["nroInstalacion_lectura"]);
                                    obj_entidad.nroEquipo_lectura = Convert.ToString(Fila["nroEquipo_lectura"]);
                                    obj_entidad.medidor_lectura = Convert.ToString(Fila["medidor_lectura"]);
                                    obj_entidad.CuentaContrato = Convert.ToString(Fila["CuentaContrato"]);

                                    //obj_entidad.FechaPlanLectura = string.IsNullOrEmpty(Fila["FechaPlanLectura"].ToString()) ? "" : Convert.ToDateTime(Fila["FechaPlanLectura"]).ToString("dd/MM/yyyy");
                                    //obj_entidad.FechaRealLectura = string.IsNullOrEmpty(Fila["FechaRealLectura"].ToString()) ? "" : Convert.ToDateTime(Fila["FechaRealLectura"]).ToString("dd/MM/yyyy");
                                    obj_entidad.FechaPlanLectura = Fila["FechaPlanLectura"].ToString();
                                    obj_entidad.FechaRealLectura = Fila["FechaRealLectura"].ToString();

                                    obj_entidad.HoraLecturaReal = Convert.ToString(Fila["HoraLecturaReal"]);
                                    obj_entidad.Lectura = Convert.ToString(Fila["Lectura"]);

                                    obj_entidad.NotaLecturista = Convert.ToString(Fila["NotaLecturista"]);
                                    obj_entidad.CodigoComentario = Convert.ToString(Fila["CodigoComentario"]);
                                    obj_entidad.Comentario = Convert.ToString(Fila["Comentario"]);
                                    obj_entidad.CodigoLector = Convert.ToString(Fila["CodigoLector"]);
                                }
                                else if (id_tipo_servicio == 3)
                                {
                                    obj_entidad.unidadLectura = Convert.ToString(Fila["unidadLectura"]);
                                    obj_entidad.ejecutante = Convert.ToString(Fila["ejecutante"]);
                                    obj_entidad.orden = Convert.ToString(Fila["orden"]);
                                    obj_entidad.creado = Convert.ToString(Fila["creado"]);

                                    obj_entidad.estatusSistema = Convert.ToString(Fila["estatusSistema"]);
                                    obj_entidad.codCodificacion = Convert.ToString(Fila["codCodificacion"]);
                                    obj_entidad.codificacion = Convert.ToString(Fila["codificacion"]);
                                    obj_entidad.contador = Convert.ToString(Fila["contador"]);

                                    obj_entidad.tcos = Convert.ToString(Fila["tcos"]);
                                    obj_entidad.hora = Convert.ToString(Fila["hora"]);
                                    obj_entidad.lectura = Convert.ToString(Fila["lectura"]);
                                    obj_entidad.imposibilidad = Convert.ToString(Fila["imposibilidad"]);
                                    obj_entidad.observacion = Convert.ToString(Fila["observacion"]);
                                }
                                else if (id_tipo_servicio == 4)
                                {
                                    obj_entidad.claseAviso = Convert.ToString(Fila["claseAviso"]);
                                    obj_entidad.aviso = Convert.ToString(Fila["aviso"]);
                                    obj_entidad.fechaAviso = Convert.ToString(Fila["fechaAviso"]);

                                    obj_entidad.docBloqueo = Convert.ToString(Fila["docBloqueo"]);
                                    obj_entidad.ctaContrato = Convert.ToString(Fila["ctaContrato"]);
                                    obj_entidad.nombreInterlocutor = Convert.ToString(Fila["nombreInterlocutor"]);

                                    obj_entidad.claseCuenta = Convert.ToString(Fila["claseCuenta"]);
                                    obj_entidad.aviso = Convert.ToString(Fila["aviso"]);
                                    obj_entidad.cantRecibos = Convert.ToString(Fila["cantRecibos"]);
                                    obj_entidad.instalacion = Convert.ToString(Fila["instalacion"]);

                                    obj_entidad.nroSerieMedidor = Convert.ToString(Fila["nroSerieMedidor"]);
                                    obj_entidad.direcInstalacion = Convert.ToString(Fila["direcInstalacion"]);
                                    obj_entidad.distritoInstalacion = Convert.ToString(Fila["distritoInstalacion"]);

                                    obj_entidad.unidadLectura = Convert.ToString(Fila["unidadLectura"]);
                                    obj_entidad.Ejecutante = Convert.ToString(Fila["Ejecutante"]);
                                    obj_entidad.orden = Convert.ToString(Fila["orden"]);

                                    obj_entidad.creadoPor = Convert.ToString(Fila["creadoPor"]);
                                    obj_entidad.contador = Convert.ToString(Fila["contador"]);
                                    obj_entidad.estadoInstalacion = Convert.ToString(Fila["estadoInstalacion"]);

                                    obj_entidad.estatusUsuario = Convert.ToString(Fila["estatusUsuario"]);
                                    obj_entidad.estatusSistema = Convert.ToString(Fila["estatusSistema"]);
                                    obj_entidad.codificacion = Convert.ToString(Fila["codificacion"]);

                                    obj_entidad.codCodificacion = Convert.ToString(Fila["codCodificacion"]);
                                    obj_entidad.tcos = Convert.ToString(Fila["tcos"]);
                                    obj_entidad.hora = Convert.ToString(Fila["hora"]);
                                    obj_entidad.lectura = Convert.ToString(Fila["lectura"]);

                                    obj_entidad.imposibilidad = Convert.ToString(Fila["imposibilidad"]);
                                    obj_entidad.nombreCliente = Convert.ToString(Fila["nombreCliente"]);

                                }

                                ListTecnico.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListTecnico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura> Capa_Dato_LecturasRelectura_Historico(int empresa, string suministro, int tiposervicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura> ListHistorico = new List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ORDENTRABAJO_HISTORICO_LECTURARELECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@codSuministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = tiposervicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura();
                                obj_entidad.medidor_lectura = Convert.ToString(Fila["medidor_lectura"]);
                                //obj_entidad.fechaLecturaMovil_lectura =  string.IsNullOrEmpty(Fila["fechaLecturaMovil_lectura"].ToString()) ? "" : Convert.ToDateTime(Fila["fechaLecturaMovil_lectura"]).ToString("dd/MM/yyyy");
                                obj_entidad.fechaLecturaMovil_lectura = Fila["fechaLecturaMovil_lectura"].ToString();
                                obj_entidad.lectura_Anterior = Convert.ToString(Fila["lectura_Anterior"]);
                                obj_entidad.ope_nombre = Convert.ToString(Fila["ope_nombre"]);
                                obj_entidad.Observacion_lectura = Convert.ToString(Fila["Observacion_lectura"]);
                                obj_entidad.abreviatura_estado = Convert.ToString(Fila["abreviatura_estado"]);
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

        public List<Cls_Entidad_AsignaOrdenTrabajo.Historico_CortesReconexiones> Capa_Dato_CortesReconexiones_Historico(int empresa, string suministro, int tiposervicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Historico_CortesReconexiones> ListHistorico = new List<Cls_Entidad_AsignaOrdenTrabajo.Historico_CortesReconexiones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ORDENTRABAJO_HISTORICO_CORTERECONEXION", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                        cmd.Parameters.Add("@codSuministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = tiposervicio;


                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Historico_CortesReconexiones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Historico_CortesReconexiones();
                                obj_entidad.medidor_corte = Convert.ToString(Fila["medidor_corte"]);
                                //obj_entidad.fechaAsignacion_corte = string.IsNullOrEmpty(Fila["fechaAsignacion_corte"].ToString()) ? "" : Convert.ToDateTime(Fila["fechaAsignacion_corte"]).ToString("dd/MM/yyyy");
                                obj_entidad.fechaAsignacion_corte = Fila["fechaAsignacion_corte"].ToString();
                                obj_entidad.id_Corte = Convert.ToString(Fila["id_Corte"]);
                                obj_entidad.ope_nombre = Convert.ToString(Fila["ope_nombre"]);
                                obj_entidad.Observacion_corte = Convert.ToString(Fila["Observacion_corte"]);
                                obj_entidad.abreviatura_estado = Convert.ToString(Fila["abreviatura_estado"]);
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

        public List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> DListaFotos(int idLectura, int opcion)
        {
            try
            {
                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> lOb = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_ENVIAR_TRABAJO_CLIENTE_FOTOS", idLectura, opcion))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio()
                                    {
                                        id_Lectura = Convert.ToInt32(iDr["fot_id_lectura"]),
                                        foto = ruta + Convert.ToString(iDr["fot_nombre"]),
                                        nombreFoto = Convert.ToString(iDr["fot_nombre"])
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


        public string Capa_Dato_UpdateFechaAplicativoMovil(int empresa, string fechamovil, int id_usuario, int tipoServicio, List<int> List_codigos)
        {

            // Resultado 0 error, 1 OK
            var Resultado = "";

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                string lecturas = "";
                string id_cadenaCodigos = "";

                List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura> ListHistorico = new List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    if (tipoServicio == 1 || tipoServicio == 2)
                    {
                        foreach (var row in List_codigos)
                        {
                            lecturas = lecturas + row.ToString() + ',';
                        }
                        id_cadenaCodigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                        using (SqlCommand cmd = new SqlCommand("SP_U_ORDENTRABAJO_ENVIARMOVIL", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                            cmd.Parameters.Add("@id_lecturas", SqlDbType.VarChar).Value = id_cadenaCodigos;
                            cmd.Parameters.Add("@fechamovil", SqlDbType.VarChar).Value = fechamovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                            cmd.ExecuteNonQuery();
                            Resultado = "1|Ok";
                        }
                    }
                    else if (tipoServicio == 3 || tipoServicio == 4)
                    {
                        foreach (var row in List_codigos)
                        {
                            lecturas = lecturas + row.ToString() + ',';
                        }
                        id_cadenaCodigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                        using (SqlCommand cmd = new SqlCommand("SP_U_ORDENTRABAJO_ENVIARMOVIL_CORTESRECONEXION", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                            cmd.Parameters.Add("@id_cortes", SqlDbType.VarChar).Value = id_cadenaCodigos;
                            cmd.Parameters.Add("@fechamovil", SqlDbType.VarChar).Value = fechamovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                            cmd.ExecuteNonQuery();
                            Resultado = "1|Ok";
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Resultado = "0|" + ex.Message;
            }
            return Resultado;
        }



        public string Capa_Dato_UpdateFechaAplicativoMovilLectura(int empresa, string fechamovil, int id_usuario, int tipoServicio, List<int> List_codigos)
        {

            // Resultado 0 error, 1 OK
            var Resultado = "";

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                string lecturas = "";
                string id_cadenaCodigos = "";

                List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura> ListHistorico = new List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    if (tipoServicio == 1 || tipoServicio == 2)
                    {
                        foreach (var row in List_codigos)
                        {
                            lecturas = lecturas + row.ToString() + ',';
                        }
                        id_cadenaCodigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                        using (SqlCommand cmd = new SqlCommand("SP_U_LECT_ENV_LECTRELECT", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                            cmd.Parameters.Add("@id_lecturas", SqlDbType.VarChar).Value = id_cadenaCodigos;
                            cmd.Parameters.Add("@fechamovil", SqlDbType.VarChar).Value = fechamovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                            cmd.ExecuteNonQuery();
                            Resultado = "1|Ok";
                        }
                    }
                    else if (tipoServicio == 3 || tipoServicio == 4)
                    {
                        foreach (var row in List_codigos)
                        {
                            lecturas = lecturas + row.ToString() + ',';
                        }
                        id_cadenaCodigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                        using (SqlCommand cmd = new SqlCommand("SP_U_LECT_ENV_RECOXCORTE", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.Add("@empresa", SqlDbType.Int).Value = empresa;
                            cmd.Parameters.Add("@id_cortes", SqlDbType.VarChar).Value = id_cadenaCodigos;
                            cmd.Parameters.Add("@fechamovil", SqlDbType.VarChar).Value = fechamovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                            cmd.ExecuteNonQuery();
                            Resultado = "1|Ok";
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Resultado = "0|" + ex.Message;
            }
            return Resultado;
        }



        public string Capa_Dato_UpdateReasignaOperador(int empresa, int usuario, int TipoServicio, int tecnico, string fechaReasigna, List<int> List_codigos, int opcion)
        {
            // Resultado 0 error, 1 OK
            var Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                string lecturas = "";
                string id_codigos = "";

                List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura> ListHistorico = new List<Cls_Entidad_AsignaOrdenTrabajo.Historico_LecturaRelectura>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    if (TipoServicio == 1 || TipoServicio == 2)
                    {
                        foreach (var row in List_codigos)
                        {
                            lecturas = lecturas + row.ToString() + ',';
                        }
                        id_codigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                        using (SqlCommand cmd = new SqlCommand("dsige_asigna_lec_relec_update_resig_operador", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@codLectura", SqlDbType.VarChar).Value = id_codigos;
                            cmd.Parameters.Add("@idOperador", SqlDbType.Int).Value = tecnico;
                            cmd.Parameters.Add("@fechToma", SqlDbType.VarChar).Value = fechaReasigna;
                            cmd.Parameters.Add("@usuEdi", SqlDbType.Int).Value = usuario;
                            cmd.Parameters.Add("@idEmp", SqlDbType.Int).Value = empresa;
                            cmd.Parameters.Add("@opcion", SqlDbType.Int).Value = opcion;
                            cmd.ExecuteNonQuery();
                            Resultado = "1|Ok";
                        }
                    }
                    else if (TipoServicio == 3 || TipoServicio == 4)
                    {
                        foreach (var row in List_codigos)
                        {
                            lecturas = lecturas + row.ToString() + ',';
                        }
                        id_codigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                        using (SqlCommand cmd = new SqlCommand("dsige_asigna_corte_reconexion_update_resig_operador", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@codCorte", SqlDbType.VarChar).Value = id_codigos;
                            cmd.Parameters.Add("@idOperador", SqlDbType.Int).Value = tecnico;
                            cmd.Parameters.Add("@fechToma", SqlDbType.VarChar).Value = fechaReasigna;
                            cmd.Parameters.Add("@usuEdi", SqlDbType.Int).Value = usuario;
                            cmd.Parameters.Add("@idEmp", SqlDbType.Int).Value = empresa;
                            cmd.Parameters.Add("@opcion", SqlDbType.Int).Value = opcion;
                            cmd.ExecuteNonQuery();
                            Resultado = "1|Ok";
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Resultado = "0|" + ex.Message;
            }
            return Resultado;
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> Capa_Dato_Get_ListaVisualizacionesPendientes(int local,string fechaAsigna, int TipoServicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> ListPendientes = new List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_VISUALIZAR_PENDIENTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@TipoServicio", SqlDbType.Int).Value = TipoServicio;


                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente();

                                obj_entidad.id_operario = Convert.ToString(Fila["id_operario"]);
                                obj_entidad.Operario = Convert.ToString(Fila["Operario"]);
                                obj_entidad.id_estado = Convert.ToString(Fila["id_estado"]);
                                obj_entidad.Estado = Convert.ToString(Fila["Estado"]);
                                obj_entidad.Cantidad = Convert.ToInt32(Fila["Cantidad"]);
                                obj_entidad.fechaAsignacion = Convert.ToString(Fila["fechaAsignacion"]);

                                ListPendientes.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListPendientes;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> Capa_Dato_Get_ListaVisualizacionesPendientes_Detallado(int TipoServicio, string id_operario, string id_estado, string fechaAsigna)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> ListPendientes_Detallado = new List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_VISUALIZAR_PENDIENTES_DETALLADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@TipoServicio", SqlDbType.Int).Value = TipoServicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.VarChar).Value = id_operario;
                        cmd.Parameters.Add("@id_estado", SqlDbType.VarChar).Value = id_estado;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente();

                                obj_entidad.Orden = Fila["Orden"].ToString();
                                obj_entidad.Cuenta_Contrato = Convert.ToString(Fila["Cuenta_Contrato"]);
                                obj_entidad.Nro_Instalacion = Convert.ToString(Fila["Nro_Instalacion"]);
                                obj_entidad.Nro_Serie = Convert.ToString(Fila["Nro_Serie"]);
                                obj_entidad.Direccion = Convert.ToString(Fila["Direccion"]);
                                obj_entidad.Ubicacion_Med = Convert.ToString(Fila["Ubicacion_Med"]);
                                obj_entidad.Fecha_Alta = Convert.ToString(Fila["Fecha_Alta"]);
                                obj_entidad.Cliente = Convert.ToString(Fila["Cliente"]);
                                obj_entidad.Inter_Locutor = Convert.ToString(Fila["Inter_Locutor"]);
                                obj_entidad.Operario = Convert.ToString(Fila["Operario"]);
                                ListPendientes_Detallado.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListPendientes_Detallado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_AsignaOrdenTrabajo.validacion> Capa_Dato_Get_ListaLecturaEnviarCliente_Validacion(int estado, string fechaAsignacion, string estado_new, int resultado, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.validacion> ListTecnico = new List<Cls_Entidad_AsignaOrdenTrabajo.validacion>();
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    //using (SqlCommand cmd = new SqlCommand("NEW_SP_S_LECTURA_ENVIAR_CLIENTE_VALIDACION_LECTURAS", cn))
                    using (SqlCommand cmd = new SqlCommand("NEW_SP_S_LECTURA_ENVIAR_CLIENTE_VALIDACION_LECTURAS_v2", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@estadoNew", SqlDbType.VarChar).Value = estado_new;
                        cmd.Parameters.Add("@resultado", SqlDbType.Int).Value = Convert.ToInt32(resultado);
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = Convert.ToInt32(id_supervisor);
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = Convert.ToInt32(id_operario_supervisor);

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.validacion obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.validacion();

                                obj_entidad.checkeado = false;
                                obj_entidad.estado = Fila["estado"].ToString();
                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.foto = Fila["foto"].ToString();
                                obj_entidad.fecha_lectura = Fila["fecha_lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();

                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.operario = Fila["operario"].ToString();

                                obj_entidad.id_observacion = Fila["id_observacion"].ToString();
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                           
                                obj_entidad.lectura_movil = string.IsNullOrEmpty(Fila["lectura_movil"].ToString()) ? obj_entidad.lectura_movil = null : Convert.ToDecimal(Fila["lectura_movil"].ToString()).ToString("#,##0.00");                                                               


                                obj_entidad.lectura_max = string.IsNullOrEmpty(Fila["lectura_max"].ToString()) ? obj_entidad.Consu_act = null : Convert.ToDecimal(Fila["lectura_max"].ToString()).ToString("#,##0.00");
                                obj_entidad.lectura_min = string.IsNullOrEmpty(Fila["lectura_min"].ToString()) ? obj_entidad.Porcen = null : Math.Round(Convert.ToDecimal(Fila["lectura_min"]), 2).ToString("#,##0");                                
                                obj_entidad.comparativa_max = string.IsNullOrEmpty(Fila["comparativa_max"].ToString()) ? obj_entidad.L_ant = null : Convert.ToDecimal(Fila["comparativa_max"].ToString()).ToString("#,##0.00");                                                             
                                obj_entidad.comparativa_min = string.IsNullOrEmpty(Fila["comparativa_min"].ToString()) ? obj_entidad.Prom_3mes = null : Convert.ToDecimal(Fila["comparativa_min"].ToString()).ToString("#,##0.00");
 

                                obj_entidad.Relecturas = Fila["Relecturas"].ToString();
                                obj_entidad.resultado = Fila["resultado"].ToString();
                                obj_entidad.Observacion_lectura = Fila["Observacion_lectura"].ToString();

                                ListTecnico.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListTecnico;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string Capa_dato_DescargarArchivoTexto_EnvioCliente(int local, int servicio, int estado, string fechaAsigna, int id_supervisor, int id_operario_supervisor)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string _correlativo = "";
                string _correlativo2 = "";
                string _rutafile = "";
                string _nombreServicio = "";
                string _rutaServer = "";

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListData = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("NEW_DSIGE_Enviar_Trabajos_al_Cliente_txt", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@Servicios", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            string[] linesArchivo = new string[dt_detalle.Rows.Count];
                            _nombreServicio = "";
                            int i = 0;
                            var medidor = "";

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                medidor = "";
                                medidor = Fila["MEDIDOR"].ToString();

                                if (medidor.Length < 8)
                                {
                                    medidor = medidor.PadRight(8);
                                }

                                linesArchivo[i] = Fila["ITEM"].ToString() + '\t' + Fila["INSTALACION"].ToString() + '\t' + Fila["EQUIPO"].ToString() + '\t' + medidor + '\t' + Fila["CTACTO"].ToString()
                                                  + '\t' + Fila["FecPlanLectura"].ToString() + '\t' + Fila["FecRealPlanLectura"].ToString() + '\t' + Fila["Hora_Plantilla"].ToString() + '\t' + Fila["Lectura"].ToString() + '\t' + Fila["Nota"].ToString()
                                                  + '\t' + Fila["Cod_Comentario"].ToString() + '\t' + Fila["Comentario"].ToString() + '\t' + Fila["Codigo_Lector"].ToString() + '\t' + Fila["CONTRATISTA"].ToString();
                                i = i + 1;
                            }


                            if (dt_detalle.Rows.Count <= 0)
                            {
                                Resultado = "0|No hay informacion disponible";
                            }
                            else
                            {
                                if (servicio == 1)
                                {
                                    _nombreServicio = "LECTURAS_";
                                }
                                if (servicio == 2)
                                {
                                    _nombreServicio = "RELECTURAS_";
                                }
                                if (servicio == 9)
                                {
                                    _nombreServicio = "RECLAMOS_";
                                }
                                _correlativo = String.Format("{0:ddMMyyyy_hhmmss}.txt", DateTime.Now);

                                _rutafile = System.Web.Hosting.HostingEnvironment.MapPath("~/Upload/" + _nombreServicio + _correlativo);
                                _rutaServer = ConfigurationManager.AppSettings["servidor-archivos"];

                                System.IO.File.WriteAllLines(_rutafile, linesArchivo);
                                Resultado = "1|" + _rutaServer + _nombreServicio + _correlativo;
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Resultado = "-1|" + ex.Message;
            }

            return Resultado;
        }


        public string Capa_dato_DescargarArchivoTexto_cortesReconexiones(int servicio, int estado, string fechaAsigna, int operario, int id_usuario, int tipo, string horainicial, string horafinal)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
 
                string _rutafile = "";
                string _nombreServicio = "";
                string _rutaServer = "";
                string _nombreFile = "";

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListData = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_ENVIAR_CLIENTE_CORTE_RECONEXION_TXT_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@tecnicoAsignado", SqlDbType.Int).Value = operario;
                        cmd.Parameters.Add("@estadoAsignacion", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@tipo", SqlDbType.Int).Value = tipo;

                        cmd.Parameters.Add("@horainicial", SqlDbType.VarChar).Value = horainicial;
                        cmd.Parameters.Add("@horafinal", SqlDbType.VarChar).Value = horafinal;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            if (dt_detalle.Rows.Count <= 0)
                            {
                                Resultado = "0|No hay informacion disponible";
                            }
                            else
                            {
                                string[] linesArchivo = new string[dt_detalle.Rows.Count];
                                _nombreServicio = "";
                                int i = 0;

                                if (servicio == 3)
                                {
                                    if (tipo == 1) ////- EFECTIVA - EXITOSA  
                                    {
                                        _nombreServicio = id_usuario + "CARGA_EFECTIVA_CORTES_" + servicio;
                                    } else if (tipo == 2) { /// NO EFECTIVA - INFRUCTUOSA 
                                        _nombreServicio = id_usuario + "CARGA_NO_EFECTIVA_CORTES_" + servicio;
                                    }

                                    foreach (DataRow Fila in dt_detalle.Rows)
                                    {
                                        linesArchivo[i] = Fila["Aviso"].ToString() + ',' + Fila["tcos"].ToString() + ',' + Fila["dni"].ToString() + ',' + Fila["LECTURA"].ToString() + ',' + Fila["ResultadoGRP"].ToString()
                                                          + ',' + Fila["ResultadoCOD"].ToString() + ',' + Fila["CausaGRP"].ToString() + ',' + Fila["CausaCOD"].ToString() + ',' + Fila["Fecha_aviso"].ToString() + ',' + Fila["HORA"].ToString();
                                        i = i + 1;
                                    }
                                }
                                else if (servicio == 4)
                                {
                                    if (tipo == 1) ////- EFECTIVA - EXITOSA  
                                    {
                                        _nombreServicio = id_usuario + "CARGA_EXITOSA_RECONEXIONES_" + servicio;
                                        foreach (DataRow Fila in dt_detalle.Rows)
                                        {
                                            linesArchivo[i] = Fila["Aviso"].ToString() + ',' + Fila["tcos"].ToString() + ',' + Fila["dni"].ToString() + ',' + Fila["LECTURA"].ToString() + ',' + Fila["nombreInterLocutor_Corte"].ToString() + ',' + Fila["ResultadoGRP"].ToString()
                                                              + ',' + Fila["ResultadoCOD"].ToString() + ',' + Fila["Fecha_aviso"].ToString() + ',' + Fila["HORA"].ToString();
                                            i = i + 1;
                                        }
                                    }
                                    else if (tipo == 2)
                                    { /// NO EFECTIVA - INFRUCTUOSA 
                                        _nombreServicio = id_usuario + "CARGA_INFRUCTUOSA_RECONEXIONES_" + servicio;
                                        foreach (DataRow Fila in dt_detalle.Rows)
                                        {
                                            linesArchivo[i] = Fila["Aviso"].ToString() + ',' + Fila["tcos"].ToString() + ',' + Fila["dni"].ToString() + ',' + Fila["LECTURA"].ToString() + ',' + Fila["nombreInterLocutor_Corte"].ToString() + ',' + Fila["ResultadoGRP"].ToString()
                                                              + ',' + Fila["ResultadoCOD"].ToString() + ',' + Fila["CausaGRP"].ToString() + ',' + Fila["CausaCOD"].ToString() + ',' + Fila["Fecha_aviso"].ToString() + ',' + Fila["HORA"].ToString();
                                            i = i + 1;
                                        }
                                    }
                                }

                                _nombreFile = _nombreServicio + ".csv";

                                _rutafile = System.Web.Hosting.HostingEnvironment.MapPath("~/temp/" + _nombreFile);
                                _rutaServer = ConfigurationManager.AppSettings["Archivos"];

                                FileInfo _file = new FileInfo(_rutafile);
                                if (_file.Exists)
                                {
                                    _file.Delete();
                                }
                                System.IO.File.WriteAllLines(_rutafile, linesArchivo);
                                Resultado = "1|" + _rutaServer + _nombreFile;
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Resultado = "-1|" + ex.Message;
            }

            return Resultado;
        }



        public string Capa_Dato_Set_procesarRecepcionLecturas(int servicio, string fechaAsigna)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_LECTURA_ENVIAR_CLIENTE_Registrar_Lectura", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.ExecuteNonQuery();
                        Resultado = "Ok";
                    }
                }
            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }


        public string Capa_Dato_Set_procesarRecepcion_Trabajos(int servicio, string fechaAsigna)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_PROCESO_RECEPCION", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@idTipoServicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.ExecuteNonQuery();
                        Resultado = "Ok";
                    }
                }
            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }










        public string Capa_Dato_Proceso_Verificacion_Porcentaje(string ListaLecturas, int servicio)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_ENVIAR_CLIENTE_PORCENTAJE_MAYOR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@CodigoLecturas", SqlDbType.VarChar).Value = ListaLecturas;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.ExecuteNonQuery();
                        Resultado = "Ok";
                    }
                }
            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }

 
        public string Capa_Dato_Proceso_Actualizar_lecturas(int id_lectura, string valor_lectura, int id_usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    //using (SqlCommand cmd = new SqlCommand("SP_U_ENVIAR_TRABAJO_CLIENTE_LECT_MANUAL", cn))
                    using (SqlCommand cmd = new SqlCommand("SP_U_ENVIAR_TRABAJO_CLIENTE_LECT_MANUAL_V2", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = id_lectura;
                        cmd.Parameters.Add("@valor_lectura", SqlDbType.VarChar).Value = valor_lectura;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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


        public string Capa_Dato_Proceso_Almacenar_lecturas(int id_usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    //using (SqlCommand cmd = new SqlCommand("SP_U_I_SUMINISTROS_MANUAL", cn))
                    using (SqlCommand cmd = new SqlCommand("SP_I_U_SUMINISTROS_MANUAL_V2", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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

        public object Capa_Dato_cambiarfoto(int id_usuario, int idfotoLectura, string nombrefotoLectura)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            object Resultado;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_U_CAMBIAR_FOTO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.Parameters.Add("@idfotoLectura", SqlDbType.Int).Value = idfotoLectura;
                        cmd.Parameters.Add("@nombrefoto", SqlDbType.VarChar).Value = nombrefotoLectura;

                        cmd.ExecuteNonQuery();

                        Resultado = new
                               {
                                   ok =  true,
                                   mensaje = "../Content/foto/foto/" + nombrefotoLectura
                               };
 
                    }
                }
            }
            catch (Exception ex)
            {
                Resultado = new
                    {
                        ok = false,
                        mensaje = ex.Message
                    };
            }
            return Resultado;
        }

        public string Capa_Dato_Proceso_Almacenar_lecturas_vacias(int id_usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_U_I_SUMINISTROS_MANUAL_VACIAS_V2", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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
               
        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_ListaObservaciones(int TipoServicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_OBSERVACIONES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@tipoServicio", SqlDbType.Int).Value = TipoServicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();

                                obj_entidad.id_Observacion = Convert.ToInt32(Fila["id_Observacion"].ToString());
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                ListObservaciones.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
               
        public string Capa_Dato_set_anulandoFoto(int id_lectura, int id_usuario)
        {
            var Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_U_VALIDACION_LECTURA_ANULAR_FOTO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = id_lectura;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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
               
        public List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> Capa_Dato_Get_Alertas_lecturas(string fecha_asignacion, int id_usuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones> ListObservaciones = new List<Cls_Entidad_AsignaOrdenTrabajo.Observaciones>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_ENVIAR_TRABAJO_CLIENTE_ALERTAS_II", cn))
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
                                Cls_Entidad_AsignaOrdenTrabajo.Observaciones obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Observaciones();
                                
                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.unidad_lectura = Fila["unidad_lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.operario = Fila["operario"].ToString();

                                obj_entidad.id_observacion = Fila["id_observacion"].ToString();
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                obj_entidad.lectura_movil = Fila["lectura_movil"].ToString();
                                obj_entidad.lectura_mxa = Fila["lectura_mxa"].ToString();
                                obj_entidad.lectura_min = Fila["lectura_min"].ToString();

                                obj_entidad.comparativa_max = Fila["comparativa_max"].ToString();
                                obj_entidad.comparativa_min = Fila["comparativa_min"].ToString();
                                obj_entidad.resultado = Fila["resultado"].ToString();
                                                               
                                ListObservaciones.Add(obj_entidad);

                                /// t 

                            }
                        }
                    }
                }

                return ListObservaciones;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
                
        public List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> Capa_Dato_ListandoLectura_Seguimiento(int id_local, int id_servicio, string fechainicial, string fechafinal, int id_usuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente> ListPendientes = new List<Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_REPORTE_SEGUIMIENTO_FOTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@fechainicial", SqlDbType.VarChar).Value = fechainicial;
                        cmd.Parameters.Add("@fechafinal", SqlDbType.VarChar).Value = fechafinal;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                              
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.VisualizarPendiente();

                                obj_entidad.fecha =  Fila["fecha"].ToString();
                                obj_entidad.nota_lectura = Fila["nota_lectura"].ToString();
                                obj_entidad.estimacion = Fila["estimacion"].ToString();

                                obj_entidad.parametros = Fila["parametros"].ToString();
                                obj_entidad.comercio = Fila["comercio"].ToString();
                                obj_entidad.total = Fila["total"].ToString();

                                ListPendientes.Add(obj_entidad);
                            }
                        }
                    }
                }
                return ListPendientes;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        public string Capa_Negocio_ListandoFotosDescargar_seguimiento(int local, int servicio, string fecha, string option, int id_usuario)
        {
            var resultado = "";
            var pathExist = "";
            var nombreFile = "";
            var nombreFoto = "";
            string ruta_descarga = ConfigurationManager.AppSettings["Archivos"];

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];
                string[] fileEntries = null; 
                int ac = 0;

                if (option=="N")
                {
                    nombreFile = id_usuario + "_notalectura_fotos.zip";
                }
                else if (option == "E")
                {
                    nombreFile = id_usuario + "_estimacion_fotos.zip";
                }
                else if (option == "P")
                {
                    nombreFile = id_usuario + "_parametros_fotos.zip";
                }
                else if (option == "C")
                {
                    nombreFile = id_usuario + "_comercio_fotos.zip";
                }

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListFotos = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_REPORTE_SEGUIMIENTO_DESCARGAR_FOTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = option;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            ac = 0;                             

                            if (dt_detalle.Rows.Count > 0)
                            {
                                if (dt_detalle.Rows.Count > 0)
                                {
                                    fileEntries = new string[dt_detalle.Rows.Count];
                                }

                                foreach (DataRow Fila in dt_detalle.Rows)
                                {
                                    Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio();
                                    obj_entidad.fotourl = ruta + Convert.ToString(Fila["fotourl"]);
                                    ListFotos.Add(obj_entidad);

                                    //almacenando un matriz
                                    pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                    if (System.IO.File.Exists(pathExist))
                                    {
                                        fileEntries[ac] = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/" + Convert.ToString(Fila["fotourl"]));
                                        ac = ac + 1;
                                    }
                                    else {
                                        //-----almacenando la url de la foto para borrarlo luego
                                        nombreFoto = Convert.ToString(Fila["fotourl"]);
                                        using (SqlCommand cmds = new SqlCommand("sp_i_fotos_error", cn))
                                        {
                                            cmds.CommandTimeout = 0;
                                            cmds.CommandType = CommandType.StoredProcedure;
                                            cmds.Parameters.Add("@ruta_foto", SqlDbType.VarChar).Value = nombreFoto;
                                            cmds.ExecuteNonQuery();
                                        }
                                    }
                                }

                                List<string> listFotos = new List<string>(fileEntries);

                                for (int index = 0; index < listFotos.Count; index++)
                                {
                                    bool nullOrEmpty = string.IsNullOrEmpty(listFotos[index]);
                                    if (nullOrEmpty)
                                    {
                                        listFotos.RemoveAt(index);
                                        --index;
                                    }
                                }                                

                                pathExist = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/"+ nombreFile);

                                if (System.IO.File.Exists(pathExist))
                                {
                                    System.IO.File.Delete(pathExist);
                                }

                                //Guardando el archivo en Zip                
                                using (ZipFile zip = new ZipFile())
                                {
                                    zip.AddFiles(listFotos, "");
                                    zip.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/" + nombreFile));
                                }
                                resultado = "OK|" + ruta_descarga + nombreFile;
                            }
                            else {
                                resultado = "0|No hay informacion para mostrar..";
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

        public string Capa_Negocio_Listando_Data_Descargar_seguimiento(int local, int servicio, string fecha, string option, int id_usuario)
        {
            var resultado = "";
            var nombreFile = "";
            var nombreExcel = "";
            var servidor = "";

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];


                //servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
                servidor = "";
                if (option == "N")
                {
                    nombreFile = id_usuario + "_notalectura_data_" + servidor +".xlsx";
                    nombreExcel = "Notalectura";
                }
                else if (option == "E")
                {
                    nombreFile = id_usuario + "_estimacion_data_" + servidor +".xlsx";
                    nombreExcel = "Estimacion";
                }
                else if (option == "P")
                {
                    nombreFile = id_usuario + "_parametros_data_" + servidor + ".xlsx";
                    nombreExcel = "Parametro";
                }
                else if (option == "C")
                {
                    nombreFile = id_usuario + "_comercio_data_" + servidor + ".xlsx";
                    nombreExcel = "Comercio";
                }

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_REPORTE_SEGUIMIENTO_DESCARGAR_DATA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@opcion", SqlDbType.VarChar).Value = option;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;

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

                    for (int i = 1; i <= 2; i++)
                    {
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    oWs.Cells[1, 1].Value = "#";
                    oWs.Cells[1, 2].Value = "Suministro";

                    int ac = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        ac += 1;
                        for (int j = 1; j <= 2; j++)
                        {
                            oWs.Cells[_fila, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }

                        oWs.Cells[_fila, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto  
                        oWs.Cells[_fila, 1].Value = ac;
                        oWs.Cells[_fila, 2].Value = oBj["suministro_lectura"].ToString();
                        oWs.Cells[_fila, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 2].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto
                        
                        _fila++;
                    }

                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Column(1).Style.Font.Bold = true;

                    for (int k = 1; k <= 2; k++)
                    {
                        oWs.Column(k).AutoFit();
                    }
                    oEx.Save();
                }

                Res ="1|" + ruta_descarga +  nombreFile;
            }
            catch (Exception ex)
            {
                Res = "0|" + ex.Message;
            }
            return Res;
        }
               
        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaSupervisor()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SUPERVISOR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {                            
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_Usuario = Convert.ToInt32(Fila["id_Usuario"].ToString());
                                obj_entidad.supervisor = Fila["supervisor"].ToString();
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

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaSupervisor_usuario( int id_usuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SUPERVISOR_USUARIO", cn))
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
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_Usuario = Convert.ToInt32(Fila["id_Usuario"].ToString());
                                obj_entidad.supervisor = Fila["supervisor"].ToString();
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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaSupervisor_usuario_validacion(int id_usuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_SUPERVISOR_USUARIO_VALIDACION", cn))
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
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_Usuario = Convert.ToInt32(Fila["id_Usuario"].ToString());
                                obj_entidad.supervisor = Fila["supervisor"].ToString();
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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaOperarios()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
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
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.checkeado = false;
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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaUsuarioOperario(int id_usuario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListServicio = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_USUARIO_RESPONSABLE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_usuario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();


                                obj_entidad.checkeado = false;
                                obj_entidad.id_Usuario_Responsable = Convert.ToInt32(Fila["id_Usuario_Responsable"].ToString());
                                obj_entidad.id_Usuario = Convert.ToInt32(Fila["id_Usuario"].ToString());
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

        public string Capa_Dato_set_agregarOperario(string objOperarios, int id_supervisor, int id_usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_USUARIO_RESPONSABLE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@objOperarios", SqlDbType.VarChar).Value = objOperarios;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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

        public string Capa_Dato_set_eliminarOperario(string objOperarios, int id_supervisor, int id_usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_D_USUARIO_RESPONSABLE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@objOperarios", SqlDbType.VarChar).Value = objOperarios;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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

        public string Capa_Dato_set_anulando_unidad_lectura(string Cod_UnidadLectura, int id_usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_UNIDAD_LECTURA_ANULAR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Cod_UnidadLectura", SqlDbType.VarChar).Value = Cod_UnidadLectura;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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
                     

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaAnio()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListDato = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ANIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_anio = Fila["id_anio"].ToString();
                                obj_entidad.descripcion_anio = Fila["descripcion_anio"].ToString();
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

        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_ListaMes()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListDato = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_MES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_mes = Fila["id_mes"].ToString();
                                obj_entidad.descripcion_mes = Fila["descripcion_mes"].ToString();
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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_UnidadLectura()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListDato = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_UNIDAD_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.Cod_UnidadLectura = Fila["Cod_UnidadLectura"].ToString();
                                obj_entidad.nombre_UnidadLectura = Fila["nombre_UnidadLectura"].ToString();
                                obj_entidad.Distrito_UnidadLectura = Fila["Distrito_UnidadLectura"].ToString();

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


        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_UnidadLectura_principal()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListDato = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_MANTENIMIENTO_UNIDAD_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.Cod_UnidadLectura = Fila["Cod_UnidadLectura"].ToString();
                                obj_entidad.nombre_UnidadLectura = Fila["nombre_UnidadLectura"].ToString();
                                obj_entidad.Distrito_UnidadLectura = Fila["Distrito_UnidadLectura"].ToString();

                                obj_entidad.Estado = Fila["Estado"].ToString();
                                obj_entidad.usuario_creacion = Fila["usuario_creacion"].ToString();
                                obj_entidad.fecha_creacion = Fila["fecha_creacion"].ToString();
                                obj_entidad.usuario_modificacion = Fila["usuario_modificacion"].ToString();
                                obj_entidad.fecha_modificacion = Fila["fecha_modificacion"].ToString();

                                obj_entidad.usuario_registra = Fila["usuario_registra"].ToString();
                                obj_entidad.usuario_edicion = Fila["usuario_edicion"].ToString();

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

               
        public string Capa_Dato_set_Insert_Update_Configuracion(int id_Configuracion_UL, int dia_configuracion_ul, string cod_unidadlectura, int id_usuario_responsable, int id_tiposervicio, int estado , int usuario_creacion)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_U_CONFIGURACION_UL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Configuracion_UL", SqlDbType.Int).Value = id_Configuracion_UL;
                        cmd.Parameters.Add("@dia_configuracion_ul", SqlDbType.Int).Value = dia_configuracion_ul;
                        cmd.Parameters.Add("@cod_unidadlectura", SqlDbType.VarChar).Value = cod_unidadlectura;
                        cmd.Parameters.Add("@id_usuario_responsable", SqlDbType.Int).Value = id_usuario_responsable;

                        cmd.Parameters.Add("@id_tiposervicio", SqlDbType.VarChar).Value = id_tiposervicio;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@usuario_creacion", SqlDbType.Int).Value = usuario_creacion;

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


        public string Capa_Dato_set_Insert_Update_Unidad_lectura(string codigo, string nombre, string distrito, string estado, int usuario)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_U_UNIDAD_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@codigo", SqlDbType.VarChar).Value = codigo;
                        cmd.Parameters.Add("@nombre", SqlDbType.VarChar).Value = nombre;
                        cmd.Parameters.Add("@distrito", SqlDbType.VarChar).Value = distrito;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;

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

            
        public List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> Capa_Dato_Get_Configuracion_ul(string dia_config, int id_supervisor , int id_servicio, int anio, int mes)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_AsignaOrdenTrabajo.Servicios> ListDato = new List<Cls_Entidad_AsignaOrdenTrabajo.Servicios>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_CONFIGURACION_UL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@dia_config", SqlDbType.VarChar).Value = dia_config;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@anio", SqlDbType.Int).Value = anio;
                        cmd.Parameters.Add("@mes", SqlDbType.Int).Value = mes;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_AsignaOrdenTrabajo.Servicios obj_entidad = new Cls_Entidad_AsignaOrdenTrabajo.Servicios();

                                obj_entidad.id_Configuracion_UL = Fila["id_Configuracion_UL"].ToString();
                                obj_entidad.Dia_Configuracion_UL = Fila["Dia_Configuracion_UL"].ToString();
                                obj_entidad.Cod_UnidadLectura = Fila["Cod_UnidadLectura"].ToString();
                                obj_entidad.Distrito_UnidadLectura = Fila["Distrito_UnidadLectura"].ToString();
                                obj_entidad.id_Usuario_Responsable = Convert.ToInt32(Fila["id_Usuario_Responsable"].ToString());
                                obj_entidad.responsable = Fila["responsable"].ToString();
                                obj_entidad.id_TipoServicio =Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.cantidad_lectura = Fila["cantidad_lectura"].ToString();

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

        public string Capa_Dato_set_generando_relectura(int servicio, string fecha_asignacion, int id_usuario)
        {
            var Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_U_ENVIAR_TRABAJOS_CERRAR_DIA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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


        public string Capa_dato_DescargarArchivoTexto_EnvioCliente_Relectura(int local, int servicio, int estado, string fechaAsigna, int id_supervisor, int id_operario_supervisor)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string _correlativo = "";
                string _rutafile = "";
                string _nombreServicio = "";
                string _rutaServer = "";

                List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio> ListData = new List<Cls_Entidad_AsignaOrdenTrabajo.LecturaEnvio>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("NEW_DSIGE_Enviar_Trabajos_al_Cliente_txt_RELECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@Servicios", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                        cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            string[] linesArchivo = new string[dt_detalle.Rows.Count];
                            _nombreServicio = "";
                            int i = 0;
                            var medidor = "";

                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                medidor = "";
                                medidor = Fila["MEDIDOR"].ToString();

                                if (medidor.Length < 8)
                                {
                                    medidor = medidor.PadRight(8);
                                }

                                linesArchivo[i] = Fila["ITEM"].ToString() + '\t' + Fila["INSTALACION"].ToString() + '\t' + Fila["EQUIPO"].ToString() + '\t' + medidor + '\t' + Fila["CTACTO"].ToString()
                                                  + '\t' + Fila["FecPlanLectura"].ToString() + '\t' + Fila["FecRealPlanLectura"].ToString() + '\t' + Fila["Hora_Plantilla"].ToString() + '\t' + Fila["Lectura"].ToString() + '\t' + Fila["Nota"].ToString()
                                                  + '\t' + Fila["Cod_Comentario"].ToString() + '\t' + Fila["Comentario"].ToString() + '\t' + Fila["Codigo_Lector"].ToString();
                                i = i + 1;
                            }


                            if (dt_detalle.Rows.Count <= 0)
                            {
                                Resultado = "0|No hay informacion disponible";
                            }
                            else
                            {
                                if (servicio == 1)
                                {
                                    _nombreServicio = "LECTURAS_";
                                }
                                if (servicio == 2)
                                {
                                    _nombreServicio = "RELECTURAS_";
                                }
                                _correlativo = String.Format("{0:ddMMyyyy_hhmmss}.txt", DateTime.Now);

                                _rutafile = System.Web.Hosting.HostingEnvironment.MapPath("~/Upload/" + _nombreServicio + _correlativo);
                                _rutaServer = ConfigurationManager.AppSettings["servidor-archivos"];

                                System.IO.File.WriteAllLines(_rutafile, linesArchivo);
                                Resultado = "1|" + _rutaServer + _nombreServicio + _correlativo;
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Resultado = "-1|" + ex.Message;
            }

            return Resultado;
        }
               

        public object Capa_Dato_get_grandesClientes(int estado, string fecha_inicial, string fecha_final, string codigoEmr)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_GRANDES_CLIENTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fecha_inicial", SqlDbType.VarChar).Value = fecha_inicial;
                        cmd.Parameters.Add("@fecha_final", SqlDbType.VarChar).Value = fecha_final;
                        cmd.Parameters.Add("@codigoEmr", SqlDbType.VarChar).Value = codigoEmr;
                        
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return dt_detalle;
        }                     
        

        public object Capa_Dato_get_ultimoCodigoEmr(int usuario)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_S_ULTIMO_COD_EMR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@idUsuario", SqlDbType.Int).Value = usuario;
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return dt_detalle;
        }

        public object Capa_Dato_download_grandesClientes(int estado, string fecha_inicial, string fecha_final, string codigoEmr, int id_usuario)
        {

            string resultado;
            string nombreFile = id_usuario + "_registroTomaLecturas_GrandesClientes.xlsx";

            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_GRANDES_CLIENTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@fecha_inicial", SqlDbType.VarChar).Value = fecha_inicial;
                        cmd.Parameters.Add("@fecha_final", SqlDbType.VarChar).Value = fecha_final;
                        cmd.Parameters.Add("@codigoEmr", SqlDbType.VarChar).Value = codigoEmr;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            if (dt_detalle.Rows.Count <= 0)
                            {
                                resultado = "0|No hay informacion disponible";
                            }
                            else
                            {
                                resultado = GenerarArchivoExcel_grandesClientes(dt_detalle, nombreFile, "grandesClientes");
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


        public string GenerarArchivoExcel_grandesClientes(DataTable dt_detalles, string nombreFile, string nombreExcel)
        {
            string _ruta = "";
            string Res = "";
            int _fila = 11;
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
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 8));

                                       
                    oWs.Cells[4, 1, 4, 14].Merge = true;  // combinar celdaS dt
                    oWs.Cells[4, 1].Value = "REGISTRO DE TOMA DE LECTURA DE MEDIDORES - GRANDES CLIENTES";
                    oWs.Cells[4, 1].Style.Font.Size = 15; //letra tamaño  
                    oWs.Cells[4, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[4, 1].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[4, 1].Style.Font.Bold = true; //Letra negrita


                    for (int i = 1; i <= 14; i++)
                    {
                        oWs.Cells[9, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        oWs.Cells[10, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    oWs.Cells[9, 1, 9, 7].Merge = true;  // combinar celdaS dt
                    oWs.Cells[9, 1].Value = "PUNTO DE LECTURA";
                    oWs.Cells[9, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[9, 1].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[9, 1].Style.Font.Bold = true; //Letra negrita

                    oWs.Cells[9, 8].Value = "MEDIDOR";
                    oWs.Cells[9, 8].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[9, 8].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[9, 8].Style.Font.Bold = true; //Letra negrita

                    oWs.Cells[9, 9, 9, 13].Merge = true;  // combinar celdaS dt
                    oWs.Cells[9, 9].Value = "CORRECTOR";
                    oWs.Cells[9, 9].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Cells[9, 9].Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Cells[9, 9].Style.Font.Bold = true; //Letra negrita



                    oWs.Cells[10, 1].Value = "NRO";
                    oWs.Cells[10, 2].Value = "CODIGO ERM";
                    oWs.Cells[10, 3].Value = "NOMBRE DEL CLIENTE";
                    oWs.Cells[10, 4].Value = "FECHA";
                    oWs.Cells[10, 5].Value = "HORA";

                    oWs.Cells[10, 6].Value = "NIVEL DE LEL(%) ";
                    oWs.Cells[10, 7].Value = "PRESIÓN ENTRADA(BAR)";
                    oWs.Cells[10, 8].Value = "CONTADOR MECÁNICO(m3)";
                    oWs.Cells[10, 9].Value = "VOL.ACC.SIN CORREGIR(m3)";

                    oWs.Cells[10, 10].Value = "VOL.ACC.CORREGIDO(Std m3)";
                    oWs.Cells[10, 11].Value = "PRESIÓN(Bara)";

                    oWs.Cells[10, 12].Value = "TEMP.(ºC)";
                    oWs.Cells[10, 13].Value = "TIEMPO BATERÍA(Días/%)";
                    oWs.Cells[10, 14].Value = "OBSERVACIONES";
 

                    int ac = 0;
                    foreach (DataRow oBj in dt_detalles.Rows)
                    {
                        ac += 1;

                        oWs.Cells[_fila, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        oWs.Cells[_fila, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto  
                        oWs.Cells[_fila, 1].Value = ac;

                        oWs.Cells[_fila, 2].Value = oBj["codigo_erm"].ToString();
                        oWs.Cells[_fila, 3].Value = oBj["nombre_cliente"].ToString();
                        oWs.Cells[_fila, 4].Value = oBj["fecha"].ToString();
                        oWs.Cells[_fila, 5].Value = oBj["hora"].ToString();

                        oWs.Cells[_fila, 6].Value = oBj["nivelel_porc"].ToString();
                        oWs.Cells[_fila, 7].Value = oBj["presion_entrada_bar"].ToString();
                        oWs.Cells[_fila, 8].Value = oBj["contador_mecanico_m3"].ToString();
                        oWs.Cells[_fila, 9].Value = oBj["vol_acc_sin_corregir_m3"].ToString();

                        oWs.Cells[_fila, 10].Value = oBj["vol_acc_corregido_std_m3"].ToString();
                        oWs.Cells[_fila, 11].Value = oBj["presion_bara"].ToString();

                        oWs.Cells[_fila, 12].Value = oBj["temp"].ToString();
                        oWs.Cells[_fila, 13].Value = oBj["tiempo_bateria_dias"].ToString();
                        oWs.Cells[_fila, 14].Value = oBj["observaciones"].ToString();

                        _fila++;
                    }

                    oWs.Row(10).Style.Font.Bold = true;
                    oWs.Row(10).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(10).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                    oWs.Column(1).Style.Font.Bold = true;

                    for (int k = 1; k <= 14; k++)
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




        public object Capa_Dato_get_grandesClientes_detalle(int Id_GrandeCliente)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_GRANDES_CLIENTES_DETALLE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Id_GrandeCliente", SqlDbType.Int).Value = Id_GrandeCliente;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return dt_detalle;
        }


        public class download {
            public string nombreFile { get; set; }
            public string nombreBd { get; set; }
            public string ubicacion { get; set; }
        }

        public string Capa_Dato_get_download_grandesClientes(int Id_GrandeCliente, int tipo, int id_usuario)
        {
            DataTable dt_detalle = new DataTable();
            List<download> list_files = new List<download>();
            string rutaFoto = "";
            string rutaOrig = "";
            string rutaDest = "";
            string nombreArchivoReal = "";
            string ruta_descarga = "";

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_GRANDES_CLIENTES_DESCARGAR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Id_GrandeCliente", SqlDbType.Int).Value = Id_GrandeCliente;
                        cmd.Parameters.Add("@tipo", SqlDbType.Int).Value = tipo;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);


                            if (tipo == 1) /// fotos
                            {
                                rutaFoto = System.Web.Hosting.HostingEnvironment.MapPath("~/Content/foto/foto/");
                                foreach (DataRow row in dt_detalle.Rows)
                                {
                                    foreach (DataColumn column in dt_detalle.Columns)
                                    {
                                        download obj_entidad = new download();
                                        obj_entidad.nombreFile = row[column].ToString().Replace('"', ' ').Trim();
                                        obj_entidad.ubicacion = rutaFoto;
                                        list_files.Add(obj_entidad);
                                    }
                                }
                            }
                            else { /// archivos  
    
                                rutaFoto = System.Web.Hosting.HostingEnvironment.MapPath("~/Files_GrandesClientes/Descargas/");

                                foreach (DataRow Fila in dt_detalle.Rows)
                                {
                                    download obj_entidad = new download();
                                    obj_entidad.nombreFile = Fila["archivo"].ToString();
                                    obj_entidad.nombreBd = Fila["nombreBD"].ToString();
                                    
                                    obj_entidad.ubicacion = rutaFoto;
                                    list_files.Add(obj_entidad);
                                }

                                ////restaurando el archivo...
                                foreach (download item in list_files)
                                {
                                    nombreArchivoReal = "";
                                    nombreArchivoReal = item.nombreBd.Replace(item.nombreBd, item.nombreFile);

                                    rutaOrig = System.Web.Hosting.HostingEnvironment.MapPath("~/Files_GrandesClientes/" + item.nombreBd);  
                                    rutaDest = System.Web.Hosting.HostingEnvironment.MapPath("~/Files_GrandesClientes/Descargas/" + nombreArchivoReal);

                                    if (System.IO.File.Exists(rutaDest)) //--- restaurarlo
                                    {
                                        System.IO.File.Delete(rutaDest);
                                        System.IO.File.Copy(rutaOrig, rutaDest);
                                    }
                                    else
                                    {
                                        System.IO.File.Copy(rutaOrig, rutaDest);
                                    }
                                    Thread.Sleep(1000);
                                }
                            }

                            if (list_files.Count > 0)
                            {
                                if (list_files.Count == 1)
                                { 
                                    if (tipo == 1)
                                    {
                                        ruta_descarga = "1|" + ConfigurationManager.AppSettings["Archivos"] + "Descargas/" + list_files[0].nombreFile;
                                    }
                                    else {
                                        ruta_descarga = "1|" + ConfigurationManager.AppSettings["ArchivosFile"] + "Descargas/" + list_files[0].nombreFile;
                                    }
                                }
                                else {
                                    ruta_descarga = "1|" + comprimir_Files(list_files, id_usuario, tipo);
                                }
                            }
                            else {
                                throw new System.InvalidOperationException("No hay informacion para mostrar");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ruta_descarga = "0|" + ex.Message;
            }
            return ruta_descarga;
        }



        public string comprimir_Files(List<download> list_download, int usuario_creacion, int tipo)
        {
            string resultado = "";
            try
            {
                string ruta_destino =  "";
                string ruta_descarga = "";

                if (tipo == 1) ///--fotos
                {
                    ruta_destino = System.Web.Hosting.HostingEnvironment.MapPath("~/Temp/Descargas/" + usuario_creacion + "_FotosDescarga.zip");
                    ruta_descarga = ConfigurationManager.AppSettings["Archivos"] + "Descargas/" + usuario_creacion + "_FotosDescarga.zip";
                }
                else {
                    ruta_destino = System.Web.Hosting.HostingEnvironment.MapPath("~/Files_GrandesClientes/Descargas/" + usuario_creacion + "_ArchivosDescarga.zip");
                    ruta_descarga = ConfigurationManager.AppSettings["ArchivosFile"] + "Descargas/" + usuario_creacion + "_ArchivosDescarga.zip";
                }

                if (File.Exists(ruta_destino)) //--- restaurarlo
                {
                    System.IO.File.Delete(ruta_destino);
                }
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    foreach (download item in list_download)
                    {
                        var nombre = item.ubicacion + item.nombreFile;
                        if (System.IO.File.Exists(item.ubicacion + item.nombreFile))
                        {
                            zip.AddFile(item.ubicacion + item.nombreFile, "");
                        }
                    }
                    // Guardando el archivo zip 
                    zip.Save(ruta_destino);
                }
                Thread.Sleep(2000);

                if (File.Exists(ruta_destino))
                {
                    resultado = ruta_descarga;
                }
                else
                {
                    throw new System.InvalidOperationException("No se pudo generar la Descarga del Archivo");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return resultado;
        }




        public object Capa_Dato_get_grandesClientes_detalleFile(int Id_GrandeCliente)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_GRANDES_CLIENTES_DETALLE_FILE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Id_GrandeCliente", SqlDbType.Int).Value = Id_GrandeCliente;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return dt_detalle;
        }


    }
}
