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
namespace DSIGE.Dato
{
    public class Cls_Dato_RecepcionLect
    {
        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Dato_Get_GerReporteFueraRangoCorLds(int id_servicio, string id_sector, string fecha_asignacion,int tipo)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_ReproteFueraRangoCorLds> ListServicio = new List<Cls_Entidad_ReproteFueraRangoCorLds>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Lectura_Calculo_GSP", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Sector", SqlDbType.VarChar).Value = id_sector;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@Tipo", SqlDbType.Int).Value = tipo;
                        
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_ReproteFueraRangoCorLds obj_entidad = new Cls_Entidad_ReproteFueraRangoCorLds();
                                if (tipo == 1)
                                {
                                     obj_entidad.zona_lectura = Fila["zona_lectura"].ToString();
                                }
                                else {
                                    obj_entidad.Observacion = Fila["Observacion"].ToString();
                                
                                }
                                
                                obj_entidad.Cantidad = Fila["Cantidad"].ToString();
                                obj_entidad.FueraRango = Fila["FueraRango"].ToString();
                                obj_entidad.DentroRango = Fila["DentroRango"].ToString();
                                obj_entidad.PFueraRango = Fila["PFueraRango"].ToString();
                                obj_entidad.PDentroRango = Fila["PDentroRango"].ToString();
                                
                                
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









        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Dato_Get_GerReporteFueraRangoCorLdsNot(int id_servicio, string id_sector, string fecha_asignacion)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_ReproteFueraRangoCorLds> ListServicio = new List<Cls_Entidad_ReproteFueraRangoCorLds>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Lectura_Calculo_GSP_NOT", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Sector", SqlDbType.VarChar).Value = id_sector;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_ReproteFueraRangoCorLds obj_entidad = new Cls_Entidad_ReproteFueraRangoCorLds();
                                obj_entidad.id_operario_lectura = Fila["id_operario_lectura"].ToString();
                                obj_entidad.Cantidad = Fila["Cantidad"].ToString();
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






        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Dato_Get_GerReporteFueraRangoCorLds_Det(int id_servicio, string id_sector, string fecha_asignacion, string zona_lectura, int tipo)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_ReproteFueraRangoCorLds> ListServicio = new List<Cls_Entidad_ReproteFueraRangoCorLds>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Lectura_Calculo_GSP_Detalle", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Sector", SqlDbType.VarChar).Value = id_sector;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@zona_lectura", SqlDbType.VarChar).Value = zona_lectura;
                        cmd.Parameters.Add("@tipo", SqlDbType.Int).Value = tipo;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_ReproteFueraRangoCorLds obj_entidad = new Cls_Entidad_ReproteFueraRangoCorLds();

                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.Operario = Fila["Operario"].ToString();
                                obj_entidad.fechaLecturaMovil_lectura = Fila["fechaLecturaMovil_lectura"].ToString();
                                obj_entidad.confirmacion_Lectura = Fila["confirmacion_Lectura"].ToString();
                                obj_entidad.Observacion_lectura = Fila["Observacion_lectura"].ToString();
                                obj_entidad.latitudMovil_lectura = Fila["latitudMovil_lectura"].ToString();
                                obj_entidad.longitudMovil_lectura = Fila["longitudMovil_lectura"].ToString();
                                obj_entidad.Metros = Fila["Metros"].ToString();
                                obj_entidad.Ver = Fila["Ver"].ToString();                               


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







        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Dato_Get_GerReporteFueraRangoCorLds_DetNot(int id_servicio, string id_sector, string fecha_asignacion, int id_operario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_ReproteFueraRangoCorLds> ListServicio = new List<Cls_Entidad_ReproteFueraRangoCorLds>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Lectura_Calculo_GSP_Detalle_NOT", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Sector", SqlDbType.VarChar).Value = id_sector;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@Id_Operario", SqlDbType.Int).Value = id_operario;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_ReproteFueraRangoCorLds obj_entidad = new Cls_Entidad_ReproteFueraRangoCorLds();

                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.Operario = Fila["Operario"].ToString();
                                obj_entidad.fechaLecturaMovil_lectura = Fila["fechaLecturaMovil_lectura"].ToString();
                                obj_entidad.confirmacion_Lectura = Fila["confirmacion_Lectura"].ToString();
                                obj_entidad.Observacion_lectura = Fila["Observacion_lectura"].ToString();
                                obj_entidad.latitudMovil_lectura = Fila["latitudMovil_lectura"].ToString();
                                obj_entidad.longitudMovil_lectura = Fila["longitudMovil_lectura"].ToString();
                                obj_entidad.Metros = Fila["Metros"].ToString();
                                obj_entidad.tieneFoto_lectura = Fila["tieneFoto_lectura"].ToString();
                                obj_entidad.URL = Fila["URL"].ToString();

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



        public List<Cls_Entidad_ReproteFueraRangoCorLds> Capa_Dato_Get_GerReporteFueraRangoCorLds_DetObs(int id_servicio, string id_sector, string fecha_asignacion, string zona_lectura, int tipo)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_ReproteFueraRangoCorLds> ListServicio = new List<Cls_Entidad_ReproteFueraRangoCorLds>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Lectura_Calculo_GSP_DetalleObs", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Sector", SqlDbType.VarChar).Value = id_sector;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@zona_lectura", SqlDbType.VarChar).Value = zona_lectura;
                        cmd.Parameters.Add("@tipo", SqlDbType.Int).Value = tipo;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_ReproteFueraRangoCorLds obj_entidad = new Cls_Entidad_ReproteFueraRangoCorLds();

                                obj_entidad.Observacion = Fila["Observacion"].ToString();
                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.Operario = Fila["Operario"].ToString();
                                obj_entidad.fechaLecturaMovil_lectura = Fila["fechaLecturaMovil_lectura"].ToString();
                                obj_entidad.confirmacion_Lectura = Fila["confirmacion_Lectura"].ToString();
                                obj_entidad.Observacion_lectura = Fila["Observacion_lectura"].ToString();
                                obj_entidad.latitudMovil_lectura = Fila["latitudMovil_lectura"].ToString();
                                obj_entidad.longitudMovil_lectura = Fila["longitudMovil_lectura"].ToString();
                                obj_entidad.Metros = Fila["Metros"].ToString();
                                obj_entidad.Ver = Fila["Ver"].ToString();


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
        public List<Cls_Entidad_RecepcionLect> Capa_Dato_Get_ListaServicio()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_RecepcionLect> ListServicio = new List<Cls_Entidad_RecepcionLect>();
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
                                Cls_Entidad_RecepcionLect obj_entidad = new Cls_Entidad_RecepcionLect();

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

        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }
        public string Capa_Dato_GuardarTemporalSuministro(List<Cls_Entidad_Temporal_Sum> listLecturas, int tipoServicio, string fecha_asignacion)
        {
            string Resultado = "";

            try
            {
                DataTable dt = new DataTable();
                dt = ConvertToDataTable(listLecturas);
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_DELETE_TEMPORAL_SUMINISTRO", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = listLecturas[0].Usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "Tmp_Asigna_Operario";
                        bulkCopy.WriteToServer(dt);

                    }
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Seguimiento_Lecturas_GuardaAsignacion_new", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = listLecturas[0].Usuario;
                        cmd.Parameters.Add("@id_Operario", SqlDbType.Int).Value = listLecturas[0].id_Operario;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@TipoServicio", SqlDbType.Int).Value = tipoServicio;
                        cmd.ExecuteNonQuery();
                        Resultado = "success";
                    }
                }


            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }



        public string Capa_Dato_AnularLectura(int id_lectura)
        {
            string Resultado = "";

            try
            {
              
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_ANULAR_LECTURA", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = id_lectura;
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



        public string Capa_Dato_GenerarSecuencia(int idLecturista, int idServicios, string idSector, string FechaInicial, int Ano, int Mes)
        {
            string Resultado = "";
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Actualizar_Secuencia_Operario", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@TipoServicio", SqlDbType.Int).Value = idServicios;
                        cmd.Parameters.Add("@FechaAsignacion", SqlDbType.VarChar).Value = FechaInicial;
                        cmd.Parameters.Add("@Sector", SqlDbType.VarChar).Value = idSector.ToString();
                        cmd.Parameters.Add("@IdOperario", SqlDbType.Int).Value = idLecturista;
                        cmd.Parameters.Add("@Año", SqlDbType.Int).Value = Ano;
                        cmd.Parameters.Add("@Mes", SqlDbType.Int).Value = Mes;
                        cmd.ExecuteNonQuery();
                        Resultado = "success";
                    }
                }


            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }


        public List<Cls_Entidad_Lecturas_Dia> Capa_Dato_Get_ListadoLecturasDia(int id_local, string fecha, int id_servicio, int tipoProceso, int pendiente, int asignado, int verificado)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_Lecturas_Dia> ListLectura = new List<Cls_Entidad_Lecturas_Dia>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Seguimiento_Lecturas_Dias", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@TipoProceso", SqlDbType.Int).Value = tipoProceso;
                        cmd.Parameters.Add("@ChekPendientes", SqlDbType.Int).Value = pendiente;
                        cmd.Parameters.Add("@ChekAsignados", SqlDbType.Int).Value = asignado;
                        cmd.Parameters.Add("@ChekVerificados", SqlDbType.Int).Value = verificado;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Dia obj_entidad = new Cls_Entidad_Lecturas_Dia();

                                obj_entidad.Tipo = Fila["Tipo"].ToString();
                                obj_entidad.Estado = Convert.ToInt32(Fila["Estado"].ToString());
                                obj_entidad.id_tipoServicio = Convert.ToInt32(Fila["id_tipoServicio"].ToString());
                                obj_entidad.id_lectura = Convert.ToInt32(Fila["id_lectura"].ToString());
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.Color = Fila["Color"].ToString();
                                obj_entidad.Medidor = Fila["Medidor"].ToString();
                                obj_entidad.Operador = Fila["Operador"].ToString();
                                obj_entidad.Direccion = Fila["Direccion"].ToString();
                                obj_entidad.Cliente = Fila["Cliente"].ToString();
                                obj_entidad.Lectura = Fila["Lectura"].ToString();
                                obj_entidad.FechaHora = Fila["FechaHora"].ToString();
                                obj_entidad.TomaFoto = Fila["TomaFoto"].ToString();
                                obj_entidad.Url = Fila["Url"].ToString();
                                

                                ListLectura.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListLectura;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Cls_Entidad_Lectura_OperarioZona> Capa_Dato_Get_ListadoOperariosZona(int id_local, string fecha, int id_servicio, int tipoProceso)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_Lectura_OperarioZona> ListLectura = new List<Cls_Entidad_Lectura_OperarioZona>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_LISTA_OPERARIOS_ZONA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@TipoProceso", SqlDbType.Int).Value = tipoProceso;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lectura_OperarioZona obj_entidad = new Cls_Entidad_Lectura_OperarioZona();

                                obj_entidad.id_operario_lectura = Convert.ToInt32(Fila["id_operario_lectura"].ToString());
                                obj_entidad.cant_zona = Convert.ToInt32(Fila["cant_zona"].ToString());
                                obj_entidad.cant_sum = Convert.ToInt32(Fila["cant_sum"].ToString());
                                obj_entidad.nroDoc_Operario = Fila["nroDoc_Operario"].ToString();
                                

                                ListLectura.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListLectura;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Cls_Entidad_Resumen_Personal> Capa_Dato_Get_ListadoResumenOperario(string fecha, int id_Operario)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_Resumen_Personal> ListLectura = new List<Cls_Entidad_Resumen_Personal>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Gestion_Ubicacion_Personal_Resumen", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechainI", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@Usuario", SqlDbType.Int).Value = 1;
                        cmd.Parameters.Add("@Cliente", SqlDbType.Int).Value = 1;
                        cmd.Parameters.Add("@Personal", SqlDbType.Int).Value = id_Operario;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_Resumen_Personal obj_entidad = new Cls_Entidad_Resumen_Personal();

                                obj_entidad.Cliente = Fila["Cliente"].ToString();
                                obj_entidad.Operario = Fila["Operario"].ToString();
                                obj_entidad.Asignado = Fila["Asignado"].ToString();
                                obj_entidad.Pendientes = Fila["Pendientes"].ToString();
                                obj_entidad.Ejecutado = Fila["Ejecutado"].ToString();
                                obj_entidad.Efectivos = Fila["Efectivos"].ToString();
                                obj_entidad.NoEfectivos = Fila["NoEfectivos"].ToString();
                                obj_entidad.Avance = Fila["Avance"].ToString();
                                obj_entidad.HoraInicio = Fila["HoraInicio"].ToString();
                                obj_entidad.HoraFinal = Fila["HoraFinal"].ToString();
                                obj_entidad.HorasGeneral = Fila["HorasGeneral"].ToString();
                                obj_entidad.Promedio = Fila["Promedio"].ToString();
                                ListLectura.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListLectura;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Cls_Entidad_RecepcionLect> Capa_Dato_Get_ListaOperario()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_RecepcionLect> ListOper = new List<Cls_Entidad_RecepcionLect>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_RECEPCION_LECTURAS_OPERARIOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_RecepcionLect obj_entidad = new Cls_Entidad_RecepcionLect();

                                obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.flag_limpiar_historial = Convert.ToInt32(Fila["flag_limpiar_historial"].ToString());
                                obj_entidad.DeviceId = Fila["IMAE"].ToString();
                                obj_entidad.VersionApp = Fila["VersionApp"].ToString();
                                ListOper.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListOper;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }





        public List<Cls_Entidad_Lecturas_Trabajo> Capa_Dato_Get_ListaTrabjosResumen(string fecha, int id_servicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_Lecturas_Trabajo> ListOper = new List<Cls_Entidad_Lecturas_Trabajo>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Seguimiento_Lecturas_Resumen", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Trabajo obj_entidad = new Cls_Entidad_Lecturas_Trabajo();

                                obj_entidad.Asignado = (Fila["Asignado"].ToString());
                                obj_entidad.Ejecutado = Fila["Ejecutado"].ToString();
                                obj_entidad.PendDia = Fila["PendDia"].ToString();
                                obj_entidad.Porcentaje = Fila["Porcentaje"].ToString();
                                ListOper.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListOper;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_Lecturas_Trabajo> Capa_Dato_Get_ListaTrabjosDetalle(string fecha, int id_servicio)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_Lecturas_Trabajo> ListOper = new List<Cls_Entidad_Lecturas_Trabajo>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("DSIGE_Seguimiento_Lecturas_Detalle", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@Fecha", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Trabajo obj_entidad = new Cls_Entidad_Lecturas_Trabajo();
                                obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                                obj_entidad.apellidos_operario = (Fila["apellidos_operario"].ToString());
                                obj_entidad.Asignado = (Fila["Asignado"].ToString());
                                obj_entidad.Ejecutado = Fila["Ejecutado"].ToString();
                                obj_entidad.PendDia = Fila["PendDia"].ToString();
                                obj_entidad.Porcentaje = Fila["Porcentaje"].ToString();
                                obj_entidad.Latitud = (Fila["Latitud"].ToString());
                                obj_entidad.Longitud = (Fila["Longitud"].ToString());
                                obj_entidad.Minimo = (Fila["Minimo"].ToString());
                                obj_entidad.Maximo = (Fila["Maximo"].ToString());
                                obj_entidad.totalHora = (Fila["totalHora"].ToString());
                                obj_entidad.nroDoc_Operario = (Fila["nroDoc_Operario"].ToString());

                                ListOper.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListOper;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<Cls_Entidad_RecepcionLect> Capa_Dato_Get_ListandoEstados()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Cls_Entidad_RecepcionLect> ListEstado = new List<Cls_Entidad_RecepcionLect>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_RECEPCION_LECTURAS_ESTADOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_RecepcionLect obj_entidad = new Cls_Entidad_RecepcionLect();

                                obj_entidad.id_Estado = Convert.ToInt32(Fila["id_Estado"].ToString());
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                ListEstado.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListEstado;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<Cls_Entidad_RecepcionLect> Capa_Dato_Get_ListarInformacionLecturas(int local, string fechaAsigna, int servicio, int id_operador, string rt, string s1, string sector, string zona, string suministro, string medidor, int estado)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Cls_Entidad_RecepcionLect> ListData = new List<Cls_Entidad_RecepcionLect>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_RECEPCION_LECTURAS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operador", SqlDbType.Int).Value = id_operador;
                        cmd.Parameters.Add("@rt", SqlDbType.VarChar).Value = rt;
                        cmd.Parameters.Add("@s1", SqlDbType.VarChar).Value = s1;
                        cmd.Parameters.Add("@sector", SqlDbType.VarChar).Value = sector;
                        cmd.Parameters.Add("@zona", SqlDbType.VarChar).Value = zona;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            int cor = 0;
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_RecepcionLect obj_entidad = new Cls_Entidad_RecepcionLect();

                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                obj_entidad.nombreCliente_lectura = Fila["nombreCliente_lectura"].ToString();
                                obj_entidad.operario_lectura = Fila["operario_lectura"].ToString();
                                obj_entidad.orden_lectura = Fila["orden_lectura"].ToString();
                                obj_entidad.lectura_Anterior = Fila["lectura_Anterior"].ToString();
                                obj_entidad.LecturaMovil_Lectura = Fila["LecturaMovil_Lectura"].ToString();
                                obj_entidad.id_Observacion_Lectura = Fila["id_Observacion_Lectura"].ToString();
                                obj_entidad.r1_lectura = Fila["r1_lectura"].ToString();
                                obj_entidad.s1_lectura = Fila["s1_lectura"].ToString();
                                obj_entidad.Sector_lectura = Fila["Sector_lectura"].ToString();
                                obj_entidad.Zona_lectura = Fila["Zona_lectura"].ToString();
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                obj_entidad.DNS_lectura = Fila["DNS_lectura"].ToString();
                                obj_entidad.CodUbicacion_lectura = Fila["CodUbicacion_lectura"].ToString();
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

        public List<Cls_Entidad_RecepcionLect> Capa_Dato_Get_ListarInformacionLecturas_new(int local, string fechaAsigna, int servicio, int id_operador, string rt, string s1, string sector, string zona, string suministro, string medidor, int estado, int id_zona)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Cls_Entidad_RecepcionLect> ListData = new List<Cls_Entidad_RecepcionLect>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_RECEPCION_LECTURAS_new", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@local", SqlDbType.Int).Value = local;
                        cmd.Parameters.Add("@fechaAsigna", SqlDbType.VarChar).Value = fechaAsigna;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@id_operador", SqlDbType.Int).Value = id_operador;
                        cmd.Parameters.Add("@rt", SqlDbType.VarChar).Value = rt;
                        cmd.Parameters.Add("@s1", SqlDbType.VarChar).Value = s1;
                        cmd.Parameters.Add("@sector", SqlDbType.VarChar).Value = sector;
                        cmd.Parameters.Add("@zona", SqlDbType.VarChar).Value = zona;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;
                        cmd.Parameters.Add("@id_zona", SqlDbType.Int).Value = id_zona;


                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            int cor = 0;
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_RecepcionLect obj_entidad = new Cls_Entidad_RecepcionLect();

                                obj_entidad.id_Lectura = Fila["id_Lectura"].ToString();
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                obj_entidad.nombreCliente_lectura = Fila["nombreCliente_lectura"].ToString();
                                obj_entidad.operario_lectura = Fila["operario_lectura"].ToString();
                                obj_entidad.orden_lectura = Fila["orden_lectura"].ToString();
                                obj_entidad.lectura_Anterior = Fila["lectura_Anterior"].ToString();
                                obj_entidad.LecturaMovil_Lectura = Fila["LecturaMovil_Lectura"].ToString();
                                obj_entidad.id_Observacion_Lectura = Fila["id_Observacion_Lectura"].ToString();
                                obj_entidad.r1_lectura = Fila["r1_lectura"].ToString();
                                obj_entidad.s1_lectura = Fila["s1_lectura"].ToString();
                                obj_entidad.Sector_lectura = Fila["Sector_lectura"].ToString();
                                obj_entidad.Zona_lectura = Fila["Zona_lectura"].ToString();
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                obj_entidad.DNS_lectura = Fila["DNS_lectura"].ToString();
                                obj_entidad.CodUbicacion_lectura = Fila["CodUbicacion_lectura"].ToString();

                                obj_entidad.usuario_modificacion = Fila["usuario_modificacion"].ToString();
                                obj_entidad.fecha_modificacion = Fila["fecha_modificacion"].ToString();
                                obj_entidad.URL = Fila["URL"].ToString();

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


        public List<Cls_Entidad_RecepcionLect> Capa_Negocio_Get_ListaInformacionLecturas_Bak(string fecha_movil, string fecha_servidor)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<Cls_Entidad_RecepcionLect> ListData = new List<Cls_Entidad_RecepcionLect>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_RECEPCION_LECTURAS_BAK", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_movil", SqlDbType.VarChar).Value = fecha_movil;
                        cmd.Parameters.Add("@fecha_servidor", SqlDbType.VarChar).Value = fecha_servidor;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            int cor = 0;
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_RecepcionLect obj_entidad = new Cls_Entidad_RecepcionLect();

                                obj_entidad.lectura_bak = Fila["lectura"].ToString();
                                obj_entidad.suministro = Fila["suministro"].ToString();
                                obj_entidad.fecha_movil = Fila["fecha_movil"].ToString();
                                obj_entidad.fecha_servidor = Fila["fecha_sql"].ToString();

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


        public string Capa_Dato_Update_Lectura(int id_lectura, string direccion, string lectura_actual, string observacion, string medidor, int id_usuario, string dns, string ubicacion)
        {
            string Resultado = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_RECEPCION_LECTURAS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_lectura", SqlDbType.Int).Value = id_lectura;
                        cmd.Parameters.Add("@direccion", SqlDbType.VarChar).Value = direccion;
                        cmd.Parameters.Add("@lectura_actual", SqlDbType.VarChar).Value = lectura_actual;
                        cmd.Parameters.Add("@observacion", SqlDbType.VarChar).Value = observacion;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.Parameters.Add("@dns", SqlDbType.VarChar).Value = dns;
                        cmd.Parameters.Add("@ubicacion", SqlDbType.VarChar).Value = ubicacion;
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



        public string Capa_Dato_Update_Historial_Operarios(string listIds)
        {
            string Resultado = "";
            string[] listIdsR = listIds.Split('|');

            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                for (int i = 0; i < listIdsR.Length; i++)
                {
                    using (SqlConnection cn = new SqlConnection(cadenaCnx))
                    {
                        cn.Open();

                        using (SqlCommand cmd = new SqlCommand("SP_U_HISTORIAL_OPERARIOS", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = listIdsR[i];
                            cmd.ExecuteNonQuery();
                            Resultado = "OK";
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Resultado = ex.Message;
            }
            return Resultado;
        }

    }
}
