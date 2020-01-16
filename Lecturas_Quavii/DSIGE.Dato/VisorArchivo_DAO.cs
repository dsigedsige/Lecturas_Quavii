using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using System.ComponentModel;
using System.Transactions;
using System.Configuration;
namespace DSIGE.Dato
{
    public class VisorArchivo_DAO
    {

        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;


        public string Capa_Dato_GuardarArchivos(int id_empresa, int usuario, string local, string servicio, string tipodoc, string contrato, string fecharegistro, string nombreArchivo, string url)
        {
            string resultado="";
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    ///eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_I_ARCHIVOS", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_empresa", SqlDbType.Int).Value = id_empresa;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = Convert.ToInt32(local);
                        cmd.Parameters.Add("@id_tiposervicio", SqlDbType.Int).Value = Convert.ToInt32(servicio);
                        cmd.Parameters.Add("@id_tipoarchivo", SqlDbType.Int).Value = Convert.ToInt32(tipodoc); ;

                        cmd.Parameters.Add("@cuentacontrato_archivo", SqlDbType.VarChar).Value = contrato;
                        cmd.Parameters.Add("@fecharegistro_archivo", SqlDbType.VarChar).Value = fecharegistro;
                        cmd.Parameters.Add("@nombre_archivo", SqlDbType.VarChar).Value = nombreArchivo;
                        cmd.Parameters.Add("@url_archivo", SqlDbType.VarChar).Value = url;
                        cmd.Parameters.Add("@usuario_creacion", SqlDbType.Int).Value = usuario;

                        cmd.ExecuteNonQuery();
                    }

                }
                resultado = "1|OK";

            }
            catch (Exception e)
            {
                resultado = "0|" + e.Message;
            }
            return resultado;
        }


 
        public List<ImportarArchivoPlano> Capa_Dato_Get_ListaArchivosAlmacenados(string fechaAsignacion)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ImportarArchivoPlano> ListArchivos = new List<ImportarArchivoPlano>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_ARCHIVO_PLANO_VERIFICAR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion_Lectura", SqlDbType.VarChar).Value = fechaAsignacion;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                ImportarArchivoPlano obj_entidad = new ImportarArchivoPlano();
                                obj_entidad.nombre_ArchivoImportado = Fila["archivoImportacion_lectura"].ToString();
                                ListArchivos.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListArchivos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        public List<ImportarArchivo> Capa_Dato_Get_Locales()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ImportarArchivo> ListLocales = new List<ImportarArchivo>();
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
                                ImportarArchivo obj_entidad = new ImportarArchivo();
                                obj_entidad.id_Local = Convert.ToInt32(Fila["id_Local"].ToString());
                                obj_entidad.nombre_local = Fila["nombre_local"].ToString();
                                ListLocales.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListLocales;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<ImportarArchivo> Capa_Dato_Get_Servicio()
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ImportarArchivo> ListServicios = new List<ImportarArchivo>();
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
                                ImportarArchivo obj_entidad = new ImportarArchivo();
                                obj_entidad.id_TipoServicio = Convert.ToInt32(Fila["id_TipoServicio"].ToString());
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                                ListServicios.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListServicios;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<ImportarArchivo> Capa_Dato_Get_ListandoDetallesDocumentos(int id_empresa, string id_local, string id_servicio, string fecha_ini, string fecha_fin, string contrato)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ImportarArchivo> ListArchivos = new List<ImportarArchivo>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_VISOR_ARCHIVOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@empresa",SqlDbType.Int).Value= id_empresa;
                            cmd.Parameters.Add("@local",SqlDbType.Int).Value= Convert.ToInt32(id_local);
                            cmd.Parameters.Add("@servicio",SqlDbType.Int).Value= Convert.ToInt32(id_servicio);
                            cmd.Parameters.Add("@fecha_ini",SqlDbType.VarChar).Value= fecha_ini;
                            cmd.Parameters.Add("@fecha_fin", SqlDbType.VarChar).Value = fecha_fin;
                            cmd.Parameters.Add("@contrato", SqlDbType.VarChar).Value = contrato;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                ImportarArchivo obj_entidad = new ImportarArchivo();
                                obj_entidad.checkeado = false;
                                obj_entidad.id_Archivo = Convert.ToInt32(Fila["id_Archivo"].ToString());
                                obj_entidad.nombre_local = Fila["nombre_local"].ToString();
                                obj_entidad.nombre_tiposervicio = Fila["nombre_tiposervicio"].ToString();
                                obj_entidad.nombre_TipoArchivo = Fila["nombre_TipoArchivo"].ToString();
                                obj_entidad.fechaRegistro_Archivo = Fila["fechaRegistro_Archivo"].ToString();
                                obj_entidad.cuentaContrato_Archivo = Fila["cuentaContrato_Archivo"].ToString();
                                ListArchivos.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListArchivos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public List<ImportarArchivo> Capa_Dato_Get_RutaArchivos(string id_Archivo)
        {
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string rutaServer = ConfigurationManager.AppSettings["servidor-archivos"];

                List<ImportarArchivo> ListArchivos = new List<ImportarArchivo>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_VISOR_ARCHIVOS_RUTA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@codArchivo", SqlDbType.VarChar).Value = id_Archivo;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                ImportarArchivo obj_entidad = new ImportarArchivo();
                                obj_entidad.ruta_archivo = rutaServer + Fila["ruta_archivo"].ToString(); 
                                ListArchivos.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListArchivos;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
