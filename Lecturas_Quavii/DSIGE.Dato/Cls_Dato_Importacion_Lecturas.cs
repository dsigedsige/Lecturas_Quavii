using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Data.OleDb;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Transactions;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Web;
using System.IO;
using System.Drawing;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;

namespace DSIGE.Dato
{
    public class Cls_Dato_Importacion_Lecturas
    {

        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
        // string cadenaCnx = "Data Source = .; Initial Catalog = Cobra_Toma_Datos; Integrated Security = True";


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
            catch (Exception ex)
            {
                cn.Close();
                throw;
            }
        }

        public bool Capa_Dato_ListaExcel(string fileLocation, int usuario, int idlocal, string idfechaAsignacion)
        {
            DataTable dt = new DataTable();
            try
            {
                string sql = "SELECT *FROM [Importar$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //creando la tabla temporal
                    using (SqlCommand cmd = new SqlCommand("SP_I_T_LECTURA_TEMPORAL", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.ExecuteNonQuery();
                    }

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_TABLA_IMPORTACION", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "T_LECTURA_TEMPORAL";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE T_LECTURA_TEMPORAL SET   loc_id ='" + idlocal + "' , id_UsuarioExpor='" + usuario + "', fecha_asigna='" + idfechaAsignacion + "'  WHERE id_UsuarioExpor IS NULL   ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                return true;
            }
            catch (Exception)
            {
                cn.Close();
                return false;
            }
        }


        public bool Capa_Dato_Eliminar_Tabla_Temporal(int id_usuario)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_D_TABLA_IMPORTACION", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception)
            {

                return false;
            }
        }


        public bool Capa_Dato_Guardar_Informacion(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_TBL_LECTURA_IMPORTAR_EXCEL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fechaAsignacion_Lectura", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_TipoServicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@archivoImportacion_lectura", SqlDbType.VarChar).Value = nombre_archivo;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
            }
            catch (Exception)
            {

                return false;
            }
        }
        

        public List<Servicio> Capa_Dato_ListarPermisoUsuarioServicio(int idusuario)
        {
            try
            {
                List<Servicio> user = new List<Servicio>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_SERVIUSUARIO_V2", idusuario))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            user.Add(
                                    new Servicio()
                                    {
                                        ser_id = Convert.ToInt32(iDr["id_TipoServicio"]),
                                        ser_descripcion = Convert.ToString(iDr["nombre_tiposervicio"])




                                    }
                                );
                        }
                    }
                }

                return user;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public object Capa_Dato_get_marcaMedidor()
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                    using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_MEDIDOR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
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


        public object Capa_Dato_get_buscarCodgioEmr(string codigo, string fechaCarga)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                    using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_BUSCAR_EMR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@codigo", SqlDbType.VarChar).Value = codigo;
                        cmd.Parameters.Add("@fechaCarga", SqlDbType.VarChar).Value = fechaCarga;

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



        public int Capa_Dato_grabarGrandesClienteFile(int Id_GrandeCliente, string CodigoEMR, string nameFile_GrandeClienteFile, string urlNameFile_GrandeClienteFile, int id_marcaMedidor, string fechaCarga, int idUsuario)
        {
            int resultado = 0;
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                    using (SqlCommand cmd = new SqlCommand("SP_I_GRANDES_CLIENTE_FILE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Id_GrandeCliente", SqlDbType.Int).Value = Id_GrandeCliente;
                        cmd.Parameters.Add("@CodigoEMR", SqlDbType.VarChar).Value = CodigoEMR;
                        cmd.Parameters.Add("@nameFile_GrandeClienteFile", SqlDbType.VarChar).Value = nameFile_GrandeClienteFile;

                        cmd.Parameters.Add("@urlNameFile_GrandeClienteFile", SqlDbType.VarChar).Value = urlNameFile_GrandeClienteFile;
                        cmd.Parameters.Add("@id_marcaMedidor", SqlDbType.Int).Value = id_marcaMedidor;
                        cmd.Parameters.Add("@fechaCarga", SqlDbType.VarChar).Value = fechaCarga;
                        cmd.Parameters.Add("@idUsuario", SqlDbType.Int).Value = idUsuario;
                        cmd.Parameters.Add("@id_file", SqlDbType.Int).Direction = ParameterDirection.Output;
                        cmd.ExecuteNonQuery();

                        resultado = Convert.ToInt32(cmd.Parameters["@id_file"].Value.ToString());
 
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return resultado;
        }

         


        public List<Servicio> Capa_Dato_Listado_Servicios()
        {
            try
            {

                List<Servicio> Lista_Servicios = new List<Servicio>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECTURA_IMPORTACION_SERVICIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;


                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Servicio Entidad = new Servicio();

                                Entidad.ser_id = Convert.ToInt32(row[0].ToString());
                                Entidad.ser_descripcion = row[1].ToString();
                                Lista_Servicios.Add(Entidad);
                            }
                        }


                    }
                }

                return Lista_Servicios;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public bool Capa_Dato_ImportarArchivoExcel(Cls_Entidad_Importacion_Lecturas Entidad)
        {
            try
            {

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_EXCEL_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@CLIENTE ", SqlDbType.VarChar).Value = Entidad.CLIENTE;
                        cmd.Parameters.Add("@DIRECCION ", SqlDbType.VarChar).Value = Entidad.DIRECCION;
                        cmd.Parameters.Add("@Sc ", SqlDbType.VarChar).Value = Entidad.Sc;
                        cmd.Parameters.Add("@Zn ", SqlDbType.VarChar).Value = Entidad.Zn;
                        cmd.Parameters.Add("@CRR ", SqlDbType.VarChar).Value = Entidad.CRR;
                        cmd.Parameters.Add("@MEDIDOR ", SqlDbType.VarChar).Value = Entidad.MEDIDOR;
                        cmd.Parameters.Add("@MARCA ", SqlDbType.VarChar).Value = Entidad.MARCA;
                        cmd.Parameters.Add("@MesUno ", SqlDbType.VarChar).Value = Entidad.MesUno;
                        cmd.Parameters.Add("@MesDos ", SqlDbType.VarChar).Value = Entidad.MesDos;
                        cmd.Parameters.Add("@MesTres ", SqlDbType.VarChar).Value = Entidad.MesTres;
                        cmd.Parameters.Add("@MesCuatro ", SqlDbType.VarChar).Value = Entidad.MesCuatro;
                        cmd.Parameters.Add("@MesCinco ", SqlDbType.VarChar).Value = Entidad.MesCinco;
                        cmd.Parameters.Add("@DNI ", SqlDbType.VarChar).Value = Entidad.DNI;
                        cmd.Parameters.Add("@Nombre_Operador ", SqlDbType.VarChar).Value = Entidad.Nombre_Operador;
                        cmd.Parameters.Add("@Bloque ", SqlDbType.VarChar).Value = Entidad.Bloque;

                        cmd.Parameters.Add("@Orden ", SqlDbType.Int).Value = Entidad.Orden;
                        cmd.Parameters.Add("@loc_id ", SqlDbType.Int).Value = Entidad.loc_id;
                        cmd.Parameters.Add("@id_UsuarioExpor ", SqlDbType.Int).Value = Entidad.id_UsuarioExpor;
                        cmd.Parameters.Add("@Promedio ", SqlDbType.Int).Value = Entidad.Promedio;
                        cmd.Parameters.Add("@Minimo ", SqlDbType.Int).Value = Entidad.Minimo;
                        cmd.Parameters.Add("@Maximo ", SqlDbType.Int).Value = Entidad.Maximo;
                        cmd.Parameters.Add("@Lect_Minima ", SqlDbType.Int).Value = Entidad.Lect_Minima;
                        cmd.Parameters.Add("@Lect_Maxima ", SqlDbType.Int).Value = Entidad.Lect_Maxima;
                        cmd.Parameters.Add("@Ultimas_Lect_3 ", SqlDbType.Int).Value = Entidad.Ultimas_Lect_3;

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


        public List<Cls_Entidad_Importacion_Lecturas> Capa_Dato_Listar_Temporal_Lecturas(int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Importacion_Lecturas> obj_list_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_LECTURA_TEMPORAL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Importacion_Lecturas Entidad = new Cls_Entidad_Importacion_Lecturas();

                                Entidad.CLIENTE = row[0].ToString();
                                Entidad.DIRECCION = row[1].ToString();
                                Entidad.Sc = row[2].ToString();
                                Entidad.Zn = row[3].ToString();
                                Entidad.CRR = row[4].ToString();
                                Entidad.MEDIDOR = row[5].ToString();
                                Entidad.MARCA = row[6].ToString();
                                Entidad.MesUno = row[7].ToString();
                                Entidad.MesDos = row[8].ToString();
                                Entidad.MesTres = row[9].ToString();
                                Entidad.MesCuatro = row[10].ToString();
                                Entidad.MesCinco = row[11].ToString();
                                Entidad.DNI = row[12].ToString();
                                Entidad.Nombre_Operador = row[13].ToString();
                                Entidad.Bloque = row[14].ToString();
                                Entidad.Orden = Convert.ToInt32(row[15].ToString());

                                Entidad.loc_id = Convert.ToInt32(row[16].ToString());
                                Entidad.id_UsuarioExpor = Convert.ToInt32(row[17].ToString());

                                Entidad.Promedio = Convert.ToInt32(row[18].ToString());
                                Entidad.Minimo = Convert.ToInt32(row[19].ToString());
                                Entidad.Maximo = Convert.ToInt32(row[20].ToString());
                                Entidad.Lect_Minima = Convert.ToInt32(row[21].ToString());
                                Entidad.Lect_Maxima = Convert.ToInt32(row[22].ToString());
                                Entidad.Ultimas_Lect_3 = Convert.ToInt32(row[23].ToString());
                                Entidad.FLAG_LECTURA = row[24].ToString();

                                obj_list_Lecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return obj_list_Lecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }



        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Dato_LISTA_TEMPORAL_LECTURA(int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> obj_list_Lecturas = new List<Cls_Entidad_Lecturas_Relecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("sp_S_LISTA_TEMPORAL_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Relecturas Entidad = new Cls_Entidad_Lecturas_Relecturas();

                                Entidad.item = row[0].ToString();
                                Entidad.instalacion = row[1].ToString();
                                Entidad.equipo = row[2].ToString();
                                Entidad.aparato = row[3].ToString();
                                Entidad.lecturaAnterior = row[20].ToString();
                                Entidad.fechaAsignacion = Convert.ToDateTime(row[40].ToString());
                                obj_list_Lecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return obj_list_Lecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
            finally
            {
                cn.Close();
            }
        }





        public List<Cls_Entidad_Importacion_Lecturas> Capa_Dato_Listar_Temporal_Lecturas_Agrupado(int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Importacion_Lecturas> obj_list_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_LECTURA_TEMPORAL_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Importacion_Lecturas Entidad = new Cls_Entidad_Importacion_Lecturas();
                                Entidad.Sc = row[0].ToString();
                                Entidad.DNI = row[1].ToString();
                                Entidad.Nombre_Operador = row[2].ToString();
                                Entidad.cant_correcto = Convert.ToInt32(row[3]);
                                Entidad.cant_erroneo = Convert.ToInt32(row[4]);
                                obj_list_Lecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return obj_list_Lecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }



        public List<Cls_Entidad_Importacion_Lecturas> Capa_Dato_Listar_Temporal_Lecturas_Detallado(int cod_usuario, string SC, string DNI)
        {
            try
            {
                List<Cls_Entidad_Importacion_Lecturas> obj_list_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_LECTURA_TEMPORAL_DETALLADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@SC", SqlDbType.NVarChar).Value = SC;
                        cmd.Parameters.Add("@DNI", SqlDbType.NVarChar).Value = DNI;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Importacion_Lecturas Entidad = new Cls_Entidad_Importacion_Lecturas();

                                Entidad.CLIENTE = row[0].ToString();
                                Entidad.DIRECCION = row[1].ToString();
                                Entidad.Sc = row[2].ToString();
                                Entidad.Zn = row[3].ToString();
                                Entidad.CRR = row[4].ToString();
                                Entidad.MEDIDOR = row[5].ToString();
                                Entidad.MARCA = row[6].ToString();
                                Entidad.MesUno = row[7].ToString();
                                Entidad.MesDos = row[8].ToString();
                                Entidad.MesTres = row[9].ToString();
                                Entidad.MesCuatro = row[10].ToString();
                                Entidad.MesCinco = row[11].ToString();
                                Entidad.DNI = row[12].ToString();
                                Entidad.Nombre_Operador = row[13].ToString();
                                Entidad.Bloque = row[14].ToString();
                                Entidad.Orden = Convert.ToInt32(row[15].ToString());

                                Entidad.loc_id = Convert.ToInt32(row[16].ToString());
                                Entidad.id_UsuarioExpor = Convert.ToInt32(row[17].ToString());

                                Entidad.Promedio = Convert.ToInt32(row[18].ToString());
                                Entidad.Minimo = Convert.ToInt32(row[19].ToString());
                                Entidad.Maximo = Convert.ToInt32(row[20].ToString());
                                Entidad.Lect_Minima = Convert.ToInt32(row[21].ToString());
                                Entidad.Lect_Maxima = Convert.ToInt32(row[22].ToString());
                                Entidad.Ultimas_Lect_3 = Convert.ToInt32(row[23].ToString());
                                Entidad.FLAG_LECTURA = row[24].ToString();

                                obj_list_Lecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return obj_list_Lecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }




        //ACTUALIZANDO FECHA DE LECTURA


        public bool Capa_Dato_ListaExcel_Actualizar_fecha(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion)
        {
            DataTable dt = new DataTable();
            try
            {
                string sql = "SELECT * FROM [Lectura_Actualizar_Fecha$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //creando la tabla temporal
                    using (SqlCommand cmd = new SqlCommand("SP_I_T_LECTURA_TEMPORAL_ACTUALIZAR_FECHA", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.ExecuteNonQuery();
                    }

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_TABLA_IMPORTACION_ACTUALIZAR_FECHA", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "T_LECTURA_TEMPORAL_ACTUALIZAR_FECHA";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE T_LECTURA_TEMPORAL_ACTUALIZAR_FECHA SET   LOCAL_ID ='" + idlocal + "', SERVICIO_ID='" + idservicio + "' , USER_ID='" + usuario + "', FCH_ASIGNA='" + idfechaAsignacion + "'  WHERE USER_ID IS NULL   ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                return true;
            }
            catch (Exception)
            {
                cn.Close();
                return false;
            }
        }



        public List<Cls_Entidad_Importacion_Lecturas> Capa_Datos_Listar_Temporal_Actualizar_Fecha(int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Importacion_Lecturas> obj_list_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_LECTURA_TEMPORAL_ACTUALIZAR_FECHA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Importacion_Lecturas Entidad = new Cls_Entidad_Importacion_Lecturas();

                                Entidad.reg_importado = Convert.ToInt32(row[0]);
                                Entidad.reg_correctos = Convert.ToInt32(row[1]);
                                Entidad.reg_incorrecto = Convert.ToInt32(row[2]);
                                Entidad.diferencia = Convert.ToInt32(row[3]);

                                obj_list_Lecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return obj_list_Lecturas;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public bool Capa_Dato_Eliminar_Tabla_Temporal_Actualizar_Fecha(int id_usuario)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_D_TABLA_IMPORTACION_ACTUALIZAR_FECHA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
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



        public List<Cls_Entidad_Importacion_Lecturas> Capa_Dato_Registros_Incorrectos(int idlocal, int idservicio, int iduser)
        {
            try
            {

                List<Cls_Entidad_Importacion_Lecturas> Lista_incorrectos = new List<Cls_Entidad_Importacion_Lecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LECT_TEM_INCORRECTOS_ACTUALIZAR_FECHA", cn))
                    {

                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@LOCAL_ID", SqlDbType.Int).Value = idlocal;
                        cmd.Parameters.Add("@SERVICIO_ID", SqlDbType.Int).Value = idservicio;
                        cmd.Parameters.Add("@USER_ID", SqlDbType.Int).Value = iduser;
                        /// cmd.ExecuteNonQuery();

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Importacion_Lecturas Entidad = new Cls_Entidad_Importacion_Lecturas();

                                Entidad.SUMINISTRO = row["SUMINISTRO"].ToString();
                                Entidad.MEDIDOR = row["MEDIDOR"].ToString();
                                Entidad.CAMBIO_FECHA = row["CAMBIO_FECHA"].ToString();
                                Lista_incorrectos.Add(Entidad);
                            }
                        }


                    }
                }

                return Lista_incorrectos;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool Capa_Dato_Guardando_Informacion_Actualizacion_Fecha_Lectura(int idusuario)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_ACTUALIZAR_FECHA_LECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USER_ID", SqlDbType.Int).Value = idusuario;
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


        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Dato_Listar_TemporalLectura_Relectura(int cod_usuario, int idtecnico, string  distrito)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> oCls_Entidad_Lecturas_Relecturas = new List<Cls_Entidad_Lecturas_Relecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_DETALLE_LECTURARELECTURA_TEMPORAL_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@TECNICO", SqlDbType.Int).Value = idtecnico;
                        cmd.Parameters.Add("@DISTRITO", SqlDbType.VarChar).Value = distrito;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Relecturas Entidad = new Cls_Entidad_Lecturas_Relecturas();
                                Entidad.interlocutor = row[0].ToString();
                                Entidad.direccion_lectura = row[1].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[2].ToString());
                                Entidad.nombre_ope = row[3].ToString();
                                Entidad.unidadLectura = row[4].ToString();
                                oCls_Entidad_Lecturas_Relecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return oCls_Entidad_Lecturas_Relecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        /// <summary>
        /// insert del temporal de lecturas y relecturas
        /// </summary>
        /// <param name="fechaAsignacion"></param>
        /// <param name="id_servicio"></param>
        /// <param name="nombre_archivo"></param>
        /// <param name="usuario"></param>
        /// <returns></returns>

        public string Capa_Dato_MigracionTemporalLectura(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario)
        {
            var res = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_T_ISERTARTEMPORALLECT_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fechaAsignacion_Lectura", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_TipoServicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@archivoImportacion_lectura", SqlDbType.VarChar).Value = nombre_archivo;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }
                }
                res = "1|OK";
            }
            catch (Exception ex)
            {
                res = "0|" + ex.Message;
            }
            return res;
        }



        /// <summary>
        /// lista la tabla temporal agrupado para mostrar en la vista
        /// </summary>
        /// <param name="cod_usuario"></param>
        /// <returns></returns>
        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Dato_Agrupado_temporal_lectura(string fechaAsignacion, int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> oCls_Entidad_Lecturas_Relecturas = new List<Cls_Entidad_Lecturas_Relecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_T_LISTARTEMPORALLECTURA_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaasignasionlectura", SqlDbType.VarChar).Value = fechaAsignacion;
                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            int total = dt_detalle.Rows.Count;
                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Relecturas Entidad = new Cls_Entidad_Lecturas_Relecturas();
                                Entidad.cor_distrito = row[0].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[1].ToString());
                                Entidad.nombre_ope = row[2].ToString();
                                Entidad.cant_correcto = Convert.ToInt32(row[3].ToString()); 
                                Entidad.cant_erroneo = Convert.ToInt32(row[4].ToString()); 
                                oCls_Entidad_Lecturas_Relecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return oCls_Entidad_Lecturas_Relecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        public object Capa_Dato_Agrupado_lectura(string fechaAsignacion, int cod_usuario)
        {
            DataTable dt_detalle = new DataTable();
            resul res = new resul();
            object resul = null;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                   {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_EXCEL_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fecha_asignar", SqlDbType.VarChar).Value = fechaAsignacion;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            resul = dt_detalle;
                        }
                    }
                }
                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            return res;
        }


        public object Capa_Dato_Agrupado_grandesClientes(string fechaAsignacion, int cod_usuario)
        {
            DataTable dt_detalle = new DataTable();
            resul res = new resul();
            object resul = null;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_EXCEL_GRANDES_CLIENTES_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fecha_asignar", SqlDbType.VarChar).Value = fechaAsignacion;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            resul = dt_detalle;
                        }
                    }
                }
                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            return res;
        }


        public object Capa_Dato_Agrupado_lectura_reclamos(string fechaAsignacion, int cod_usuario)
        {
            DataTable dt_detalle = new DataTable();
            resul res = new resul();
            object resul = null;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_EXCEL_AGRUPADO_RECLAMO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fecha_asignar", SqlDbType.VarChar).Value = fechaAsignacion;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            resul = dt_detalle;
                        }
                    }
                }
                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            return res;
        }

        public object Capa_Dato_Agrupado_reclamos(string fechaAsignacion, int cod_usuario)
        {
            DataTable dt_detalle = new DataTable();
            resul res = new resul();
            object resul = null;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_EXCEL_RECLAMO_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fecha_asignar", SqlDbType.VarChar).Value = fechaAsignacion;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            resul = dt_detalle;
                        }
                    }
                }
                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            return res;
        }


        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Dato_Agrupado_temporal_Relectura(string fechaAsignacion, int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> oCls_Entidad_Lecturas_Relecturas = new List<Cls_Entidad_Lecturas_Relecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_T_LISTARTEMPORALRELECTURA_III", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaasignasionlectura", SqlDbType.VarChar).Value = fechaAsignacion;
                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            int total = dt_detalle.Rows.Count;
                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Relecturas Entidad = new Cls_Entidad_Lecturas_Relecturas();
                                Entidad.cor_distrito = row[0].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[1].ToString());
                                Entidad.nombre_ope = row[2].ToString();
                                Entidad.cant_correcto = Convert.ToInt32(row[3].ToString()); 
                                Entidad.cant_erroneo = Convert.ToInt32(row[4].ToString()); 

                                oCls_Entidad_Lecturas_Relecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return oCls_Entidad_Lecturas_Relecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }
        


        /// <summary>
        /// lista la tabla temporal agrupado para mostrar en la vista
        /// </summary>sdffffffffffffffffffffff
        /// <param name="cod_usuario"></param>
        /// <returns></returns>
        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Dato_Agrupado_temporal_lectura_Repetidas(string fechaAsignacion, int cod_usuario)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> oCls_Entidad_Lecturas_Relecturas = new List<Cls_Entidad_Lecturas_Relecturas>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_T_LISTARTEMPORALLECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaasignasionlectura", SqlDbType.VarChar).Value = fechaAsignacion;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            int total = dt_detalle.Rows.Count;
                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Relecturas Entidad = new Cls_Entidad_Lecturas_Relecturas();
                                Entidad.cor_distrito = row[0].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[1].ToString());
                                Entidad.nombre_ope = row[2].ToString();
                                Entidad.cant_correcto = Convert.ToInt32(row[3].ToString()); ;
                                Entidad.cant_erroneo = Convert.ToInt32(row[4].ToString()); ;





                                oCls_Entidad_Lecturas_Relecturas.Add(Entidad);
                            }
                        }


                    }
                }

                return oCls_Entidad_Lecturas_Relecturas;
            }
            catch (Exception e)
            {

                throw e;
            }
        }



        DataTable TablePrincipal = new DataTable();

        private void GenerarCabeceraDatatable()
        {

            TablePrincipal.Columns.Add("medidorlectura", typeof(string));
            TablePrincipal.Columns.Add("mesUno", typeof(string));
            TablePrincipal.Columns.Add("mesDos", typeof(string));
            TablePrincipal.Columns.Add("mesTres", typeof(string));
            TablePrincipal.Columns.Add("mesCuatro", typeof(string));
            TablePrincipal.Columns.Add("mesCinco", typeof(string));
            TablePrincipal.Columns.Add("consumo1", typeof(string));
            TablePrincipal.Columns.Add("consumo2", typeof(string));
            TablePrincipal.Columns.Add("consumo3", typeof(string));
            TablePrincipal.Columns.Add("consumo4", typeof(string));
            TablePrincipal.Columns.Add("promedioConsumo", typeof(decimal));
            TablePrincipal.Columns.Add("parametroMax", typeof(decimal));
            TablePrincipal.Columns.Add("parametroMin", typeof(decimal));
            TablePrincipal.Columns.Add("lecturaMax", typeof(decimal));
            TablePrincipal.Columns.Add("lecMin", typeof(decimal));
        }

        public string Capa_Dato_InsertandoUpdate_LecturaMeses(int id_user)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {
                    using (SqlConnection con = new SqlConnection(cadenaCnx))
                    {
                        con.Open();
                        //generando tabla
                        GenerarCabeceraDatatable();
                        // generando la estructura

                        using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_LECTURA_TEMPORAL_CALCULO_CONSUMO", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = id_user;
                            cmd.ExecuteNonQuery();
                        }

                        // verificando si hay registros para editar

                        using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_ARCHIVO_EXCEL_DATOS_LECTURA", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@id_user", SqlDbType.Int).Value = id_user;
                            cmd.ExecuteNonQuery();
                        }
                        //haciendo update a las validaciones de temporal_lectura

                        using (SqlCommand cmd = new SqlCommand("SP_U_TEMPORAL_LECTURA_VALIDACION", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.ExecuteNonQuery();
                        }
                        // insertando Nuevo Datos

                        using (SqlCommand cmd = new SqlCommand("SP_S_IMPORTAR_ARCHIVO_EXCEL_DATOS_LECTURA", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@id_user", SqlDbType.Int).Value = id_user;

                            DataTable dt_detalle = new DataTable();
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt_detalle);

                                foreach (DataRow fila in dt_detalle.Rows)
                                {
                                    DataRow row = TablePrincipal.NewRow();
                                    row["medidorlectura"] = fila["aparato"].ToString();
                                    row["mesUno"] = String.IsNullOrEmpty(fila["lecturaAnterior"].ToString()) ? 0 : fila["lecturaAnterior"];
                                    row["mesDos"] = "0";
                                    row["mesTres"] = "0";
                                    row["mesCuatro"] = "0";
                                    row["mesCinco"] = "0";
                                    row["consumo1"] = "0";
                                    row["consumo2"] = "0";
                                    row["consumo3"] = "0";
                                    row["consumo4"] = "0";
                                    row["promedioConsumo"] = (decimal)0;
                                    row["parametroMax"] = (decimal)0;
                                    row["parametroMin"] = (decimal)0;
                                    row["lecturaMax"] = (decimal)0;
                                    row["lecMin"] = (decimal)0;
                                    TablePrincipal.Rows.Add(row);
                                }
                            }

                            if (dt_detalle.Rows.Count > 0)
                            {
                                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                                {
                                    bulkCopy.BatchSize = 500;
                                    bulkCopy.NotifyAfter = 1000;
                                    bulkCopy.DestinationTableName = "tbl_lectura_Meses";
                                    bulkCopy.WriteToServer(TablePrincipal);
                                }
                            }
                        }
                    }

                 
                    Resultado = "1|Ok";
                



                return Resultado;
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }

            return Resultado;
        }

        
        public string Capa_Dato_InsertandoUpdate_Lecturas(string fechaAsignacion, int id_user)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                    {
                        con.Open();

                        using (SqlCommand cmd = new SqlCommand("SP_I_TEMPORAL_LECTURA_INSERTAR_LECTURA_MESES", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = id_user;
                            cmd.ExecuteNonQuery();
                        }
                        using (SqlCommand cmd = new SqlCommand("SP_U_TEMPORAL_LECTURA_VALIDACION", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.ExecuteNonQuery();
                        }
                        // UPDATE A TABLA LECTURA

                        using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_EXCEL_ENVIO_MOVIL", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@fecha_movil", SqlDbType.VarChar).Value = fechaAsignacion;
                            cmd.Parameters.Add("@usuario", SqlDbType.Int ).Value = id_user;
                            cmd.ExecuteNonQuery();
                        }
                    } 
                    Resultado = "1|Ok";
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }
            return Resultado;
        }
               
        public string Capa_Dato_save_Lecturas(string fechaAsignacion,string fechaMovil ,  int id_user)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_ARCHIVO_EXCEL_ENVIAR_MOVIL", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = fechaMovil;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        cmd.ExecuteNonQuery();
                    }
                }
                Resultado = "1|Ok";
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }
            return Resultado;
        }

        public string Capa_Dato_save_grandesClientes(string fechaAsignacion, string fechaMovil, int id_user)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_EXCEL_GRANDES_CLIENTES_MOVIL", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = fechaMovil;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        cmd.ExecuteNonQuery();
                    }
                }
                Resultado = "1|Ok";
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }
            return Resultado;
        }


        public string Capa_Dato_save_Lecturas_Reclamos(string fechaAsignacion, string fechaMovil, int id_user)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_ARCHIVO_ENVIAR_MOVIL_RECLAMOS", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = fechaMovil;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        cmd.ExecuteNonQuery();
                    }
                }
                Resultado = "1|Ok";
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }
            return Resultado;
        }


        public object Capa_Dato_get_generarReparto_Pdf(string fecha_Asigna, int tipo)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                    using (SqlCommand cmd = new SqlCommand("SP_S_GENERACIONREPARTO_PDF", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_Asigna;
                        cmd.Parameters.Add("@tipo", SqlDbType.Int).Value = tipo;

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



        public object Capa_Dato_get_generarReparto_Pdf_individual(string fecha_Asigna, string suministro)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                    using (SqlCommand cmd = new SqlCommand("SP_S_GENERACIONREPARTO_PDF_INDIVIDUAL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_Asigna;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;

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


        public string Capa_Dato_save_Reclamos(string fechaAsignacion, string fechaMovil, int id_user)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_EXCEL_RECLAMO_ENVIAR_MOVIL", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = fechaMovil;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_user;
                        cmd.ExecuteNonQuery();
                    }
                }
                Resultado = "1|Ok";
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }
            return Resultado;
        }


        public string Capa_Dato_ValidacionLectura(string fechaAsignacion)
        {
            var Resultado = "";

            DataTable dt = new DataTable();
            try
            {

                using (TransactionScope scope = new TransactionScope())
                {
                    using (SqlConnection con = new SqlConnection(cadenaCnx))
                    {
                        con.Open();
              

                        using (SqlCommand cmd = new SqlCommand("SP_U_TEMPORAL_LECTURA_VALIDACION", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.ExecuteNonQuery();
                        }
                        // UPDATE A TABLA LECTURA

                        using (SqlCommand cmd = new SqlCommand("SP_U_LECTURA_VALIDACION_LECTURA", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@FECHAASIGNACION", SqlDbType.VarChar).Value = fechaAsignacion;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    scope.Complete();
                    Resultado = "1|Ok";
                }
                return Resultado;
            }
            catch (Exception e)
            {
                Resultado = "0|" + e.Message;
            }
            return Resultado;
        }

        public bool Capa_Dato_Modificar_Informacion(LecturaMes Entidad)
        {
            try
            {
                bool conn = Convert.ToBoolean(DatabaseFactory.CreateDatabase().ExecuteScalar("sp_U_Lectura_Mes",
                      Entidad.medidor_lectura, Entidad.mesUno_lectura, Entidad.mesDos_lectura,
                      Entidad.mesTres_lectura, Entidad.mesCuatro_lectura, Entidad.mesCinco_lectura
                      , Entidad.consumo1, Entidad.consumo2, Entidad.consumo3, Entidad.consumo4,
                      Entidad.promedioConsumo, Entidad.parametroMaximo, Entidad.parametroMinimo
                    , Entidad.lecturaMaximo, Entidad.LecturaMinimo));
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool Capa_Dato_TemporalLectura(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            DataTable dt = new DataTable();
            try
            {
                string sql = "SELECT * FROM [Importar$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //creando la tabla temporal
                    using (SqlCommand cmd = new SqlCommand("SP_T_TEMPORAL_LECTURA", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.ExecuteNonQuery();
                    }

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_T_TEMPORALLECTURA", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "TEMPORAL_LECTURA";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMPORAL_LECTURA SET nombreArchivo='" + nombreArchivo + "',   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion='" + idfechaAsignacion + "'  WHERE idUsuarioExport IS NULL    ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                return true;
            }
            catch (Exception)
            {
                cn.Close();
                return false;
            }
            finally
            {
                cn.Close();
            }
        }


        public class resul
        {
            public bool ok { get; set; }
            public object data { get; set; }
        }



        public object Capa_Dato_save_Temporal_grandesClientes(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            object resul = null;
            resul res = new resul();

            DataTable dt = new DataTable();

            try
            {
                string sql = "SELECT * FROM [Importar$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_TEMP_GRANDES_CLIENTES", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "TEMP_GRANDES_CLIENTES";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMP_GRANDES_CLIENTES SET nombreArchivo='" + nombreArchivo + "',   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion= '" + idfechaAsignacion + "'   WHERE idUsuarioExport IS NULL    ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    resul = "OK";
                }

                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            finally
            {
                cn.Close();
            }
            return res;
        }


        public object Capa_Dato_save_temporalLectura(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            object resul = null;
            resul res = new resul();

            DataTable dt = new DataTable();

            try
            {
                string sql = "SELECT * FROM [Importar$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_TEMPORAL_LECTURA", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "TEMP_LECTURA";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMP_LECTURA SET nombreArchivo='" + nombreArchivo + "',   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion=getdate()   WHERE idUsuarioExport IS NULL    ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    resul = "OK";
                }

                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            finally
            {
                cn.Close();
            }
            return res;
        }


        public object Capa_Dato_save_temporalLectura_Reclamos(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            object resul = null;
            resul res = new resul();

            DataTable dt = new DataTable();

            try
            {
                string sql = "SELECT * FROM [Importar$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_TEMP_LECTURA_RECLAMO", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "TEMP_LECTURA_RECLAMO";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMP_LECTURA_RECLAMO SET nombreArchivo='" + nombreArchivo + "',   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion=getdate()   WHERE idUsuarioExport IS NULL    ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    resul = "OK";
                }

                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            finally
            {
                cn.Close();
            }
            return res;
        }


        public object Capa_Dato_save_temporalReclamos(string fileLocation, int usuario, int idlocal, int idservicio, string idfechaAsignacion, string nombreArchivo)
        {
            object resul = null;
            resul res = new resul();

            DataTable dt = new DataTable();

            try
            {
                string sql = "SELECT * FROM [Importar$]";

                OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
                da.SelectCommand.CommandType = CommandType.Text;
                da.Fill(dt);
                cn.Close();

                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_D_TEMP_RECLAMOS", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = usuario;
                        cmd.ExecuteNonQuery();
                    }

                    //guardando al informacion de la importacion
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {

                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "TEMP_RECLAMOS";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMP_RECLAMOS SET nombreArchivo='" + nombreArchivo + "',   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion=getdate()   WHERE idUsuarioExport IS NULL    ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    resul = "OK";
                }

                res.ok = true;
                res.data = resul;
            }
            catch (Exception ex)
            {
                res.ok = false;
                res.data = ex.Message;
            }
            finally
            {
                cn.Close();
            }
            return res;
        }

        


        public List<Cls_Entidad_Lecturas_Relecturas> Capa_Dato_Registros_IncorrectosLectura(string fechaAsignacion, int usuario, int id_servicio)
        {
            try
            {
                List<Cls_Entidad_Lecturas_Relecturas> Lista_incorrectos = new List<Cls_Entidad_Lecturas_Relecturas>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_LISTA_VALIDACIONLECTURA", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion_Lectura", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@IDUSUARIO", SqlDbType.Int).Value = usuario;
                        cmd.Parameters.Add("@TIPOSERVICIO", SqlDbType.Int).Value = id_servicio;

                        /// cmd.ExecuteNonQuery();

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Cls_Entidad_Lecturas_Relecturas Entidad = new Cls_Entidad_Lecturas_Relecturas();

                                Entidad.suministro_lectura = row["suministro_lectura"].ToString();
                                Entidad.id_tiposervicio = Convert.ToInt32(row["id_TipoServicio"].ToString());
                                Entidad.distritoLectura = row["distrito_lectura"].ToString();
                                Lista_incorrectos.Add(Entidad);
                            }
                        }
                    }
                }
                return Lista_incorrectos;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
                     
        public string Capa_Dato_set_EnviarCorreo(int servicio, string fecha_Asigna, int usuario, int local)
        {
            string mensaje = "";
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                         using (SqlCommand cmd = new SqlCommand("PROC_S_ENVIAR_CORREO_CORTES", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                            cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_Asigna;
                            cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = local;                  

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
                                    mensaje = GenerarArchivoExcel(dt_detalle,   servicio,   fecha_Asigna,   usuario);
                                }
                            }
                        }
                }
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            return mensaje;
        }
        
        public object Capa_Dato_get_repartoPDF(int servicio, string fecha_Asigna)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                List<DistribuirLecturas_E> ListData = new List<DistribuirLecturas_E>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    /// generando los archivos excel
                    using (SqlCommand cmd = new SqlCommand("SP_S_REPARTO_PDF", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha_Asigna;

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


        public string SendEmail(string correo, string rutaFile, string nombreFile, int servicio)
        {
            string path;
            string pathExist;
            string mensaje = "";
            var body = "";
            try
            {
                ////////string correo = "cesar.languasco@gmail.com";

                   body = "Listado de Cortes Calidda";
                if (servicio == 4)
                {
                    body = "Listado de Reconexiones Calidda";
                }

                var message = new MailMessage();

                foreach (var curr_address in correo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    message.To.Add(new MailAddress(curr_address));
                }
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
                    smtp.Credentials = credential;
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;
                    smtp.EnableSsl = true;
                    smtp.Send(message);
                }
                mensaje = "OK";
            }
            catch (Exception ex)
            {
                mensaje = ex.Message;
            }
            return mensaje;
        }

        public string GenerarArchivoExcel(DataTable dt_detalles, int servicio, string fecha_Asigna, int usuario)
        {

            int id_operario;
            string nombreServicio = "";
            string _ruta = "";
            string Res = "";
            int _fila = 2;
            var correo = "";
            try
            {

                id_operario = 0; 
                List<int> list_operarios = new List<int>();
                List<Archivo_E> listArchivos = new List<Archivo_E>();

                nombreServicio = "_EMAIL_CORTES_";
                if (servicio ==4)
                {
                    nombreServicio = "_EMAIL_RECONEXIONES_";
                }

                ///----obteniendo los Operarios---
                foreach (DataRow row in dt_detalles.Rows)
                {
                    if (id_operario != Convert.ToInt32(row["id_Operario_corte"].ToString()))
                    { 
                        if (String.IsNullOrEmpty(row["email_operario"].ToString()) ==false) {
                            id_operario = Convert.ToInt32(row["id_Operario_corte"].ToString());
                            list_operarios.Add(id_operario);
                        }
                    }
                }

                for (int i = 0; i < list_operarios.Count(); i++)
                {
                    _ruta = HttpContext.Current.Server.MapPath("~/Temp/" + list_operarios[i] + "_" + nombreServicio + usuario +".xlsx");

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

                        for (int j = 1; j <= 10; j++)
                        {
                            oWs.Cells[1, j].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }

                        oWs.Cells[1, 1].Value = "#";
                        oWs.Cells[1, 2].Value = "Cuenta Contrato";
                        oWs.Cells[1, 3].Value = "Nombre Interlocutor";
                        oWs.Cells[1, 4].Value = "Clase Cuenta";

                        oWs.Cells[1, 5].Value = "Instalación";
                        oWs.Cells[1, 6].Value = "Medidor";
                        oWs.Cells[1, 7].Value = "Dirección de Instalación";
                        oWs.Cells[1, 8].Value = "Distrito de Instalación";
                        oWs.Cells[1, 9].Value = "Unidad de Lectura";
                        oWs.Cells[1, 10].Value = "Técnico";

                        int ac = 0;
                         _fila = 2;
                        id_operario = 0;
                        correo = "";
                        foreach (DataRow oBj in dt_detalles.Rows)
                        {
                            id_operario = Convert.ToInt32(oBj["id_Operario_corte"].ToString());
                            if (list_operarios[i] == Convert.ToInt32(oBj["id_Operario_corte"].ToString()))
                            {
                                ac += 1;
                                oWs.Cells[_fila, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                oWs.Cells[_fila, 1].Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center; // alinear texto  
                                oWs.Cells[_fila, 1].Value = ac;

                                correo = oBj["email_operario"].ToString();
                                oWs.Cells[_fila, 2].Value = oBj["suministro_corte"].ToString();
                                oWs.Cells[_fila, 3].Value = oBj["nombreInterlocutor_corte"].ToString();
                                oWs.Cells[_fila, 4].Value = oBj["claseAviso_corte"].ToString();

                                oWs.Cells[_fila, 5].Value = oBj["nroInstalacion_corte"].ToString();
                                oWs.Cells[_fila, 6].Value = oBj["medidor_corte"].ToString();

                                oWs.Cells[_fila, 7].Value = oBj["direccion_corte"].ToString();
                                oWs.Cells[_fila, 8].Value = oBj["distrito_corte"].ToString();
                                oWs.Cells[_fila, 9].Value = oBj["unidad_corte"].ToString();
                                oWs.Cells[_fila, 10].Value = oBj["id_Operario_corte"].ToString();

                                if (id_operario != Convert.ToInt32(oBj["id_Operario_corte"].ToString()))
                                {
                                    id_operario = Convert.ToInt32(oBj["id_Operario_corte"].ToString());
                                }          
                                _fila++;
                            }

                            if (id_operario != Convert.ToInt32(oBj["id_Operario_corte"].ToString()))
                            {
                                break;
                            }
                        }
                        oWs.Row(1).Style.Font.Bold = true;
                        oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                        oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                        oWs.Column(1).Style.Font.Bold = true;

                        for (int k = 1; k <= 9; k++)
                        {
                            oWs.Column(k).AutoFit();
                        }
                        oEx.Save();
                    }

                    if (correo != null || correo != "")
                    {
                        Archivo_E obj_entidad = new Archivo_E();
                        obj_entidad.email = correo;
                        obj_entidad.ruta = _ruta;
                        obj_entidad.nombreFile = list_operarios[i] + "_" + nombreServicio + usuario + ".xlsx";

                        listArchivos.Add(obj_entidad);
                    }

                    //------suspendemos el hilo, y esperamos ..
                    System.Threading.Thread.Sleep(2000);

                }

                //---generando el envio de correo
                for (int i = 0; i < listArchivos.Count(); i++)
                {
                    Res = SendEmail(listArchivos[i].email, listArchivos[i].ruta, listArchivos[i].nombreFile, servicio);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return Res;
        }
    }
}
