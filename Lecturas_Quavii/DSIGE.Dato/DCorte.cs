using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using DSIGE;
using DSIGE.Modelo;
using System.ComponentModel;

namespace DSIGE.Dato
{
    public class DCorte
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-17
        /// </summary>

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
            catch (Exception)
            {
                cn.Close();
                throw;
            }
        }



        public DCorte()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-17
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DBulkInfo(List<Corte> oRq)
        {
            try
            {
                return new Conexion().ExecuteBulkCopy(
                    "tbl_corte_bulk",
                    (from __l in oRq
                     select new
                     {
                         emp_id = __l.emp_id,
                         ser_id = __l.ser_id,
                         cor_importacion_fecha = __l.cor_importacion_fecha,
                         cor_importacion_archivo = __l.cor_importacion_archivo,
                         loc_id = __l.loc_id,
                         cor_aviso_clase = __l.cor_aviso_clase,
                         cor_aviso = __l.cor_aviso,
                         cor_aviso_fecha = __l.cor_aviso_fecha,
                         cor_documento_bloqueo = __l.cor_documento_bloqueo,
                         cor_suministro = __l.cor_suministro,
                         cor_interlocutor = __l.cor_interlocutor,
                         cor_clase_cuenta = __l.cor_clase_cuenta,
                         cor_deuda = __l.cor_deuda,
                         cor_cantidad = __l.cor_cantidad,
                         cor_instalacion_numero = __l.cor_instalacion_numero,
                         cor_medidor = __l.cor_medidor,
                         cor_direccion = __l.cor_direccion,
                         cor_distrito = __l.cor_distrito,
                         cor_seccion = __l.cor_seccion,
                         ope_id = __l.ope_id,
                         cor_orden = __l.cor_orden
                     })
                );
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-17
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Corte> DListaValidado(Request_Corte_Validacion oRq)
        {
            try
            {
                List<Corte> oLs = new List<Corte>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_corte_validacion", oRq.archivo))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new Corte()
                                    {
                                        cor_suministro = Convert.ToString(iDr["suministro_corte"]),
                                        cor_medidor = Convert.ToString(iDr["medidor_corte"]),
                                        cor_direccion = Convert.ToString(iDr["direccion_corte"]),
                                        cor_interlocutor = Convert.ToString(iDr["nombreinterlocutor_corte"]),
                                        cor_importacion_archivo = Convert.ToString(iDr["archivoimportacion_corte"]),
                                        cor_orden = Convert.ToInt32(iDr["orden_corte"]),
                                        clase = Convert.ToString(iDr["clase"]),
                                        cor_estado = Convert.ToInt32(iDr["estado"])
                                    }
                                );
                        }
                    }
                }

                return oLs;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-17
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DInsertaMasivo(Request_Corte_Inserta_Masivo oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_corte_inserta_masivo", oRq.cor_archivo, oRq.cor_asignacion_fecha, oRq.cor_id, oRq.usu_id));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool Capa_Dato_ListaExcelcorte(string fileLocation, int usuario, int idlocal, string idfechaAsignacion)
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

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_E_TEMPORALCORTE", con))
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
                        bulkCopy.DestinationTableName = "TEMPORAL_CORTE_QUAVII";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMPORAL_CORTE_QUAVII SET  fechaCreacion= getdate() ,   local ='" + idlocal + "' , usuarioImportador='" + usuario + "', fechaCarga='" + idfechaAsignacion + "'  WHERE usuarioImportador IS NULL   ";

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
            catch (Exception ex)
            {

              throw new Exception(ex.Message);
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
                     

        public object Capa_Dato_ListaExcelcorte_New(string fileLocation, int usuario, int idlocal, string idfechaAsignacion, int tipo_formato)
        {
            object resul = null;
            resul res = new resul();

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

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_E_TEMPORALCORTE", con))
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
                        bulkCopy.DestinationTableName = "TEMPORAL_CORTE_QUAVII";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMPORAL_CORTE_QUAVII SET  fechaCreacion= getdate() ,   local ='" + idlocal + "' , usuarioImportador='" + usuario + "', fechaCarga='" + idfechaAsignacion + "'  , tipo_formato='" + tipo_formato + "' WHERE usuarioImportador IS NULL   ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
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
            finally
            {
                cn.Close();
            }
            return res;
        }


        public object Capa_Dato_ListaExcel_Reconexion_New(string fileLocation, int usuario, int idlocal, string idfechaAsignacion, int tipoFormato)
        {
            object resul = null;
            resul res = new resul();

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

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_E_TEMPORALRECONEXION", con))
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
                        bulkCopy.DestinationTableName = "TEMPORAL_RECONEXION_QUAVII";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMPORAL_RECONEXION_QUAVII SET  fechaCreacionDB= getdate() ,   local ='" + idlocal + "' , usuarioImportador='" + usuario + "', fechaCarga='" + idfechaAsignacion + "' , tipo_formato='" + tipoFormato + "'  WHERE usuarioImportador IS NULL   ";

                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
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
            finally
            {
                cn.Close();
            }
            return res;
        }



        public List<Reparto> Capa_Dato_Listar_Reparto_Reporte(string fechaAsignacion)
        {
            try
            {
                List<Reparto> oCortes = new List<Reparto>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_GET_REPORTE_REPARTO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@FECHA", SqlDbType.VarChar).Value = fechaAsignacion;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Reparto Entidad = new Reparto();
                                Entidad.NRORECIBO_REPARTO = row["NRORECIBO_REPARTO"].ToString();
                                Entidad.EMPRESA = row["EMPRESA"].ToString();
                                Entidad.CLIENTE = row["CLIENTE"].ToString();
                                Entidad.ORDEN = row["ORDEN"].ToString();
                                Entidad.CTA = row["CTA"].ToString();
                                Entidad.FECHA_MAX = row["FECHA_MAX"].ToString();
                                Entidad.DIRECCION = row["DIRECCION"].ToString();
                                Entidad.DISTRITO = row["DISTRITO"].ToString();
                                Entidad.TIPO_REPARTO = row["TIPO_REPARTO"].ToString();
                                Entidad.CICLO = row["CICLO"].ToString();
                                Entidad.UL = row["UL"].ToString();

                                oCortes.Add(Entidad);
                            }
                        }
                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public bool Capa_Dato_Actualizar_Reparto(int id_operario, string unidad_lectura, string fechaAsignatura, int id_operario_cambiar , string tipo)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_UPDATE_OPERARIO_REPARTO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@unidad_lectura", SqlDbType.VarChar).Value = unidad_lectura;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignatura;
                        cmd.Parameters.Add("@id_operario_cambiar", SqlDbType.Int).Value = id_operario_cambiar;
                        cmd.Parameters.Add("@tipo", SqlDbType.VarChar).Value = tipo;
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


        public bool Capa_Dato_Actualizar_GrandesClientes(int id_operario, string distrito, string fechaAsignatura, int id_operario_cambiar)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_U_GRANDES_CLIENTES_ENVIO_MOVIL", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@distrito", SqlDbType.VarChar).Value = distrito;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignatura;
                        cmd.Parameters.Add("@id_operario_cambiar", SqlDbType.Int).Value = id_operario_cambiar;

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


        public string Capa_Dato_Generando_EnvioMovil_Distribucion_Detallado(List<RepartoDetalle> ListaRepartos, string FechaAsigna, string FechaMovil, int id_usuario, string Forma, string Tipo_recibo)
        {
            string Resultado = "";
            try
            {
                PropertyDescriptorCollection properties = System.ComponentModel.TypeDescriptor.GetProperties(typeof(RepartoDetalle));
                DataTable table = new DataTable();

                foreach (PropertyDescriptor prop in properties)
                {
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }

                foreach (RepartoDetalle item in ListaRepartos)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    table.Rows.Add(row);
                }


                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_D_T_DistribuirReparto", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.ExecuteNonQuery();
                    }
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {
                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "T_DistribuirReparto";
                        bulkCopy.WriteToServer(table);

                        //Actualizando campos 
                        string Sql = "";
                               Sql = "UPDATE T_DistribuirReparto SET   FechaAsigna='" + FechaAsigna + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE  id_usuario IS NULL";
                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                        // update en la tabla Lecturas
                        using (SqlCommand cmd = new SqlCommand("SP_I_REPARTO_ENVIO_MOVIL_NEW", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                            cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                            cmd.Parameters.Add("@Forma", SqlDbType.VarChar).Value = Forma;
                            cmd.Parameters.Add("@Tipo_recibo", SqlDbType.VarChar).Value = Tipo_recibo;
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
                

        public string Capa_Negocio_Generando_EnvioMovil_programacion_Detallado(List<CorteReconexionDetalle> ListaCorRec, string FechaAsigna, string FechaMovil, int id_usuario)
        {
            string Resultado = "";
            try
            {
                PropertyDescriptorCollection properties = System.ComponentModel.TypeDescriptor.GetProperties(typeof(CorteReconexionDetalle));
                DataTable table = new DataTable();

                foreach (PropertyDescriptor prop in properties)
                {
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }

                foreach (CorteReconexionDetalle item in ListaCorRec)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    table.Rows.Add(row);
                }


                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_D_T_DistribuirCorteReconexion", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.ExecuteNonQuery();
                    }
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {
                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "T_DistribuirCorteReconexion";
                        bulkCopy.WriteToServer(table);

                        //Actualizando campos 
                        string Sql = "";
                        Sql = "UPDATE T_DistribuirCorteReconexion SET   FechaAsigna='" + FechaAsigna + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE  id_usuario IS NULL";
                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                        // update en la tabla Lecturas
                        using (SqlCommand cmd = new SqlCommand("SP_I_PROGRAMACION_ENVIO_MOVIL_DETALLADO", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
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



        public string Capa_Dato_Validar_Operario(int id_operario)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_VALIDAR_OPERARIO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ID_OPERARIO", SqlDbType.VarChar).Value = id_operario;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            if (dt_detalle.Rows.Count > 0)
                            {
                                return "existe";
                            }
                            else {
                                return "no_existe";
                            }
                        
                        }
                    }
                }
            }
            catch (Exception)
            {

                return "exit";
            }
        }

        

        public string Capa_dato_Generando_Compartir_lecturas(string fecha, string cod_unidad, int id_operario, string tipo)
        {
            cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            string Resultado = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("PROC_S_REPARTO_COMPARTIR", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fecha;
                        cmd.Parameters.Add("@cod_unidad", SqlDbType.VarChar).Value = cod_unidad;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@tipo", SqlDbType.VarChar).Value = tipo;
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
                              

        public List<Reparto> Capa_Dato_Listar_Reparto(string fechaAsignacion, int cod_usuario, string tipo)
        {
            try
            {
                List<Reparto> oCortes = new List<Reparto>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_REPARTO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@FECHA", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@TIPO", SqlDbType.VarChar).Value = tipo;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Reparto Entidad = new Reparto();


                                if (tipo == "I")
                                {
                                    Entidad.total = Convert.ToInt32(row[0].ToString());
                                    Entidad.unidad_lectura = row[1].ToString();
                                    Entidad.nombre_UnidadLectura = row[2].ToString();
                                    Entidad.id_operario_aux = row[3].ToString();
                                    Entidad.id_operario = Convert.ToInt32(row[4].ToString());
                                    Entidad.flag_compartido = row[5].ToString();
                                    Entidad.id_operario_cambiar = row[6].ToString();
                                }
                                else if (tipo == "R")
                                {
                                    Entidad.total = Convert.ToInt32(row[0].ToString());
                                    Entidad.unidad_lectura = row[1].ToString();
                                    Entidad.nombre_UnidadLectura = row[2].ToString();
                                    Entidad.id_operario_aux = row[3].ToString();
                                    Entidad.id_operario = Convert.ToInt32(row[4].ToString());
                                    Entidad.flag_compartido = row[5].ToString();
                                    Entidad.id_operario_cambiar = row[6].ToString();
                                }
                                oCortes.Add(Entidad);
                            }
                        }
                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        public object Capa_Dato_Listar_grandesClientes_agrupado(string fechaAsignacion, int cod_usuario)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_GRANDES_CLIENTES_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return dt_detalle;
        }


        public object Capa_Dato_Listar_Reparto_Detallado(string fechaAsignacion,  string tipo, string cod_unidad, int id_operario, string forma)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
               using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_REPARTO_DETALLADO_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure; 
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@tipo", SqlDbType.VarChar).Value = tipo;
                        cmd.Parameters.Add("@cod_unidad", SqlDbType.VarChar).Value = cod_unidad;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@forma", SqlDbType.VarChar).Value = forma;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return dt_detalle;
        }
                       
        public object Capa_Dato_get_grandesClientes_Detallado(string fechaAsignacion,  string distrito, int id_operario, string forma)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_GRANDES_CLIENTES_DETALLADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fecha_asignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@distrito", SqlDbType.VarChar).Value = distrito;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@forma", SqlDbType.VarChar).Value = forma;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return dt_detalle;
        }


        public string Capa_Dato_Generando_EnvioMovil_grandesClientes_Detallado(List<GrandesClientesDetalle> ListaGrandesClientes, string FechaAsigna, string FechaMovil, int id_usuario, string Forma)
        {
            string Resultado = "";
            try
            {
                PropertyDescriptorCollection properties = System.ComponentModel.TypeDescriptor.GetProperties(typeof(GrandesClientesDetalle));
                DataTable table = new DataTable();

                foreach (PropertyDescriptor prop in properties)
                {
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }

                foreach (GrandesClientesDetalle item in ListaGrandesClientes)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    table.Rows.Add(row);
                }


                using (SqlConnection con = new SqlConnection(cadenaCnx))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_D_T_GrandesClientes", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                        cmd.ExecuteNonQuery();
                    }
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con))
                    {
                        bulkCopy.BatchSize = 500;
                        bulkCopy.NotifyAfter = 1000;
                        bulkCopy.DestinationTableName = "T_GrandesClientes";
                        bulkCopy.WriteToServer(table);

                        //Actualizando campos 
                        string Sql = "";
                        Sql = "UPDATE T_GrandesClientes SET   FechaAsigna='" + FechaAsigna + "' ,  FechaMovil='" + FechaMovil + "'  , id_usuario='" + id_usuario + "' WHERE  id_usuario IS NULL";
                        using (SqlCommand cmd = new SqlCommand(Sql, con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                        }
                        // update en la tabla Lecturas
                        using (SqlCommand cmd = new SqlCommand("SP_U_GRANDES_CLIENTES_ENVIO_MOVIL_DETALLADO", con))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@FechaAsigna", SqlDbType.VarChar).Value = FechaAsigna;
                            cmd.Parameters.Add("@FechaMovil", SqlDbType.VarChar).Value = FechaMovil;
                            cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = id_usuario;
                            cmd.Parameters.Add("@Forma", SqlDbType.VarChar).Value = Forma;
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




        /// <summary>
        /// lista la tabla temporal agrupado para mostrar en la vista
        /// </summary>
        /// <param name="cod_usuario"></param>
        /// <returns></returns>
        public List<CorteTemporalCorte> Capa_Dato_Listar_Temporal_cortes_Agrupado(string fechaAsignacion, int cod_usuario, int idServicio)
        {
            try
            {
                List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_CORTE_TEMPORAL_AGRUPADO_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaAsignacion_Corte", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@idservicio", SqlDbType.Int).Value = idServicio;


                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                CorteTemporalCorte Entidad = new CorteTemporalCorte();
                                Entidad.cor_distrito = row[0].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[1].ToString());
                                Entidad.nombre_ope = row[2].ToString();
                                Entidad.cant_correcto = Convert.ToInt32(row[3].ToString());
                                Entidad.cant_erroneo = Convert.ToInt32(row[4].ToString());

                                oCortes.Add(Entidad);
                            }
                        }


                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        public object Capa_Dato_Listar_Temporal_cortes_Agrupado_New(string fechaAsignacion, int cod_usuario, int idServicio, int idLocal)
        {
            DataTable dt_detalle = new DataTable();
            resul res = new resul();
            object resul = null;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_CORTE_TEMPORAL_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = idServicio;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = idLocal;

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


        public object Capa_Dato_Listar_Temporal_Reconexion_Agrupado_New(string fechaAsignacion, int cod_usuario, int idServicio, int idLocal)
        {
            DataTable dt_detalle = new DataTable();
            resul res = new resul();
            object resul = null;
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_RECONEXION_TEMPORAL_AGRUPADO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_usuario", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaAsignacion", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = idServicio;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = idLocal;

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



        /// <summary>
        /// Inserta  los datos de la tabla temporal al original TABLA TEMPORAL CORTE -TABLA CORTE
        /// </summary>
        /// <param name="fechaAsignacion"></param>
        /// <param name="id_servicio"></param>
        /// <param name="nombre_archivo"></param>
        /// <param name="usuario"></param>
        /// <returns></returns>
        public string Capa_Dato_Guardar_InformacionCorte(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario, int id_local)
        {
            string res = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_ARCHIVO_GUARDAR_CORTES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fechaAsignacion_Corte", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_TipoServicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@archivoImportacion_Corte", SqlDbType.VarChar).Value = nombre_archivo;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_local;

                        cmd.ExecuteNonQuery();
                    }
                }
                res = "1|OK";
            }
            catch (Exception e)
            {
                res = "0|" + e.Message;
            }
            return res;
        }

        /// <summary>
        /// Insertar a la tabla reconexiones
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="usuario"></param>
        /// <param name="idlocal"></param>
        /// <param name="idfechaAsignacion"></param>
        /// <returns></returns>

        public bool Capa_Dato_ListaExcelreconexiones(string fileLocation, int usuario, int idlocal, string idfechaAsignacion)
        {
            DataTable dt = new DataTable();
            string res = "";
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
                    using (SqlCommand cmd = new SqlCommand("SP_I_T_TEMPORAL_RECONEXIONES", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.ExecuteNonQuery();
                    }

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_E_TEMPORALRECONEXION", con))
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
                        bulkCopy.DestinationTableName = "TEMPORAL_RECONEXIONES";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMPORAL_RECONEXIONES SET   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion='" + idfechaAsignacion + "'  WHERE idUsuarioExport IS NULL   ";

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
            catch (Exception ex)
            {
                res = ex.Message;
                return false;
            }
            finally
            {
                cn.Close();
            }
        }

        public bool Capa_Dato_ListaExcelReparto(string fileLocation, int usuario, int idlocal, string idfechaAsignacion)
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
                    using (SqlCommand cmd = new SqlCommand("SP_I_T_TEMPORAL_REPARTO", con))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.ExecuteNonQuery();
                    }

                    //eliminando registros del usuario
                    using (SqlCommand cmd = new SqlCommand("SP_E_TEMPORALREPARTO", con))
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
                        bulkCopy.DestinationTableName = "TEMPORAL_REPARTO";
                        bulkCopy.WriteToServer(dt);

                        //Actualizando campos 

                        string Sql = "UPDATE TEMPORAL_REPARTO SET   loc_id ='" + idlocal + "' , idUsuarioExport='" + usuario + "', fechaAsignacion='" + idfechaAsignacion + "'  WHERE idUsuarioExport IS NULL   ";

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
            catch (Exception ex)
            {

                return false;
            }
            finally
            {
                cn.Close();
            }
        }

        /// <summary>
        /// lista la tabla temporal de reconexionesagrupado para mostrar en la vista
        /// </summary>
        /// <param name="cod_usuario"></param>
        /// <returns></returns>
        public List<CorteTemporalCorte> Capa_Dato_Listar_Temporal_reconexiones_Agrupado(string fechaAsignacion, int cod_usuario, int idServicio)
        {
            try
            {
                List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_RECONEXION_TEMPORAL_AGRUPADO_II", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@USUARIO", SqlDbType.Int).Value = cod_usuario;
                        cmd.Parameters.Add("@fechaAsignacion_Corte", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@idservicio", SqlDbType.Int).Value = idServicio;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                CorteTemporalCorte Entidad = new CorteTemporalCorte();
                                Entidad.cor_distrito = row[0].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[1].ToString());
                                //  Entidad.contador = Convert.ToInt32(row[2].ToString());                                                              
                                Entidad.nombre_ope = row[2].ToString();
                                Entidad.cant_correcto = Convert.ToInt32(row[3].ToString());
                                Entidad.cant_erroneo = Convert.ToInt32(row[4].ToString());
                                oCortes.Add(Entidad);
                            }
                        }
                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public List<Reparto> Capa_Dato_Listar_Temporal_Reparto_Agrupado(string fechaAsignacion, int cod_usuario, int idServicio)
        {
            try
            {
                List<Reparto> oCortes = new List<Reparto>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_REPARTO_TEMPORAL_AGRUPADO_II", cn))
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
                                Reparto Entidad = new Reparto();
                                Entidad.total = Convert.ToInt32(row[0].ToString());
                                Entidad.unidad_lectura = row[1].ToString();
                                Entidad.nombre_UnidadLectura = row[2].ToString();
                                oCortes.Add(Entidad);
                            }
                        }
                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public bool Capa_Dato_Guardar_InformacionReparto(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario, string fechaRecojo, string horaRecojo, int cantidadRecibos, string fechaMaxima, string ciclo)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_TBL_RECONEXIONES_IMPORTAR_EXCELREPARTO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fechaAsignacion_Corte", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_TipoServicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@archivoImportacion_Corte", SqlDbType.VarChar).Value = nombre_archivo;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;

                        cmd.Parameters.Add("@fechaRecojo", SqlDbType.VarChar).Value = fechaRecojo;
                        cmd.Parameters.Add("@horaRecojo", SqlDbType.VarChar).Value = horaRecojo;
                        cmd.Parameters.Add("@cantidadRecibos", SqlDbType.Int).Value = cantidadRecibos;

                        cmd.Parameters.Add("@fechaMaxima", SqlDbType.VarChar).Value = fechaMaxima;
                        cmd.Parameters.Add("@ciclo", SqlDbType.VarChar).Value = ciclo;


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

        public string Capa_Dato_Guardar_InformacionReconexiones(string fechaAsignacion, int id_servicio, string nombre_archivo, int usuario, int id_Local)
        {
            string res = "";
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_I_IMPORTAR_ARCHIVO_GUARDAR_RECONEXIONES", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@fechaAsignacion_Corte", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@id_TipoServicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@archivoImportacion_Corte", SqlDbType.VarChar).Value = nombre_archivo;
                        cmd.Parameters.Add("@usuario", SqlDbType.Int).Value = usuario;
                        cmd.Parameters.Add("@id_local", SqlDbType.Int).Value = id_Local;
                        cmd.ExecuteNonQuery();
                    }
                }
                res = "1|OK";
            }
            catch (Exception e)
            {
                res = "0|" + e.Message;
            }
            return res;
        }
        
        public List<CorteTemporalCorte> Capa_Dato_Listar_TemporalCorteReconexiones(int cod_usuario, int idtecnico, string distrito)
        {
            try
            {
                List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_DETALLE_CORTE_TEMPORAL_II", cn))
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
                                CorteTemporalCorte Entidad = new CorteTemporalCorte();
                                Entidad.cor_interlocutor = row[0].ToString();
                                Entidad.cor_direccion = row[1].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[2].ToString());
                                Entidad.nombre_ope = row[3].ToString();
                                Entidad.unidad_corte = row[4].ToString();
                                oCortes.Add(Entidad);
                            }
                        }


                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }
        
        public List<CorteTemporalCorte> Capa_Dato_Listar_TemporalReconexiones(int cod_usuario, int idtecnico, string distrito)
        {
            try
            {
                List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTAR_DETALLE_RECONEXION_TEMPORAL", cn))
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
                                CorteTemporalCorte Entidad = new CorteTemporalCorte();
                                Entidad.cor_interlocutor = row[0].ToString();
                                Entidad.cor_direccion = row[1].ToString();
                                Entidad.tecnico = Convert.ToInt32(row[2].ToString());
                                Entidad.nombre_ope = row[3].ToString();
                                Entidad.unidad_corte = row[4].ToString();
                                oCortes.Add(Entidad);
                            }
                        }


                    }
                }

                return oCortes;
            }
            catch (Exception e)
            {

                throw e;
            }
        }
                     
        public List<CorteTemporalCorte> Capa_Dato_Registros_IncorrectosCorte(string fechaAsignacion, int usuario, int id_servicio, string distritocorte)
        {
            try
            {

                List<CorteTemporalCorte> Lista_incorrectos = new List<CorteTemporalCorte>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_LISTA_VALIDACIONCORTE", cn))
                    {

                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion_corte", SqlDbType.VarChar).Value = fechaAsignacion;
                        cmd.Parameters.Add("@IDUSUARIO", SqlDbType.Int).Value = usuario;
                        cmd.Parameters.Add("@TIPOSERVICIO", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@DISTRITOCORTE", SqlDbType.VarChar).Value = distritocorte;
                        /// cmd.ExecuteNonQuery();

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                CorteTemporalCorte Entidad = new CorteTemporalCorte();

                                Entidad.cor_suministro = row["suministro_corte"].ToString();
                                Entidad.id_tiposervicio = Convert.ToInt32(row["id_TipoServicio"].ToString());
                                Entidad.cor_distrito = row["distrito_corte"].ToString();
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

        public List<corteSumnistro> DCorteListaSuministro(int id_tiposervicio, string fecha_asignacion, string suministro)
        {
            try
            {

                List<corteSumnistro> oCortee = new List<corteSumnistro>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_L_CORTE_SUMINISTRO", cn))
                    {

                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion_corte", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@TIPOSERVICIO", SqlDbType.Int).Value = id_tiposervicio;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        /// cmd.ExecuteNonQuery();

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                corteSumnistro Entidad = new corteSumnistro();


                                Entidad.cor_suministro = row["nroInstalacion_corte"].ToString();
                                Entidad.cor_medidor = row["medidor_corte"].ToString();
                                Entidad.crea_nombre = row["nombreCliente_corte"].ToString();
                                Entidad.cor_direccion = row["direccion_corte"].ToString();

                                oCortee.Add(Entidad);


                            }
                        }


                    }
                }

                return oCortee;
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public bool Capa_Dato_Cambio_Estado_Corte(int id_tiposervicio, string fecha_asignacion, string suministro)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_CORTE_SUMINISTRO", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@fechaAsignacion_corte", SqlDbType.VarChar).Value = fecha_asignacion;
                        cmd.Parameters.Add("@TIPOSERVICIO", SqlDbType.Int).Value = id_tiposervicio;
                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;

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

    }
}