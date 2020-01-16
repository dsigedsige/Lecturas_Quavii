using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using Microsoft.ApplicationBlocks.Data;
using DSIGE;
using DSIGE.Modelo;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;

namespace DSIGE.Dato
{
    public class DLectura
    {

        private SqlTransaction oSqlTransaction = null;
        String oSqlConnIN;


        OleDbConnection cn;

        


        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-03
        /// </summary>
        public DLectura()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }


        private OleDbConnection ConectarExcel(string nomExcel)
        {
            cn = new OleDbConnection();
            cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + nomExcel + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";
            cn.Open();
            return cn;
        }

        public DataTable ListaExcel(string fileLocation)
        {
            string sql = "SELECT *FROM [Hoja1$]";
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(sql, ConectarExcel(fileLocation));
            da.SelectCommand.CommandType = CommandType.Text;
            da.Fill(dt);
            cn.Close();
            return dt;
        }


        public bool InsertarExcel(int idlocal, int idUsu,DataTable dt)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_InsertarExcel", idUsu, idUsu,dt);
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //public bool ImportarArchivoExcel01(List<LeerLecturas> ListLectura)
        //{
        //    try
        //    {
        //        foreach (LeerLecturas item in ListLectura)
        //        {
        //            DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_Importar_ArchivoExcel",
        //            item.lec_importacion_archivo,
        //            item.lec_item,
        //            item.instalacion,
        //            item.Equipo,
        //            item.Aparato,
        //            item.lec_Tipocalle,
        //            item.lec_NombreCalle,
        //            item.lec_AlturaCalle,
        //            item.NumeroEdificio,
        //            item.NumeroDepart,
        //            item.PuntoSumin,
        //            item.VivPrin
                    
        //            );
        //        }
        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}


        public bool ImportarArchivoExcel(List<Lectura> ListLectura)
        {
            try
            {
                foreach (Lectura item in ListLectura)
                {
                    DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_Importar_ArchivoExcel",
                    item.lec_importacion_archivo,
                    item.lec_suministro,
                    item.lec_direccion,
                    item.lec_seccion,
                    item.lec_zona,
                    item.lec_correlativo,
                    item.lec_medidor,
                    item.lec_medidor_marca,
                    item.MesUno,
                    item.MesDos,
                    item.MesTres,
                    item.MesCuatro,
                    item.MesCinco,
                    item.dniOpe,
                    item.bloque,
                    item.lec_orden,
                    item.loc_id,                    
                    item.id_UsuarioExpor                                                         
                    );
                }
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool Guardar_TableTemp_Lecturas(int idUsu, string fechaAsignacion)
        {
            try
            {
                DatabaseFactory.CreateDatabase().ExecuteNonQuery("dsige_Guardar_Temp_Lecturas", idUsu, fechaAsignacion);            
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Lectura> DListarTablaTem(int idemp)
        {
            try
            {
                List<Lectura> oLs = new List<Lectura>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_ListarTableTemp", idemp))
                {
                    
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new Lectura()
                                    {
                                        lec_suministro = Convert.ToString(iDr["suministro_lectura"]),
                                        lec_direccion = Convert.ToString(iDr["direccion_lectura"]),
                                        lec_seccion = Convert.ToInt32(iDr["Seccion_lectura"]),
                                        lec_zona = Convert.ToString(iDr["Zona_lectura"]),
                                        lec_correlativo = Convert.ToInt32(iDr["Correlativo"]),
                                        lec_medidor = Convert.ToString(iDr["medidor_lectura"]),
                                        lec_medidor_marca = Convert.ToString(iDr["marcaMedidor_lectura"]),
                                        MesUno = Convert.ToString(iDr["MesUno"]),
                                        MesDos = Convert.ToString(iDr["MesDos"]),
                                        MesTres = Convert.ToString(iDr["MesTres"]),
                                        MesCuatro = Convert.ToString(iDr["MesCuatro"]),
                                        MesCinco = Convert.ToString(iDr["MesCinco"]),
                                        dniOpe = Convert.ToString(iDr["dni_operario"]),
                                        nombre_ope = Convert.ToString(iDr["operador"]),
                                        bloque = Convert.ToString(iDr["Block_lectura"]),
                                        lec_orden = Convert.ToInt32(iDr["orden_lectura"]),
                                        loc_id = Convert.ToInt32(iDr["loc_id"]),
                                        id_UsuarioExpor = Convert.ToInt32(iDr["id_UsuarioExpor"]),
                                        Promedio = Convert.ToInt32(iDr["Promedio"]),
                                        Minimo = Convert.ToInt32(iDr["Minimo"]),
                                        Maximo = Convert.ToInt32(iDr["Maximo"]),
                                        Lect_Minima = Convert.ToInt32(iDr["Lect_Minima"]),
                                        Lect_Maxima = Convert.ToInt32(iDr["Lect_Maxima"]),
                                        Ultimas_Lect = Convert.ToInt32(iDr["3Ultimas_Lect"]),
                                       
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
        /// Fecha: 2015-07-03
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DBulkInfo(List<Lectura> oRq)
        {
            try
            {
                return new Conexion().ExecuteBulkCopy(
                    "tbl_lectura_bulk",
                    (from __l in oRq
                     select new
                     {
                         archivo = __l.lec_importacion_archivo,
                         numero = __l.lec_numero_instalacion,
                         medidor = __l.lec_medidor,
                         direccion = __l.lec_direccion,
                         distrito = __l.lec_distrito,
                         anterior = __l.lec_anterior,
                         interlocutor = __l.lec_interlocutor,
                         suministro = __l.lec_suministro,
                         tipo_cliente = __l.lec_cliente_tipo,
                         categoria = __l.lec_categoria,
                         zona = __l.lec_zona,
                         medidor_marca = __l.lec_medidor_marca,
                         nombre_cliente = __l.lec_cliente_nombre,
                         operario = __l.ope_id,
                         orden = __l.lec_orden,
                         local = __l.loc_id,
                         servicio = __l.ser_id,
                         empresa = __l.emp_id,
                         bloque = __l.bloque,
                         seccion = __l.lec_seccion,
                         correlativo = __l.lec_correlativo,
                         dni_operario = __l.dniOpe
                     })
                );
            }
            catch (Exception e)
            {
                throw e;
            }
        }



        public List<Cls_Entidad_ReporteDiarioNew> Capa_Dato_Get_ListadoReporteDiario(int cliente, string fechaini, string fechafin, int idServicio)
        {
            SqlConnection cn = null;
            string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
            try
            {
                List<Cls_Entidad_ReporteDiarioNew> obj_Lista = new List<Cls_Entidad_ReporteDiarioNew>();
                using (cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Gestion_Estadistica_Operario_SoloEfectivo_Fecha", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Cliente", SqlDbType.Int).Value = cliente;
                        cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                        cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechafin;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = idServicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Cls_Entidad_ReporteDiarioNew obj_entidad = new Cls_Entidad_ReporteDiarioNew();

                                obj_entidad.id_Operario = Convert.ToInt32(Fila["id_Operario"].ToString());
                                obj_entidad.apellidos_operario = Fila["apellidos_operario"].ToString();
                                obj_entidad.Efectivo = Fila["Efectivo"].ToString();
                                obj_entidad.totalHora = Fila["totalHora"].ToString();
                                obj_entidad.Fecha = Fila["Fecha"].ToString();

                                obj_Lista.Add(obj_entidad);
                            }

                        }

                        cn.Close();
                    }
                }

                return obj_Lista;
            }
            catch (Exception ex)
            {
                cn.Close();
                throw ex;
            }
        }
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-07
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Lectura> DListaValidado(Request_Lectura_Validacion oRq){ 
            try{
                List<Lectura> oLs = new List<Lectura>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_lectura_validacion", oRq.archivo)) {
                    if (iDr != null) {
                        while (iDr.Read()) {
                            oLs.Add(
                                    new Lectura() { 
                                        lec_suministro = Convert.ToString(iDr["suministro_lectura"]),
                                        lec_medidor = Convert.ToString(iDr["medidor_lectura"]),
                                        lec_direccion = Convert.ToString(iDr["direccion_lectura"]),
                                        lec_cliente_nombre = Convert.ToString(iDr["nombrecliente_lectura"]),
                                        lec_importacion_archivo = Convert.ToString(iDr["archivoimportacion_lectura"]),
                                        lec_orden = Convert.ToInt32(iDr["orden_lectura"]),
                                        clase = Convert.ToString(iDr["clase"]),
                                        lec_estado = Convert.ToInt32(iDr["estado"]),
                                        lec_seccion = Convert.ToInt32(iDr["seccion"]),
                                        lec_zona = Convert.ToString(iDr["Zona_lectura"]),
                                        lec_correlativo = Convert.ToInt32(iDr["correlativo"]),
                                        lec_medidor_marca = Convert.ToString(iDr["marcaMedidor_lectura"]),
                                        lec_anterior = Convert.ToString(iDr["lectura_Anterior"]),
                                        dniOpe = Convert.ToString(iDr["dni_ope"]),
                                        bloque = Convert.ToString(iDr["bloque"]),
                                    }
                                );
                        }
                    }
                }

                return oLs;
            }
            catch(Exception e){
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-08
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DInsertaMasivo(Request_Lectura_Inserta_Masivo oRq)
        {
            try
            {//oRq.emp_id, oRq.ser_id,
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_lectura_inserta_masivo", oRq.lec_archivo, oRq.lec_asignacion_fecha, oRq.lec_id, oRq.usu_id));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        ///// <summary>
        ///// Autor: jlucero
        ///// Fecha: 2015-07-09
        ///// </summary>
        ///// <param name="oRq"></param>
        ///// <returns></returns>
        //public int DEliminaMasivo(Request_Elimina_Masivo oRq)
        //{
        //    try
        //    {
        //        return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_elimina_masivo", oRq.archivo, oRq.opcion));
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}



        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaHistorico> DListaLecturaResumen(Request_Lectura_Historico oRq)
        {
            try
            {
                List<LecturaHistorico> lOb = new List<LecturaHistorico>(); 

                string[] Servicios = oRq.lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = Servicios.Length;

                if (total == 1) //UN SERVICIO
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Reporte_Resumen_Lecturas", oRq.local, oRq.f_ini, oRq.f_fin, Servicios[0]))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new LecturaHistorico()
                                        {
                                            id_ope = Convert.ToInt32(iDr["id_Operario"]),
                                            des_ope = Convert.ToString(iDr["Operario"]),
                                            total = Convert.ToDouble(iDr["Total"]),
                                            realizado = Convert.ToDouble(iDr["Realizado"]),
                                            conFoto = Convert.ToDouble(iDr["ConFoto"]),
                                            pendiente = Convert.ToDouble(iDr["Pendiente"]),
                                            avance = Convert.ToDouble(iDr["Avance"]),
                                            f_ini = Convert.ToString(iDr["FecIni"]),
                                            f_fin = Convert.ToString(iDr["FecFin"]),
                                            horas = Convert.ToString(iDr["Horas"])
                                        }
                                    );
                            }
                        }
                    }
                }
                else
                { // TODOS LOS SERVICIOS
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Reporte_Resumen_Lecturas_todos", oRq.local, oRq.f_ini, oRq.f_fin))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new LecturaHistorico()
                                        {
                                            id_ope = Convert.ToInt32(iDr["id_Operario"]),
                                            des_ope = Convert.ToString(iDr["Operario"]),
                                            total = Convert.ToDouble(iDr["Total"]),
                                            realizado = Convert.ToDouble(iDr["Realizado"]),
                                            conFoto = Convert.ToDouble(iDr["ConFoto"]),
                                            pendiente = Convert.ToDouble(iDr["Pendiente"]),
                                            avance = Convert.ToDouble(iDr["Avance"]),
                                            f_ini = Convert.ToString(iDr["FecIni"]),
                                            f_fin = Convert.ToString(iDr["FecFin"]),
                                            horas = Convert.ToString(iDr["Horas"])
                                        }
                                    );
                            }
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

        public List<Fotos> Capa_Dato_Get_MostrarFotos(int idLectura, int idservicio)
        {
            try
            {
                string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<Fotos> ListData = new List<Fotos>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_S_MOSTRAR_FOTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Lectura", SqlDbType.Int).Value = idLectura;
                        cmd.Parameters.Add("@id_servicio", SqlDbType.Int).Value = idservicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Fotos obj_entidad = new Fotos();

                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.foto = ruta + Fila["foto"].ToString();

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

        public List<ResumenLecturasCab> DListaLecturaAgrupado(int id_local, int id_servicio, string fechaini, string fechaFin)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ResumenLecturasCab> ListData = new List<ResumenLecturasCab>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    if (id_servicio == 0)
                    {
                        /// Todos los servicios
                        using (SqlCommand cmd = new SqlCommand("DSIGE_Reporte_Resumen_Lecturas_Cabecera_Todos", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@Local", SqlDbType.Int).Value = id_local;
                            cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                            cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechaFin;
                            cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;

                            DataTable dt_detalle = new DataTable();
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt_detalle);
                                foreach (DataRow iDr in dt_detalle.Rows)
                                {
                                    ResumenLecturasCab obj_entidad = new ResumenLecturasCab();
                                    obj_entidad.id_servicio = Convert.ToInt32(iDr["id_servicio"]);
                                    obj_entidad.id_ope = Convert.ToInt32(iDr["id_Operario"]);
                                    obj_entidad.des_ope = Convert.ToString(iDr["Operario"]);
                                    obj_entidad.total = Convert.ToDouble(iDr["Total"]);
                                    obj_entidad.realizado = Convert.ToDouble(iDr["Realizado"]);
                                    obj_entidad.conFoto = Convert.ToDouble(iDr["ConFoto"]);
                                    obj_entidad.pendiente = Convert.ToDouble(iDr["Pendiente"]);
                                    obj_entidad.avance = Convert.ToDouble(iDr["Avance"]);
                                    obj_entidad.f_ini = Convert.ToDateTime(iDr["FecIni"]).ToString("dd/MM/yyyy hh:mm:ss");
                                    obj_entidad.f_fin = Convert.ToDateTime(iDr["FecFin"]).ToString("dd/MM/yyyy hh:mm:ss");
                                    obj_entidad.horas = Convert.ToString(iDr["Horas"]);
                                    ListData.Add(obj_entidad);
                                }
                            }
                        }

                    }
                    else {
                        using (SqlCommand cmd = new SqlCommand("DSIGE_Reporte_Resumen_Lecturas_Cabecera", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@Local", SqlDbType.Int).Value = id_local;
                            cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                            cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechaFin;
                            cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;

                            DataTable dt_detalle = new DataTable();
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt_detalle);
                                foreach (DataRow iDr in dt_detalle.Rows)
                                {
                                    ResumenLecturasCab obj_entidad = new ResumenLecturasCab();
                                    obj_entidad.id_servicio = Convert.ToInt32(iDr["id_servicio"]);
                                    obj_entidad.id_ope = Convert.ToInt32(iDr["id_Operario"]);
                                    obj_entidad.des_ope = Convert.ToString(iDr["Operario"]);
                                    obj_entidad.total = Convert.ToDouble(iDr["Total"]);
                                    obj_entidad.realizado = Convert.ToDouble(iDr["Realizado"]);
                                    obj_entidad.conFoto = Convert.ToDouble(iDr["ConFoto"]);
                                    obj_entidad.pendiente = Convert.ToDouble(iDr["Pendiente"]);
                                    obj_entidad.avance = Convert.ToDouble(iDr["Avance"]);

                                    obj_entidad.f_ini = Convert.ToDateTime(iDr["FecIni"]).ToString("dd/MM/yyyy hh:mm:ss");
                                    obj_entidad.f_fin = Convert.ToDateTime(iDr["FecFin"]).ToString("dd/MM/yyyy hh:mm:ss");

                                    //obj_entidad.f_ini =  Convert.ToString(iDr["FecIni"]);
                                    //obj_entidad.f_fin = Convert.ToString(iDr["FecFin"]);
                                    obj_entidad.horas = Convert.ToString(iDr["Horas"]);
                                    ListData.Add(obj_entidad);
                                }
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

        public List<ResumenLecturasCab> DListaLecturaAgrupado_new(int id_local, int id_servicio, string fechaini, string fechaFin, int id_supervisor, int id_operario_supervisor)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ResumenLecturasCab> ListData = new List<ResumenLecturasCab>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("NEW_DSIGE_Reporte_Resumen_Lecturas_Cabecera", cn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@Local", SqlDbType.Int).Value = id_local;
                            cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                            cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechaFin;
                            cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                            cmd.Parameters.Add("@id_supervisor", SqlDbType.Int).Value = id_supervisor;
                            cmd.Parameters.Add("@id_operario_supervisor", SqlDbType.Int).Value = id_operario_supervisor;

                        DataTable dt_detalle = new DataTable();
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dt_detalle);
                                foreach (DataRow iDr in dt_detalle.Rows)
                                {
                                    ResumenLecturasCab obj_entidad = new ResumenLecturasCab();

                                    obj_entidad.id_ope = Convert.ToInt32(iDr["id_Operario"]);
                                    obj_entidad.des_ope = Convert.ToString(iDr["Operario"]);
                                    obj_entidad.total = Convert.ToDouble(iDr["Total"]);
                                    obj_entidad.realizado = Convert.ToDouble(iDr["Realizado"]);
                                    obj_entidad.conFoto = Convert.ToDouble(iDr["ConFoto"]);
                                    obj_entidad.pendiente = Convert.ToDouble(iDr["Pendiente"]);
                                    obj_entidad.avance = Convert.ToDouble(iDr["Avance"]);
                                    obj_entidad.f_ini = Convert.ToString(iDr["FecIni"]);
                                    obj_entidad.f_fin = Convert.ToString(iDr["FecFin"]);
                                    obj_entidad.horas = Convert.ToString(iDr["Horas"]);

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


        public List<ResumenLecturas> DListaLecturaDetallada(int id_local, int id_servicio, string fechaini, string fechaFin, int id_operario)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ResumenLecturas> ListDetalles = new List<ResumenLecturas>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Reporte_Resumen_Lecturas_Detallado", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                        cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechaFin;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                ResumenLecturas obj_entidad = new ResumenLecturas();

                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.id_servicio = Convert.ToInt32(Fila["id_servicio"].ToString());
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.LecturaMovil_Lectura = Fila["LecturaMovil_Lectura"].ToString();
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                obj_entidad.abreviatura_observacion = Fila["abreviatura_observacion"].ToString();
                                obj_entidad.foto = Fila["foto"].ToString();
                                obj_entidad.latitud_lectura = Fila["latitud_lectura"].ToString();
                                obj_entidad.longitud_lectura = Fila["longitud_lectura"].ToString();
                                obj_entidad.mapa = Fila["mapa"].ToString();

                                ListDetalles.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListDetalles;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        
        public List<LecturaHistorico> DListaLecturaResumenObs(Request_Lectura_Historico oRq, string lista)
        {
            try
            {
                List<LecturaHistorico> oLs = new List<LecturaHistorico>();
                string[] split = lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = split.Length;
                
                if (total > 1)
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Reporte_Resumen_Oper_Obs_Todos", oRq.local, oRq.f_ini, oRq.f_fin, lista))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                oLs.Add(
                                        new LecturaHistorico()
                                        {
                                            des_ope = Convert.ToString(iDr["Operario"]),
                                            obs = Convert.ToString(iDr["OBS"]),
                                            total = Convert.ToDouble(iDr["Total"]),
                                            sinFoto = Convert.ToInt32(iDr["SinFoto"]),
                                            conFoto = Convert.ToDouble(iDr["ConFoto"])
                                        }
                                    );
                            }
                        }
                    }
                }
                else if (total == 1)
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Reporte_Resumen_Oper_Obs", oRq.local, oRq.f_ini, oRq.f_fin, lista))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                oLs.Add(
                                        new LecturaHistorico()
                                        {
                                            des_ope = Convert.ToString(iDr["Operario"]),
                                            obs = Convert.ToString(iDr["OBS"]),
                                            total = Convert.ToDouble(iDr["Total"]),
                                            sinFoto = Convert.ToInt32(iDr["SinFoto"]),
                                            conFoto = Convert.ToDouble(iDr["ConFoto"])
                                        }
                                    );
                            }
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

        public List<ResumenLecturas> DListaLecturaDetalladaEstado(int id_local, int id_servicio, string fechaini, string fechaFin, int id_operario, int categoriaLectura)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<ResumenLecturas> ListDetalles = new List<ResumenLecturas>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("DSIGE_Reporte_Resumen_Lecturas_Detallado_Estado", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = id_local;
                        cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                        cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechaFin;
                        cmd.Parameters.Add("@Servicio", SqlDbType.Int).Value = id_servicio;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@categoriaLectura", SqlDbType.Int).Value = categoriaLectura;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                ResumenLecturas obj_entidad = new ResumenLecturas();

                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.suministro_lectura = Fila["suministro_lectura"].ToString();
                                obj_entidad.medidor_lectura = Fila["medidor_lectura"].ToString();
                                obj_entidad.id_Operario_Lectura = Fila["id_Operario_Lectura"].ToString();
                                obj_entidad.operario = Fila["operario"].ToString();
                                obj_entidad.direccion_lectura = Fila["direccion_lectura"].ToString();
                                obj_entidad.Zona_lectura = Fila["Zona_lectura"].ToString();
                                obj_entidad.LecturaMovil_Lectura = Fila["LecturaMovil_Lectura"].ToString();
                                obj_entidad.confirmacion_Lectura = Fila["confirmacion_Lectura"].ToString();
                                obj_entidad.descripcion_estado = Fila["descripcion_estado"].ToString();
                                obj_entidad.estado = Convert.ToInt32(Fila["estado"].ToString());
                                obj_entidad.abreviatura_observacion = Fila["abreviatura_observacion"].ToString();
                                obj_entidad.foto = Fila["foto"].ToString();
                                obj_entidad.latitudMovil_lectura = Fila["latitudMovil_lectura"].ToString();
                                obj_entidad.longitudMovil_lectura = Fila["longitudMovil_lectura"].ToString();

                                ListDetalles.Add(obj_entidad);
                            }
                        }
                    }
                }

                return ListDetalles;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




        public List<ResumenLecturas> Capa_Dato_Get_MostrarFotosLecturas(int idLectura)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                List<ResumenLecturas> ListData = new List<ResumenLecturas>();
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
                                ResumenLecturas obj_entidad = new ResumenLecturas();
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

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        /// 

        //modifico el servico
        //public List<LecturaHistorico> DListaLecturaResumenObs(Request_Lectura_Historico oRq, string lista)
        //{
        //    try
        //    {
        //        List<LecturaHistorico> oLs = new List<LecturaHistorico>();

        //        using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Reporte_Resumen_Oper_Obs", oRq.local, oRq.f_ini, oRq.f_fin, lista))
        //        {
        //            if (iDr != null)
        //            {
        //                while (iDr.Read())
        //                {
        //                    oLs.Add(
        //                            new LecturaHistorico()
        //                            {
        //                                des_ope = Convert.ToString(iDr["Operario"]),
        //                                obs = Convert.ToString(iDr["OBS"]),
        //                                total = Convert.ToDouble(iDr["Total"]),
        //                                sinFoto = Convert.ToInt32(iDr["SinFoto"]),
        //                                conFoto = Convert.ToDouble(iDr["ConFoto"])
        //                            }
        //                        );
        //                }
        //            }
        //        }

        //        return oLs;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}

        public List<urlLectura> DListaUrlLectura(string id,string fechaini,string fechafin, int con)
        {
            try
            {
                List<urlLectura> oLs = new List<urlLectura>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Lista_url_foto", id,fechaini,fechafin,con))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new urlLectura()
                                    {
                                        url_lectura = Convert.ToString(iDr["fotourl"])
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
        /// Autor: rcontreras
        /// Fecha: 20-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaTomadas> DListaLecturaTomadaFoto(Request_Lectura_Tomadas oRq, string lista, int id_operario)
        {
            try
            {
                List<LecturaTomadas> oLs = new List<LecturaTomadas>();
                
                string[] split = lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = split.Length;


                if (total > 1)
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Lista_Lecturas_tomadas_Foto_Todos", oRq.local, oRq.f_ini, oRq.f_fin, oRq.suministro, oRq.medidor, lista, id_operario))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                oLs.Add(
                                        new LecturaTomadas()
                                        {
                                            id_lectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            orden_lectura = Convert.ToInt32(iDr["orden_lectura"]),
                                            suministro_lectura = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor_lectura = Convert.ToString(iDr["medidor_lectura"]),
                                            marcaMedidor_lectura = Convert.ToString(iDr["nroInstalacion_lectura"]),
                                            direccion_lectura = Convert.ToString(iDr["direccion_lectura"]),
                                            lectura_movil = Convert.ToString(iDr["LecturaMovil_Lectura"]),
                                            tieneFoto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            operario = Convert.ToString(iDr["Operario"]),
                                            fechaLectura = Convert.ToString(iDr["fechaLecturaMovil_lectura"]),
                                            obs = Convert.ToString(iDr["OBS"]),
                                            ubicacion = Convert.ToString(iDr["Ubicacion"]),
                                            notas = Convert.ToString(iDr["Observacion_lectura"]),
                                            latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),

                                        }
                                    );
                            }
                        }
                    }

                }
                else if (total == 1)
                {
                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("DSIGE_Lista_Lecturas_tomadas_Foto", oRq.local, oRq.f_ini, oRq.f_fin, oRq.suministro, oRq.medidor, lista, id_operario))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                oLs.Add(
                                        new LecturaTomadas()
                                        {
                                            id_lectura = Convert.ToInt32(iDr["id_Lectura"]),
                                            orden_lectura = Convert.ToInt32(iDr["orden_lectura"]),
                                            suministro_lectura = Convert.ToString(iDr["suministro_lectura"]),
                                            medidor_lectura = Convert.ToString(iDr["medidor_lectura"]),
                                            marcaMedidor_lectura = Convert.ToString(iDr["marcaMedidor_lectura"]),
                                            direccion_lectura = Convert.ToString(iDr["direccion_lectura"]),
                                            lectura_movil = Convert.ToString(iDr["LecturaMovil_Lectura"]),
                                            tieneFoto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                            operario = Convert.ToString(iDr["Operario"]),
                                            fechaLectura = Convert.ToString(iDr["fechaLecturaMovil_lectura"]),
                                            obs = Convert.ToString(iDr["OBS"]),
                                            ubicacion = Convert.ToString(iDr["Ubicacion"]),
                                            notas = Convert.ToString(iDr["Observacion_lectura"]),
                                            latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
                                            longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),

                                        }
                                    );
                            }
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
        /// Autor: rcontreras
        /// Fecha: 20-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaTomadas> DListaLecturaHistoricoXSuministro(Request_Lectura_Tomadas oRq)
        {
            try
            {
                List<LecturaTomadas> oLs = new List<LecturaTomadas>();
                string[] Servicios =oRq.lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = Servicios.Length;

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_LECTURAS_CORTES_HISTORICO_II", oRq.local, oRq.f_ini, oRq.f_fin, oRq.suministro, oRq.medidor, Servicios[0]))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new LecturaTomadas()
                                    {
                                        id_lectura = Convert.ToInt32(iDr["id_Lectura"]),
                                        orden_lectura = Convert.ToInt32(iDr["orden_lectura"]),
                                        suministro_lectura = Convert.ToString(iDr["suministro_lectura"]),
                                        medidor_lectura = Convert.ToString(iDr["medidor_lectura"]),
                                        marcaMedidor_lectura = Convert.ToString(iDr["marcaMedidor_lectura"]),
                                        direccion_lectura = Convert.ToString(iDr["direccion_lectura"]),
                                        lectura_movil = Convert.ToString(iDr["LecturaMovil_Lectura"]),
                                        tieneFoto_lectura = Convert.ToString(iDr["tieneFoto_lectura"]),
                                        operario = Convert.ToString(iDr["Operario"]),
                                        fechaAsignacion_Lectura = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                        fechaLectura = Convert.ToString(iDr["fechaLecturaMovil_lectura"]),
                                        obs = Convert.ToString(iDr["OBS"]),
                                        notas = Convert.ToString(iDr["Observacion_lectura"]),
                                        id_TipoServicio = Convert.ToInt32(iDr["id_TipoServicio"]),
                                        fotoActa = Convert.ToString(iDr["fotoActa"]),
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
        /// Autor: rcontreras
        /// Fecha: 29-11-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<LecturaTomadas> DListaFotosMasivo()
        {
            try
            {
                List<LecturaTomadas> oLs = new List<LecturaTomadas>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_descarga_foto_masivo"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            oLs.Add(
                                    new LecturaTomadas()
                                    {
                                        id_lectura = Convert.ToInt32(iDr["id_Lectura"]),
                                        nombre_foto = Convert.ToString(iDr["fotourl"]),
                                        suministro_lectura = Convert.ToString(iDr["suministro_lectura"]),
                                        medidor_lectura = Convert.ToString(iDr["medidor_lectura"])
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


        public List<Lecturas_Tomadas> DListaLecturasTomadas(Request_Lectura_Tomadas oRq)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;


                string[] Servicios = oRq.lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
                int total = Servicios.Length;

                List<Lecturas_Tomadas> ListData = new List<Lecturas_Tomadas>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_LISTA_LECTURAS_TOMADAS_NEW", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@Local", SqlDbType.Int).Value = oRq.local;
                        cmd.Parameters.Add("@Tipo", SqlDbType.Int).Value = Servicios[0];
                        cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = oRq.f_ini;
                        cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = oRq.f_fin;

                        cmd.Parameters.Add("@Suministro", SqlDbType.VarChar).Value = oRq.suministro;
                        cmd.Parameters.Add("@Medidor", SqlDbType.VarChar).Value = oRq.medidor;
                        cmd.Parameters.Add("@idoperario", SqlDbType.Int).Value = oRq.operario;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Lecturas_Tomadas obj_entidad = new Lecturas_Tomadas();

                                obj_entidad.id_Lectura = Convert.ToInt32(Fila["id_Lectura"].ToString());
                                obj_entidad.Suministro_lectura = Fila["Suministro_lectura"].ToString();
                                obj_entidad.Medidor_lectura = Fila["Medidor_lectura"].ToString();
                                obj_entidad.Confirmacion_lectura = Fila["Confirmacion_lectura"].ToString();

                                obj_entidad.lectura_Anterior = Fila["lectura_Anterior"].ToString();
                                obj_entidad.promedioConsumo_Lectura = Convert.ToDecimal(Fila["promedioConsumo_Lectura"].ToString());
                                obj_entidad.Operario = Fila["Operario"].ToString();

                                obj_entidad.observacion_lectura = Fila["observacion_lectura"].ToString();
                                obj_entidad.descripcion_observacion = Fila["descripcion_observacion"].ToString();
                                obj_entidad.fechaLecturaMovil_Lectura = Fila["fechaLecturaMovil_Lectura"].ToString();
                                obj_entidad.tieneFoto_lectura = Fila["tieneFoto_lectura"].ToString();

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


        public DataTable D_ListandoReparto_Tomadas(int servicio, string tipoRecibo, string cicloFacturacion, int Estado, string fecha_ini, string fecha_fin, string suministro, string medidor, int operario)
        {
            DataTable dt_detalle = new DataTable();
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;
 
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_REPORTE_REPARTOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;
                        cmd.Parameters.Add("@tipoRecibo", SqlDbType.VarChar).Value = tipoRecibo;
                        cmd.Parameters.Add("@cicloFacturacion", SqlDbType.VarChar).Value = cicloFacturacion;
                        cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = Estado;
                        cmd.Parameters.Add("@fecha_ini", SqlDbType.VarChar).Value = fecha_ini;
                        cmd.Parameters.Add("@fecha_fin", SqlDbType.VarChar).Value = fecha_fin;

                        cmd.Parameters.Add("@suministro", SqlDbType.VarChar).Value = suministro;
                        cmd.Parameters.Add("@medidor", SqlDbType.VarChar).Value = medidor;
                        cmd.Parameters.Add("@operario", SqlDbType.Int).Value = operario;


                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                        }
                    }
                }
                return dt_detalle;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




        public List<Lecturas_Tomadas> D_Descarga_ResultadosReclamosExcel(string fechaini, string fechafin, int servicio)
        {
            try
            {
                var cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

                List<Lecturas_Tomadas> ListData = new List<Lecturas_Tomadas>();
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("PROC_S_LECTURAS_TOMADAS_RECLAMOS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@FechaIni", SqlDbType.VarChar).Value = fechaini;
                        cmd.Parameters.Add("@FechaFin", SqlDbType.VarChar).Value = fechafin;
                        cmd.Parameters.Add("@servicio", SqlDbType.Int).Value = servicio;

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);
                            foreach (DataRow Fila in dt_detalle.Rows)
                            {
                                Lecturas_Tomadas obj_entidad = new Lecturas_Tomadas();

                                obj_entidad.MEDIDOR = Fila["MEDIDOR"].ToString();
                                obj_entidad.CTA_CTO = Fila["CTA_CTO"].ToString();
                                obj_entidad.FECHA_PLAN_LECTURA = Fila["FECHA_PLAN_LECTURA"].ToString();
                                obj_entidad.MES = Fila["MES"].ToString();
                                obj_entidad.LECTURA = Fila["LECTURA"].ToString();

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

        //public List<LecturaTomadas> DListaLecturasTomadas(Request_Lectura_Tomadas oRq)
        //{
        //    try
        //    {
        //        List<LecturaTomadas> oLs = new List<LecturaTomadas>();
        //        string[] Servicios = oRq.lista.Split(new Char[] { ' ', ',', '.', ':', '\t' });
        //        int total = Servicios.Length;

        //        using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_LISTA_LECTURAS_TOMADAS", oRq.local, Servicios[0], oRq.f_ini, oRq.f_fin, oRq.suministro, oRq.medidor, oRq.operario))
        //        {
        //            if (iDr != null)
        //            {
        //                while (iDr.Read())
        //                {
        //                    oLs.Add(
        //                            new LecturaTomadas()
        //                            {
        //                                id_lectura = Convert.ToInt32(iDr["id_Lectura"]),
        //                                orden_lectura = string.IsNullOrEmpty(iDr["orden_lectura"].ToString()) ? 0 : Convert.ToInt32(iDr["orden_lectura"]),
        //                                suministro_lectura = Convert.ToString(iDr["suministro_lectura"]),
        //                                medidor_lectura = Convert.ToString(iDr["medidor_lectura"]),
        //                                marcaMedidor_lectura = Convert.ToString(iDr["marcaMedidor_lectura"]),
        //                                direccion_lectura = Convert.ToString(iDr["direccion_lectura"]),
        //                                lectura_movil = Convert.ToString(iDr["LecturaMovil_Lectura"]),
        //                                operario = Convert.ToString(iDr["Operario"]),
        //                                fechaLectura = Convert.ToString(iDr["fechaLecturaMovil_lectura"]),
        //                                obs = Convert.ToString(iDr["OBS"]),
        //                                ubicacion = Convert.ToString(iDr["Ubicacion"]),
        //                                notas = Convert.ToString(iDr["Observacion_lectura"]),
        //                                latitud_lectura = Convert.ToString(iDr["latitud_lectura"]),
        //                                longitud_lectura = Convert.ToString(iDr["longitud_lectura"]),
        //                                fechaData = Convert.ToString(iDr["meses"])
        //                            }
        //                        );
        //                }
        //            }
        //        }

        //        return oLs;
        //    }
        //    catch (Exception e)
        //    {
        //        throw e;
        //    }
        //}




    }
}