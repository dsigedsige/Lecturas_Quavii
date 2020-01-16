using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DSIGE.Dato
{
    public class DEnvioTrabajoCliLecReLec
    {
        public DEnvioTrabajoCliLecReLec()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 20-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        /// 

        string cadenaCnx = System.Configuration.ConfigurationManager.ConnectionStrings["dataSige"].ConnectionString;

        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Dato_Excel_Envio_Trabajo_cliente(int id_Local,int  id_operario ,string suministro_lectura, string medidor_lectura, string fechaAsignacion_Lectura, int estado)
        {
            try
            {
                List<Edicion_Envio_Trabajo_Cliente_Lectura> Lista_Operario = new List<Edicion_Envio_Trabajo_Cliente_Lectura>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_VisualizarEnvioAlCliente", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Local", SqlDbType.Int).Value = id_Local;
                        cmd.Parameters.Add("@id_operario", SqlDbType.Int).Value = id_operario;
                        cmd.Parameters.Add("@suministro_lectura", SqlDbType.NVarChar).Value = suministro_lectura;
                        cmd.Parameters.Add("@medidor_lectura", SqlDbType.NVarChar).Value = medidor_lectura;
                        cmd.Parameters.Add("@fechaAsignacion_Lectura", SqlDbType.NVarChar).Value = fechaAsignacion_Lectura;
                        cmd.Parameters.Add("@estado", SqlDbType.Int).Value = estado;                       

                        DataTable dt_detalle = new DataTable();
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Edicion_Envio_Trabajo_Cliente_Lectura Entidad = new Edicion_Envio_Trabajo_Cliente_Lectura();

                                Entidad.Seccion_Lectura = row["Seccion_Lectura"].ToString();
                                Entidad.Zona_lectura = row["Zona_lectura"].ToString();
                                Entidad.suministro_lectura = row["suministro_lectura"].ToString();
                                Entidad.lecturaMovil_lectura = Convert.ToString(row["lecturaMovil_lectura"].ToString());
                                Entidad.abreviatura_observacion = row["abreviatura_observacion"].ToString();
                                Entidad.Observacion_lectura = row["Observacion_lectura"].ToString();
                                Entidad.Block_lectura = row["Block_lectura"].ToString();
                                Entidad.fechaAsignacion_Lectura = row["fechaAsignacion_Lectura"].ToString();
                                Entidad.confirmacion_Lectura = row["confirmacion_Lectura"].ToString();
                                Lista_Operario.Add(Entidad);
                            }
                        }


                    }
                }

                return Lista_Operario;
            }
            catch (Exception e)
            {

                throw e;
            }
        }






        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Dato_Listado_Operario()
        {
            try
            {
                List<Edicion_Envio_Trabajo_Cliente_Lectura> Lista_Operario = new List<Edicion_Envio_Trabajo_Cliente_Lectura>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_ENVIAR_TRABAJO_CLIENTE_OPER", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Edicion_Envio_Trabajo_Cliente_Lectura Entidad = new Edicion_Envio_Trabajo_Cliente_Lectura();


                                Entidad.id_Operario_Lectura = row["id_Operario"].ToString();
                                Entidad.nombre_operario = row["nombre_operario"].ToString();

                                Lista_Operario.Add(Entidad);
                            }
                        }


                    }
                }

                return Lista_Operario;
            }
            catch (Exception e)
            {

                throw e;
            }
        }




        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Dato_Listado_Observaciones()
        {
            try
            {
                List<Edicion_Envio_Trabajo_Cliente_Lectura> Lista_Observaciones = new List<Edicion_Envio_Trabajo_Cliente_Lectura>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_ENVIAR_TRABAJO_CLIENTE_OBS", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Edicion_Envio_Trabajo_Cliente_Lectura Entidad = new Edicion_Envio_Trabajo_Cliente_Lectura();

                                Entidad.id_Observacion_Lectura = row["id_Observacion"].ToString();
                                Entidad.descripcion_observacion = row["descripcion_observacion"].ToString();

                                Lista_Observaciones.Add(Entidad);
                            }
                        }


                    }
                }

                return Lista_Observaciones;
            }
            catch (Exception e)
            {

                throw e;
            }
        }


        public List<Edicion_Envio_Trabajo_Cliente_Lectura> Capa_Dato_Listando_Registros_Modo_edicion( int id_lectura )
        {
            try
            {
                List<Edicion_Envio_Trabajo_Cliente_Lectura> Lista_Observaciones = new List<Edicion_Envio_Trabajo_Cliente_Lectura>();

                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();
                    using (SqlCommand cmd = new SqlCommand("SP_S_ENVIAR_TRABAJO_CLIENTE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@id_Lectura", SqlDbType.Int).Value = id_lectura;

                        DataTable dt_detalle = new DataTable();

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dt_detalle);

                            foreach (DataRow row in dt_detalle.Rows)
                            {
                                Edicion_Envio_Trabajo_Cliente_Lectura Entidad = new Edicion_Envio_Trabajo_Cliente_Lectura();

                                Entidad.id_Lectura =Convert.ToInt32(row["id_Lectura"].ToString());
                                Entidad.suministro_lectura = row["suministro_lectura"].ToString();
                                Entidad.medidor_lectura = row["medidor_lectura"].ToString();

                                Entidad.marcaMedidor_lectura = row["marcaMedidor_lectura"].ToString();
                                Entidad.direccion_lectura = row["direccion_lectura"].ToString();

                                Entidad.id_Operario_Lectura = row["id_Operario_Lectura"].ToString();
                                Entidad.Lectura_Manual_Lectura = row["Lectura_Manual_Lectura"].ToString();

                                Entidad.fechaLecturaMovil_lectura = row["fechaLecturaMovil_lectura"].ToString();
                                Entidad.id_Observacion_Lectura = row["id_Observacion_Lectura"].ToString();

                                Lista_Observaciones.Add(Entidad);
                            }
                        }


                    }
                }

                return Lista_Observaciones;
            }
            catch (Exception e)
            {

                throw e;
            }
        }


//@id_Lectura INT,
//@id_Operario_Lectura VARCHAR(10),
//@Lectura_Manual_Lectura VARCHAR(30),
//@fechaLecturaMovil_lectura DATETIME,
//@id_Observacion_Lectura  VARCHAR(10),
//@id_usuario INT



        public bool Capa_Dato_Guardar_Informacion(int id_Lectura, string id_Operario_Lectura, string Lectura_Manual_Lectura, DateTime fechaLecturaMovil_lectura, string id_Observacion_Lectura, int id_usuario)
        {
            try
            {
                using (SqlConnection cn = new SqlConnection(cadenaCnx))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("SP_U_ENVIAR_TRABAJO_CLIENTE", cn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.CommandType = CommandType.StoredProcedure;

                        cmd.Parameters.Add("@id_Lectura", SqlDbType.Int).Value = id_Lectura;
                        cmd.Parameters.Add("@id_Operario_Lectura", SqlDbType.VarChar).Value = id_Operario_Lectura;
                        cmd.Parameters.Add("@Lectura_Manual_Lectura", SqlDbType.VarChar).Value = Lectura_Manual_Lectura;
                        cmd.Parameters.Add("@fechaLecturaMovil_lectura", SqlDbType.DateTime).Value = fechaLecturaMovil_lectura;
                        cmd.Parameters.Add("@id_Observacion_Lectura", SqlDbType.VarChar).Value = id_Observacion_Lectura;
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



        


        public List<EnvioTrabajoCliLecReLec> DLista(Request_Asigna_Lectura_ReLectura oRq)
        {
            try
            {
                List<EnvioTrabajoCliLecReLec> lOb = new List<EnvioTrabajoCliLecReLec>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_envio_trabajos_cliente_lectura_relectura", oRq.emp_id, oRq.tip_ser_id, oRq.id_local, oRq.suministro, oRq.medidor, oRq.tecnico_asig_id, oRq.estado_asig, oRq.fecha_asig))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new EnvioTrabajoCliLecReLec()
                                    {
                                        corre = Convert.ToInt32(iDr["corre"]),
                                        idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                        foto = Convert.ToString(iDr["foto"]),
                                        fecLectura = Convert.ToString(iDr["fechLectura"]),
                                        suministroLectura = Convert.ToString(iDr["suministro_lectura"]),
                                        medidorLectura = Convert.ToString(iDr["medidor_lectura"]),
                                        lecMin = Convert.ToString(iDr["limite_Minimo_lectura"]),
                                        lecMax = Convert.ToString(iDr["limite_Maximo_lectura"]),
                                        lecMovil = Convert.ToString(iDr["lecMovil"]),
                                        consuActu = Convert.ToString(iDr["consuActu"]),
                                        lecProm = Convert.ToString(iDr["lecProm"]),
                                        lecPorcen = Convert.ToString(iDr["lecPorcen"]),
                                        lecVali = Convert.ToString(iDr["lecVali"]),
                                        lecManual = Convert.ToString(iDr["lecManual"]),
                                        idlecObs = Convert.ToString(iDr["idlecObs"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        zonaLectura = Convert.ToString(iDr["Zona_lectura"]),
                                        obsLectura = Convert.ToString(iDr["lecObs"]),
                                        dirLectura = Convert.ToString(iDr["direccion_lectura"]),
                                        fecAsignaLectura = Convert.ToString(iDr["fechaAsignacion_Lectura"]),
                                        clienLectura = Convert.ToString(iDr["nombreCliente_lectura"]),
                                        Estado = Convert.ToString(iDr["abreviatura_estado"]),
                                        nroinstalacion = Convert.ToString(iDr["nroInstalacion_lectura"])
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

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 20-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DActualizaEnviaCliente(string idLectura, int usuEdi, int idEmp)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_envia_trabajo_lec_relec_update_envio_clie", idLectura, usuEdi, idEmp));
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 28-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLec> DListaFotos(int idLectura)
        {
            try
            {
                List<EnvioTrabajoCliLecReLec> lOb = new List<EnvioTrabajoCliLecReLec>();
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_envio_trabajos_cliente_fotos", idLectura))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new EnvioTrabajoCliLecReLec()
                                    {
                                        idLectura = Convert.ToInt32(iDr["fot_id_lectura"]),
                                        foto = ruta + Convert.ToString(iDr["fot_nombre"])
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


        public List<Fotoselfie_reparto> DFotosSelfie_Reparto(string fecha, int id_operario)
        {
            try
            {
                List<Fotoselfie_reparto> lOb = new List<Fotoselfie_reparto>();
                string ruta = ConfigurationManager.AppSettings["servidor-foto-lectura"];

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("PROC_S_REPARTO_FOTOS_SELFIE", fecha, id_operario))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Fotoselfie_reparto()
                                    {
                                        RutaFoto = ruta + Convert.ToString(iDr["RutaFoto"])
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



        public List<EnvioTrabajoCliLecReLec> DListaFotos_Historico(int idLectura, int Tiposervicio)
        {
            try
            {
                List<EnvioTrabajoCliLecReLec> lOb = new List<EnvioTrabajoCliLecReLec>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("SP_S_FOTOS_LECTURA_CORTES_HISTORICO", idLectura, Tiposervicio))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new EnvioTrabajoCliLecReLec()
                                    {
                                        idLectura = Convert.ToInt32(iDr["fot_id_lectura"]),
                                        foto = Convert.ToString(iDr["fot_nombre"])
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

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLec> DListaPreVisualizaEnvio(string codLectura)
        {
            try
            {
                List<EnvioTrabajoCliLecReLec> lOb = new List<EnvioTrabajoCliLecReLec>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_envio_cliente_pre_visualiza_envio", codLectura))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new EnvioTrabajoCliLecReLec()
                                    {
                                        corre = Convert.ToInt32(iDr["corre"]),
                                        //lecSucursal = Convert.ToString(iDr["sucursal"]),
                                        lecSeccionLectura = Convert.ToString(iDr["Seccion_Lectura"]),
                                        zonaLectura = Convert.ToString(iDr["Zona_lectura"]),
                                        lecCorrelativo = Convert.ToString(iDr["lecCorrelativo"]),
                                        suministroLectura = Convert.ToString(iDr["suministro_lectura"]),
                                        medidorLectura = Convert.ToString(iDr["medidor_lectura"]),
                                        lecMovil = Convert.ToString(iDr["LecturaMovil_Lectura"]),
                                        idlecObs = Convert.ToString(iDr["id_Observacion_Lectura"]),
                                        obsLectura = Convert.ToString(iDr["Observacion_lectura"]),
                                        dirLectura = Convert.ToString(iDr["direccion_lectura"]),
                                        lecMarcaMedidor = Convert.ToString(iDr["marcaMedidor_lectura"]),
                                        ope_nombre = Convert.ToString(iDr["lecturista"]),
                                        fecLectura = Convert.ToString(iDr["fechaLecturaMovil_lectura"]),
                                        confirLectura = Convert.ToString(iDr["confirmacion_Lectura"]),
                                        blockLectura = Convert.ToString(iDr["block_lectura"]),
                                        foto = Convert.ToString(iDr["tieneFoto_lectura"])

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


        /// <summary>
        /// Lectura Envio CLiente
        /// </summary>
        /// <param name="codLectura"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLec> DListaPreVisualizaLecturaEnvioCliente(string fechamovil, int tipoServicio, List<int> List_codigos)
        {
            try
            {
                List<EnvioTrabajoCliLecReLec> lOb = new List<EnvioTrabajoCliLecReLec>();

                string lecturas = "";
                string id_cadenaCodigos = "";


                if (tipoServicio == 1 || tipoServicio == 2)
                {
                    foreach (var row in List_codigos)
                    {
                        lecturas = lecturas + row.ToString() + ',';
                    }
                    id_cadenaCodigos = lecturas.Substring(0, lecturas.ToString().Length - 1);

                    using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_LecturaEnvioCliente", id_cadenaCodigos))
                    {
                        if (iDr != null)
                        {
                            while (iDr.Read())
                            {
                                lOb.Add(
                                        new EnvioTrabajoCliLecReLec()
                                        {
                                            corre = Convert.ToInt32(iDr["corre"]),
                                            //lecSucursal = Convert.ToString(iDr["sucursal"]),
                                            //lecSeccionLectura = Convert.ToString(iDr["Seccion_Lectura"]),
                                            zonaLectura = Convert.ToString(iDr["Zona_lectura"]),
                                            //lecCorrelativo = Convert.ToString(iDr["lecCorrelativo"]),
                                            suministroLectura = Convert.ToString(iDr["suministro_lectura"]),
                                            medidorLectura = Convert.ToString(iDr["medidor_lectura"]),
                                            lecMovil = Convert.ToString(iDr["LecturaMovil_Lectura"]),
                                            idlecObs = Convert.ToString(iDr["id_Observacion_Lectura"]),
                                            obsLectura = Convert.ToString(iDr["Observacion_lectura"]),
                                            dirLectura = Convert.ToString(iDr["direccion_lectura"]),
                                            lecMarcaMedidor = Convert.ToString(iDr["marcaMedidor_lectura"]),
                                            ope_nombre = Convert.ToString(iDr["lecturista"]),
                                            fecLectura = Convert.ToString(iDr["fechaLecturaMovil_lectura"]),
                                            confirLectura = Convert.ToString(iDr["confirmacion_Lectura"]),
                                            //blockLectura = Convert.ToString(iDr["block_lectura"]),
                                            foto = Convert.ToString(iDr["tieneFoto_lectura"])

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


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<EnvioTrabajoCliLecReLecLecturaPendiente> DListaLecturasPendientes(DateTime fechaAsigna, int idEmp)
        {
            try
            {
                List<EnvioTrabajoCliLecReLecLecturaPendiente> lOb = new List<EnvioTrabajoCliLecReLecLecturaPendiente>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_lista_operarios_lecturas_pendientes", fechaAsigna, idEmp))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new EnvioTrabajoCliLecReLecLecturaPendiente()
                                    {
                                        idOpeLectura = Convert.ToInt32(iDr["id_Operario_Lectura"]),
                                        nomOpe = Convert.ToString(iDr["ope_nombre"]),
                                        nomEstado = Convert.ToString(iDr["est_nombre"]),
                                        cantidad = Convert.ToInt32(iDr["cantidad"])
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

    }
}
