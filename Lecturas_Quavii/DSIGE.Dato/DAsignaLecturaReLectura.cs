using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using Microsoft.Practices.EnterpriseLibrary.Data;

namespace DSIGE.Dato
{
    public class DAsignaLecturaReLectura
    {

        public DAsignaLecturaReLectura() {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 17-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<AsignaLecturaReLectura> DLista(Request_Asigna_Lectura_ReLectura oRq)
        {
            try
            {
                List<AsignaLecturaReLectura> lOb = new List<AsignaLecturaReLectura>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_asignacion_lectura_relectura", oRq.emp_id, oRq.tip_ser_id, oRq.id_local, oRq.suministro, oRq.medidor, oRq.tecnico_asig_id, oRq.estado_asig, oRq.fecha_asig))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new AsignaLecturaReLectura()
                                    {
                                        corre = Convert.ToInt32(iDr["corre"]),
                                        idLectura = Convert.ToInt32(iDr["id_Lectura"]),
                                        suministroLectura = Convert.ToString(iDr["suministro_lectura"]),
                                        medidorLectura = Convert.ToString(iDr["medidor_lectura"]),
                                        dirLectura = Convert.ToString(iDr["direccion_lectura"]),
                                        lecAnte = Convert.ToString(iDr["lectura_Anterior"]),
                                        lecMin = Convert.ToString(iDr["limite_Minimo_lectura"]),
                                        lecMax = Convert.ToString(iDr["limite_Maximo_lectura"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        zonaLectura = Convert.ToString(iDr["Zona_lectura"]),
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
        /// Fecha: 18-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DActualizaEnviaMovil(string idLectura, DateTime fechActu, int usuEdi, int idEmp)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_asigna_lec_relec_update_app_movil", idLectura, fechActu, usuEdi, idEmp));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Estado> DListaEstados(int opcion)
        {
            try
            {
                List<Estado> lOb = new List<Estado>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_estados_asig_lec_relec", opcion))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Estado()
                                    {
                                        idEstado = Convert.ToInt32(iDr["id_Estado"]),
                                        descEstado = Convert.ToString(iDr["descripcion_estado"])
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
        /// Fecha: 19-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<AsignaLecturaReLectura> DListaHistorico(string codSuministro)
        {
            try
            {
                List<AsignaLecturaReLectura> lOb = new List<AsignaLecturaReLectura>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_asigna_lec_relec_historico", codSuministro))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new AsignaLecturaReLectura()
                                    {
                                        medidorLectura = Convert.ToString(iDr["medidor_lectura"]),
                                        fecLectura = Convert.ToString(iDr["fechaLectura"]),
                                        lectura = Convert.ToString(iDr["lectura"]),
                                        ope_nombre = Convert.ToString(iDr["ope_nombre"]),
                                        obsLectura = Convert.ToString(iDr["observacion"]),
                                        Estado = Convert.ToString(iDr["estado"])
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
        /// Fecha: 19-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DActualizaReasignaOperador(string idLectura, int idOperador, DateTime fechActu, int usuEdi, int idEmp, int opcion)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_asigna_lec_relec_update_resig_operador", idLectura, idOperador, fechActu, usuEdi, idEmp, opcion));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
