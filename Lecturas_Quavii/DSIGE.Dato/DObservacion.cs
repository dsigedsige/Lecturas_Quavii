using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;

namespace DSIGE.Dato
{
    public class DObservacion
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// Descripcion: Declaracion de procedimientos para la tabla servicio
        /// </summary>
        public DObservacion()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <returns></returns>
        public List<Observacion> DLista()
        {
            try
            {
                List<Observacion> lOb = new List<Observacion>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_observacion", 0, 0, "", "", 0, 0, 0, 0,"0","0"))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Observacion()
                                    {
                                        obs_id = Convert.ToInt32(iDr["obs_id"]),
                                        emp_id = Convert.ToInt32(iDr["emp_id"]),
                                        obs_abreviatura = Convert.ToString(iDr["obs_abreviatura"]),
                                        obs_descripcion = Convert.ToString(iDr["obs_descripcion"]),
                                        gde_id = Convert.ToInt32(iDr["gde_id"]),
                                        gde_descripcion = Convert.ToString(iDr["gde_descripcion"]),
                                        obs_estado = Convert.ToInt32(iDr["obs_estado"]),
                                        obs_pideFoto = Convert.ToString(iDr["PideFoto"]),
                                        obs_noPideFoto = Convert.ToString(iDr["NoPideFoto"])
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
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DInserta(Request_CRUD_Observacion oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_observacion", 0, oRq.emp_id, oRq.obs_abreviatura, oRq.obs_descripcion, oRq.gde_id, oRq.obs_estado, oRq.obs_usuario, 1, oRq.obs_pideFoto, oRq.obs_noPideFoto));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DActualiza(Request_CRUD_Observacion oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_observacion", oRq.obs_id, oRq.emp_id, oRq.obs_abreviatura, oRq.obs_descripcion, oRq.gde_id, oRq.obs_estado, oRq.obs_usuario, 2, oRq.obs_pideFoto, oRq.obs_noPideFoto));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DAnula(Request_CRUD_Observacion oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_observacion", oRq.obs_id, 0, "", "", 0, oRq.obs_estado, oRq.obs_usuario, 3));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Observacion DAuditoria(Request_CRUD_Observacion oRq)
        {
            try
            {
                Observacion oBj = new Observacion();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_observacion", oRq.obs_id, 0, "", "", 0, 0, 0, 4))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.crea_id = Convert.ToInt32(oDr["obs_crea"]);
                            oBj.crea_nombre = Convert.ToString(oDr["obs_crea_nombre"]);
                            oBj.crea_fecha = Convert.ToString(oDr["obs_crea_fecha"]);
                            oBj.modifica_id = Convert.ToInt32(oDr["obs_modifica"]);
                            oBj.modifica_nombre = Convert.ToString(oDr["obs_modifica_nombre"]);
                            oBj.modifica_fecha = Convert.ToString(oDr["obs_modifica_fecha"]);
                        }
                    }
                }

                return oBj;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Observacion DObjeto(Request_CRUD_Observacion oRq)
        {
            try
            {
                Observacion oBj = new Observacion();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_observacion", oRq.obs_id, 0, "", "", 0, 0, 0, 5, oRq.obs_pideFoto, oRq.obs_noPideFoto))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.obs_id = Convert.ToInt32(oDr["obs_id"]);
                            oBj.emp_id = Convert.ToInt32(oDr["emp_id"]);
                            oBj.obs_abreviatura = Convert.ToString(oDr["obs_abreviatura"]);
                            oBj.obs_descripcion = Convert.ToString(oDr["obs_descripcion"]);
                            oBj.gde_id = Convert.ToInt32(oDr["gde_id"]);
                            oBj.obs_estado = Convert.ToInt32(oDr["obs_estado"]);
                            oBj.obs_pideFoto = Convert.ToString(oDr["PideFoto"]);
                            oBj.obs_noPideFoto = Convert.ToString(oDr["NoPideFoto"]);
                        }
                    }
                }

                return oBj;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}