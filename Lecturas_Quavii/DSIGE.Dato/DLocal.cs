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
    public class DLocal
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        public DLocal()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Local> DLista(Request_Local_Select oRq)
        {
            try
            {
                List<Local> lSe = new List<Local>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_local_select", oRq.emp_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lSe.Add(
                                    new Local()
                                    {
                                        loc_id = Convert.ToInt32(iDr["loc_id"]),
                                        emp_id = Convert.ToInt32(iDr["emp_id"]),
                                        loc_nombre = Convert.ToString(iDr["loc_nombre"]),
                                        loc_direccion = Convert.ToString(iDr["loc_direccion"]),
                                        loc_orden = Convert.ToInt32(iDr["loc_orden"]),
                                        loc_latitud = Convert.ToString(iDr["loc_latitud"]),
                                        loc_longitud = Convert.ToString(iDr["loc_longitud"]),
                                        loc_estado = Convert.ToInt32(iDr["loc_estado"])
                                    }
                                );
                        }
                    }
                }

                return lSe;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 23-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Local> DLista(Request_Local_Operario oRq)
        {
            try
            {
                List<Local> lOb = new List<Local>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_operario_servicio_select", oRq.ope_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lOb.Add(
                                    new Local()
                                    {
                                        loc_id = Convert.ToInt32(iDr["loc_id"]),
                                        loc_nombre = Convert.ToString(iDr["loc_descripcion"]),
                                        loc_estado = Convert.ToInt32(iDr["loc_estado"])
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
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Local> DLista(Request_CRUD_Local oRq)
        {
            try
            {
                List<Local> lSe = new List<Local>();
                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_local", oRq.loc_id, oRq.emp_id, oRq.loc_nombre, oRq.loc_direccion, oRq.loc_latitud, oRq.loc_longitud, oRq.loc_estado, oRq.loc_usuario, 0))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lSe.Add(
                                    new Local()
                                    {
                                        loc_id = Convert.ToInt32(iDr["loc_id"]),
                                        emp_id = Convert.ToInt32(iDr["emp_id"]),
                                        loc_nombre = Convert.ToString(iDr["loc_nombre"]),
                                        loc_direccion = Convert.ToString(iDr["loc_direccion"]),
                                        loc_latitud = Convert.ToString(iDr["loc_latitud"]),
                                        loc_longitud = Convert.ToString(iDr["loc_longitud"]),
                                        loc_estado = Convert.ToInt32(iDr["loc_estado"])
                                    }
                                );
                        }
                    }
                }

                return lSe;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DInserta(Request_CRUD_Local oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_local", 0, oRq.emp_id, oRq.loc_nombre, oRq.loc_direccion, oRq.loc_latitud, oRq.loc_longitud, oRq.loc_estado, oRq.loc_usuario, 1));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DActualiza(Request_CRUD_Local oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_local", oRq.loc_id, oRq.emp_id, oRq.loc_nombre, oRq.loc_direccion, oRq.loc_latitud, oRq.loc_longitud, oRq.loc_estado, oRq.loc_usuario, 2));
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DAnula(Request_CRUD_Local oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_local", oRq.loc_id, 0, "", "", "", "", oRq.loc_estado, oRq.loc_usuario, 3));
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
        public Local DAuditoria(Request_CRUD_Local oRq)
        {
            try
            {
                Local oBj = new Local();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_local", oRq.loc_id, 0, "", "", "", "", 0, 0, 4))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.crea_id = Convert.ToInt32(oDr["loc_crea"]);
                            oBj.crea_nombre = Convert.ToString(oDr["loc_crea_nombre"]);
                            oBj.crea_fecha = Convert.ToString(oDr["loc_crea_fecha"]);
                            oBj.modifica_id = Convert.ToInt32(oDr["loc_modifica"]);
                            oBj.modifica_nombre = Convert.ToString(oDr["loc_modifica_nombre"]);
                            oBj.modifica_fecha = Convert.ToString(oDr["loc_modifica_fecha"]);
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
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Local DObjeto(Request_CRUD_Local oRq)
        {
            try
            {
                Local oBj = new Local();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_local", oRq.loc_id, 0, "", "", "","", 0, 0, 5))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.loc_id = Convert.ToInt32(oDr["loc_id"]);
                            oBj.emp_id = Convert.ToInt32(oDr["emp_id"]);
                            oBj.loc_nombre = Convert.ToString(oDr["loc_nombre"]);
                            oBj.loc_direccion = Convert.ToString(oDr["loc_direccion"]);
                            oBj.loc_latitud = Convert.ToString(oDr["loc_latitud"]);
                            oBj.loc_longitud = Convert.ToString(oDr["loc_longitud"]);
                            oBj.loc_estado = Convert.ToInt32(oDr["loc_estado"]);
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
