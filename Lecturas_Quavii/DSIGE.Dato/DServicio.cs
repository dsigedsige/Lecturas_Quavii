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
    public class DServicio
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// Descripcion: Declaracion de procedimientos para la tabla servicio
        /// </summary>
        public DServicio()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <returns></returns>
        public List<Servicio> DLista() {
            try
            {
                List<Servicio> lSe = new List<Servicio>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_servicios", 0, "", 0, 0, 0))
                {
                    if (iDr != null) {
                        while (iDr.Read()) {
                            lSe.Add(
                                    new Servicio() {
                                        ser_id = Convert.ToInt32(iDr["ser_id"]),
                                        ser_descripcion = Convert.ToString(iDr["ser_descripcion"]),
                                        ser_estado = Convert.ToInt32(iDr["ser_estado"])
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
        /// Autor: jlucero
        /// Fecha: 2015-06-19
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Servicio> DLista(Request_Servicio_Operario oRq)
        {
            try
            {
                List<Servicio> lSe = new List<Servicio>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_servicio_select", oRq.ope_id))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lSe.Add(
                                    new Servicio()
                                    {
                                        ser_id = Convert.ToInt32(iDr["ser_id"]),
                                        ser_descripcion = Convert.ToString(iDr["ser_descripcion"]),
                                        ser_estado = Convert.ToInt32(iDr["ser_estado"])
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
        /// Autor: jlucero
        /// Fecha: 2015-06-15
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DInserta(Request_CRUD_Servicio oRq) { 
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_servicios", 0, oRq.descripcion, oRq.estado, oRq.usuario, 1));
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
        public int DActualiza(Request_CRUD_Servicio oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_servicios", oRq.id, oRq.descripcion, oRq.estado, oRq.usuario, 2));
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
        public int DAnula(Request_CRUD_Servicio oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_servicios", oRq.id, "", oRq.estado, oRq.usuario, 3));
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
        public Servicio DAuditoria(Request_CRUD_Servicio oRq)
        {
            try
            {
                Servicio oBj = new Servicio();

                using(IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_servicios", oRq.id, "", 0, 0, 4)){
                    if (oDr != null) {
                        while (oDr.Read()) {
                            oBj.crea_id = Convert.ToInt32(oDr["ser_crea"]);
                            oBj.crea_nombre = Convert.ToString(oDr["ser_crea_nombre"]);
                            oBj.crea_fecha = Convert.ToString(oDr["ser_crea_fecha"]);
                            oBj.modifica_id = Convert.ToInt32(oDr["ser_modifica"]);
                            oBj.modifica_nombre = Convert.ToString(oDr["ser_modifica_nombre"]);
                            oBj.modifica_fecha = Convert.ToString(oDr["ser_modifica_fecha"]);
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
        public Servicio DObjeto(Request_CRUD_Servicio oRq)
        {
            try
            {
                Servicio oBj = new Servicio();

                using (IDataReader oDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_servicios", oRq.id, "", 0, 0, 5))
                {
                    if (oDr != null)
                    {
                        while (oDr.Read())
                        {
                            oBj.ser_id = Convert.ToInt32(oDr["ser_id"]);
                            oBj.ser_descripcion = Convert.ToString(oDr["ser_descripcion"]);
                            oBj.ser_estado = Convert.ToInt32(oDr["ser_estado"]);
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
