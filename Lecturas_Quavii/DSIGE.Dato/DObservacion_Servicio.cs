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
    public class DObservacion_Servicio
    {

        public DObservacion_Servicio()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 23-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Observacion> DLista(int idEmp, int opcion, int obser)
        {
            try
            {
                List<Observacion> lSe = new List<Observacion>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_observacion_select", idEmp, opcion, obser))
                {
                    if (iDr != null)
                    {
                        while (iDr.Read())
                        {
                            lSe.Add(
                                    new Observacion()
                                    {
                                        obs_id = Convert.ToInt32(iDr["id_Observacion"]),
                                        obs_descripcion = Convert.ToString(iDr["descripcion_observacion"])
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
        public List<Servicio> DObservacionServicio(string obser)
        {
            try
            {
                List<Servicio> lSe = new List<Servicio>();

                using (IDataReader iDr = DatabaseFactory.CreateDatabase().ExecuteReader("dsige_observacion_servicio_select", obser))
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
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DAsignaServicio(string obs_id, int ser_id, int ser_estado, int usu_id, int emp_id)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_servicio_observacion_asignacion", obs_id, ser_id, ser_estado, usu_id, emp_id));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
