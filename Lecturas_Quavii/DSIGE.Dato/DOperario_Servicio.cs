using Microsoft.Practices.EnterpriseLibrary.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;

namespace DSIGE.Dato
{
    public class DOperario_Servicio
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-21
        /// </summary>
        public DOperario_Servicio()
        {
            DatabaseFactory.SetDatabaseProviderFactory(new DatabaseProviderFactory(), false);
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-21
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int DAsignaServicio(Request_Operario_Servicio_Asignacion oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_servicio_operario_asignacion", oRq.ope_id, oRq.ser_id, oRq.sop_estado, oRq.usu_id, oRq.emp_id, oRq.loc_id));
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
        public int DAsignaServicioL(Request_Operario_Servicio_Asignacion oRq)
        {
            try
            {
                return Convert.ToInt32(DatabaseFactory.CreateDatabase().ExecuteScalar("dsige_local_operario_asignacion", oRq.ope_id, oRq.ser_id, oRq.sop_estado, oRq.usu_id, oRq.emp_id));
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
