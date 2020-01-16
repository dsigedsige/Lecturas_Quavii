using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class NOperario_Servicio
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-21
        /// </summary>
        public NOperario_Servicio() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-21
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NAsignaServicio(Request_Operario_Servicio_Asignacion oRq)
        {
            try
            {
                return new DOperario_Servicio().DAsignaServicio(oRq);
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
        public int NAsignaServicioL(Request_Operario_Servicio_Asignacion oRq)
        {
            try
            {
                return new DOperario_Servicio().DAsignaServicioL(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
