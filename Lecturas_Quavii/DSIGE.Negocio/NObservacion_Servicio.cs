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
    public class NObservacion_Servicio
    {

        public NObservacion_Servicio() { }
        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 23-09-2015
        /// </summary>
        /// <returns></returns>
        public List<Observacion> NLista(int idEmp, int opcion, int obser)
        {
            try
            {
                return new DObservacion_Servicio().DLista(idEmp, opcion, obser);
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
        /// <returns></returns>
        public List<Servicio> NObservacionServicio(string obser)
        {
            try
            {
                return new DObservacion_Servicio().DObservacionServicio(obser);
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
        public int NAsignaServicio(string obs_id, int ser_id, int ser_estado, int usu_id, int emp_id)
        {
            try
            {
                return new DObservacion_Servicio().DAsignaServicio(obs_id, ser_id, ser_estado, usu_id, emp_id);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
