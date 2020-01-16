using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
    public class NAsignaLecturaReLectura
    {
        public NAsignaLecturaReLectura() { }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 17-09-2015
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<AsignaLecturaReLectura> NLista(Request_Asigna_Lectura_ReLectura oRq)
        {
            try
            {
                return new DAsignaLecturaReLectura().DLista(oRq);
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
        public int NActualizaEnviaMovil(Request_actualiza_app_movil_lec_relec oRq)
        {
            try
            {
                return new DAsignaLecturaReLectura().DActualizaEnviaMovil(oRq.idLectura, oRq.fechActu, oRq.usuEdi, oRq.idEmp);
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
        public List<Estado> NListaEstados(int opcion)
        {
            try
            {
                return new DAsignaLecturaReLectura().DListaEstados(opcion);
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
        public List<AsignaLecturaReLectura> NListaHistorico(string codSuministro)
        {
            try
            {
                return new DAsignaLecturaReLectura().DListaHistorico(codSuministro);
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
        public int NActualizaReasignaOperador(Request_actualiza_app_movil_lec_relec oRq)
        {
            try
            {
                return new DAsignaLecturaReLectura().DActualizaReasignaOperador(oRq.idLectura, oRq.idOperador, oRq.fechActu, oRq.usuEdi, oRq.idEmp, oRq.opcion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
