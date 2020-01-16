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
    public class NObservacion
    {
        public NObservacion() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <returns></returns>
        public List<Observacion> NLista()
        {
            try
            {
                return new DObservacion().DLista();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NInserta(Request_CRUD_Observacion oRq)
        {
            try
            {
                return new DObservacion().DInserta(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NActualiza(Request_CRUD_Observacion oRq)
        {
            try
            {
                return new DObservacion().DActualiza(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public int NAnula(Request_CRUD_Observacion oRq)
        {
            try
            {
                return new DObservacion().DAnula(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Observacion NAuditoria(Request_CRUD_Observacion oRq)
        {
            try
            {
                return new DObservacion().DAuditoria(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public Observacion NObjeto(Request_CRUD_Observacion oRq)
        {
            try
            {
                return new DObservacion().DObjeto(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}