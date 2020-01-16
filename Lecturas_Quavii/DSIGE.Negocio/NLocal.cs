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
    public class NLocal
    {
        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        public NLocal() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Local> NLista(Request_Local_Select oRq) {
            try
            {
                return new DLocal().DLista(oRq);
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
        public List<Local> NLista(Request_Local_Operario oRq)
        {
            try
            {
                return new DLocal().DLista(oRq);
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
        public List<Local> NLista(Request_CRUD_Local oRq)
        {
            try
            {
                return new DLocal().DLista(oRq);
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
        public int NInserta(Request_CRUD_Local oRq)
        {
            try
            {
                return new DLocal().DInserta(oRq);
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
        public int NAnula(Request_CRUD_Local oRq)
        {
            try
            {
                return new DLocal().DAnula(oRq);
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
        public Local NAuditoria(Request_CRUD_Local oRq)
        {
            try
            {
                return new DLocal().DAuditoria(oRq);
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
        public Local NObjeto(Request_CRUD_Local oRq)
        {
            try
            {
                return new DLocal().DObjeto(oRq);
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
        public int NActualiza(Request_CRUD_Local oRq)
        {
            try
            {
                return new DLocal().DActualiza(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

    }
}
