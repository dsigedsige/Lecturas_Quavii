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
    public class NOperario
    {
        public NOperario() { }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <returns></returns>
        public List<Operario> NLista() {
            try {
                return new DOperario().DLista();
            }
            catch (Exception e) {
                throw e;
            }
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-19
        /// </summary>
        /// <param name="oRq"></param>
        /// <returns></returns>
        public List<Operario> NLista(Request_Operario_Empresa_Local_Opcion oRq)
        {
            try
            {
                return new DOperario().DLista(oRq);
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
        public int NInserta(Request_CRUD_Operario oRq)
        {
            try
            {
                return new DOperario().DInserta(oRq);
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
        public int NActualiza(Request_CRUD_Operario oRq)
        {
            try
            {
                return new DOperario().DActualiza(oRq);
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
        public int NAnula(Request_CRUD_Operario oRq)
        {
            try
            {
                return new DOperario().DAnula(oRq);
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
        public Operario NAuditoria(Request_CRUD_Operario oRq)
        {
            try
            {
                return new DOperario().DAuditoria(oRq);
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
        public Operario NObjeto(Request_CRUD_Operario oRq)
        {
            try
            {
                return new DOperario().DObjeto(oRq);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}